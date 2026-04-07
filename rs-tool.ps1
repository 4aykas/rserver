#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Revit Server RVT Backup Script v8
    LOCAL  - run on the Revit Server host machine
    REMOTE - connect to Revit Server REST API (port 808 / HTTP)
             NO admin shares required. Model tree fetched via REST API.
.NOTES
    Revit Server REST API reference:
      http://<host>/RevitServerAdminRESTService<VER>/AdminRESTService.svc/
    Paths use | as separator: |FolderA|SubFolder|Model.rvt
    Required headers: User-Name, User-Machine-Name, Operation-GUID
#>

function Write-Title {
    param([string]$Text)
    $line = "=" * 64
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ""
}
function Write-OK   { param([string]$M) Write-Host "  [OK]  $M" -ForegroundColor Green      }
function Write-Info { param([string]$M) Write-Host "  [..] $M"  -ForegroundColor DarkGray   }
function Write-Warn { param([string]$M) Write-Host "  [!!] $M"  -ForegroundColor Yellow     }
function Write-Fail { param([string]$M) Write-Host "  [XX] $M"  -ForegroundColor Red        }
function Write-Skip { param([string]$M) Write-Host "  [--] $M"  -ForegroundColor DarkYellow }

# ----------------------------------------------------------------
# REST API helpers
# ----------------------------------------------------------------
function New-RSNHeaders {
    return @{
        "User-Name"          = $env:USERNAME
        "User-Machine-Name"  = $env:COMPUTERNAME
        "Operation-GUID"     = [guid]::NewGuid().ToString()
    }
}

function Invoke-RSNApi {
    param(
        [string]$BaseUrl,
        [string]$ApiPath     # e.g. "|/contents" or "|FolderA|SubFolder/contents"
    )
    $url = "$BaseUrl/$ApiPath"
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers (New-RSNHeaders) -Method GET -ErrorAction Stop
        return $resp
    } catch {
        return $null
    }
}

# Recursively walk the RSN folder tree via REST API
# Returns a flat list of model RSN paths like: FolderA/SubFolder/Model.rvt
function Get-RSNModels {
    param(
        [string]$BaseUrl,
        [string]$FolderRSNPath = ""   # empty = root
    )

    $results = [System.Collections.Generic.List[string]]::new()

    # Build the API path: root = "|" + "/contents", subfolder = "|Folder|Sub" + "/contents"
    if ([string]::IsNullOrEmpty($FolderRSNPath)) {
        $apiPath = "|/contents"
    } else {
        $pipeEncoded = $FolderRSNPath.Replace("/", "|")
        $apiPath = "|$pipeEncoded/contents"
    }

    $resp = Invoke-RSNApi -BaseUrl $BaseUrl -ApiPath $apiPath

    if ($null -eq $resp) { return $results }

    # Process models in this folder
    if ($resp.Models) {
        foreach ($m in $resp.Models) {
            if ([string]::IsNullOrEmpty($FolderRSNPath)) {
                $results.Add($m.Name)
            } else {
                $results.Add("$FolderRSNPath/$($m.Name)")
            }
        }
    }

    # Recurse into subfolders
    if ($resp.Folders) {
        foreach ($f in $resp.Folders) {
            if ([string]::IsNullOrEmpty($FolderRSNPath)) {
                $subPath = $f.Name
            } else {
                $subPath = "$FolderRSNPath/$($f.Name)"
            }
            $subResults = Get-RSNModels -BaseUrl $BaseUrl -FolderRSNPath $subPath
            foreach ($r in $subResults) { $results.Add($r) }
        }
    }

    return $results
}

# ----------------------------------------------------------------
# STEP 1 - Choose run mode
# ----------------------------------------------------------------
Write-Title "Revit Server RVT Backup v8 - $($env:COMPUTERNAME)"

Write-Host "  Run mode:" -ForegroundColor White
Write-Host "    [1]  LOCAL  - this machine IS the Revit Server host" -ForegroundColor Cyan
Write-Host "    [2]  REMOTE - connect to a Revit Server over the network" -ForegroundColor Cyan
Write-Host ""
$modeInput = Read-Host "  Enter 1 or 2"

$isRemote = $false
switch ($modeInput.Trim()) {
    "1" { $isRemote = $false; Write-OK "Mode: LOCAL" }
    "2" { $isRemote = $true;  Write-OK "Mode: REMOTE" }
    default { Write-Fail "Invalid choice. Exiting."; exit 1 }
}

# ----------------------------------------------------------------
# STEP 2 - Server hostname
#   LOCAL  -> use this machine name
#   REMOTE -> scan RSN.ini files + offer manual entry
# ----------------------------------------------------------------
Write-Title "Step 2: Server Hostname"

$serverHost = $null

if (-not $isRemote) {
    $serverHost = $env:COMPUTERNAME
    Write-OK "Local machine: $serverHost"
} else {
    $rsnCandidates = [System.Collections.Generic.List[string]]::new()
    $userDirs = @("C:\Users\$env:USERNAME") + @(Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)

    foreach ($v in 2020..2027) {
        $rsnCandidates.Add("C:\ProgramData\Autodesk\Revit Server $v\Config\RSN.ini")
        $rsnCandidates.Add("C:\ProgramData\Autodesk\Autodesk Revit Server $v\Config\RSN.ini")
        foreach ($u in $userDirs) {
            $rsnCandidates.Add("$u\AppData\Roaming\Autodesk\Revit\Autodesk Revit $v\RSN.ini")
            $rsnCandidates.Add("$u\AppData\Roaming\Autodesk\Revit Server $v\RSN.ini")
        }
    }

    Write-Host "  Scanning for RSN.ini files on this machine..." -ForegroundColor White
    Write-Host ""

    $foundServers = [System.Collections.Generic.List[PSCustomObject]]::new()
    $seenHosts    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($rsnPath in ($rsnCandidates | Select-Object -Unique)) {
        if (Test-Path $rsnPath -ErrorAction SilentlyContinue) {
            $lines = Get-Content $rsnPath -ErrorAction SilentlyContinue
            foreach ($line in $lines) {
                $entry = $line.Trim()
                if (-not [string]::IsNullOrWhiteSpace($entry) -and $seenHosts.Add($entry)) {
                    $foundServers.Add([PSCustomObject]@{ Host = $entry; Source = $rsnPath })
                    Write-Host "  Found server : $entry" -ForegroundColor Green
                    Write-Host "  RSN.ini      : $rsnPath" -ForegroundColor DarkGray
                    Write-Host ""
                }
            }
        }
    }

    if ($foundServers.Count -eq 0) {
        Write-Warn "No RSN.ini files found or all files are empty on this machine."
        Write-Host ""
    }

    Write-Host "  Select the Revit Server to connect to:" -ForegroundColor White
    Write-Host ""

    $menuIndex = 1
    foreach ($s in $foundServers) {
        Write-Host "    [$menuIndex]  $($s.Host)  (from RSN.ini)" -ForegroundColor Cyan
        $menuIndex++
    }
    Write-Host "    [$menuIndex]  Enter custom hostname / IP manually" -ForegroundColor Yellow
    Write-Host ""

    $sel = Read-Host "  Enter number"
    $selInt = 0
    if (-not [int]::TryParse($sel.Trim(), [ref]$selInt)) {
        Write-Fail "Invalid input. Exiting."
        exit 1
    }

    if ($selInt -ge 1 -and $selInt -le $foundServers.Count) {
        $serverHost = $foundServers[$selInt - 1].Host
        Write-OK "Selected: $serverHost"
    } elseif ($selInt -eq $menuIndex) {
        Write-Host ""
        Write-Host "  Enter hostname, IP, or FQDN of the Revit Server machine." -ForegroundColor Yellow
        Write-Host "  Examples: MYSERVER  /  192.168.1.50  /  revit.company.com" -ForegroundColor DarkGray
        Write-Host ""
        $serverHost = (Read-Host "  Hostname / IP").Trim()
        if ([string]::IsNullOrWhiteSpace($serverHost)) {
            Write-Fail "No hostname entered. Exiting."
            exit 1
        }
        Write-OK "Custom server: $serverHost"
    } else {
        Write-Fail "Invalid selection. Exiting."
        exit 1
    }

    Write-Info "Testing connectivity to $serverHost ..."
    $ping = Test-Connection -ComputerName $serverHost -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($ping) {
        Write-OK "Host is reachable (ping)."
    } else {
        Write-Warn "No ping response - continuing (ping may be blocked by firewall)."
    }
}

# ----------------------------------------------------------------
# STEP 3 - Locate revitservertool.exe on THIS machine
# ----------------------------------------------------------------
Write-Title "Step 3: Locating RevitServerTool on This Machine"

$toolPathPatterns = @(
    "C:\Program Files\Autodesk\Revit {VER}\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Autodesk Revit {VER}\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Revit {VER}\tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Autodesk Revit {VER}\tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Revit Server {VER}\Tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files (x86)\Autodesk\Revit Server {VER}\Tools\RevitServerToolCommand\revitservertool.exe"
)

$detectedTools = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "  Scanning for revitservertool.exe (versions 2020-2027)..." -ForegroundColor White
Write-Host ""

foreach ($v in 2020..2027) {
    foreach ($pat in $toolPathPatterns) {
        $p = $pat.Replace("{VER}", "$v")
        if (Test-Path $p) {
            $fv = (Get-Item $p).VersionInfo.FileVersion
            $detectedTools.Add([PSCustomObject]@{ Version = "$v"; ToolExe = $p; FileVer = $fv })
            Write-Host "  Found Revit $v : $p" -ForegroundColor Green
            break
        }
    }
}

if ($detectedTools.Count -eq 0) {
    Write-Warn "revitservertool.exe not found automatically."
    Write-Host ""
    foreach ($pat in $toolPathPatterns) {
        Write-Host "    $($pat.Replace('{VER}','XXXX'))" -ForegroundColor DarkGray
    }
    Write-Host ""
    $manualTool = Read-Host "  Full path to revitservertool.exe (or Enter to abort)"
    if ([string]::IsNullOrWhiteSpace($manualTool) -or -not (Test-Path $manualTool.Trim())) {
        Write-Fail "Tool not found. Cannot continue."
        exit 1
    }
    $manualVer = Read-Host "  Which Revit version does this tool belong to? (e.g. 2025)"
    $detectedTools.Add([PSCustomObject]@{
        Version = $manualVer.Trim()
        ToolExe = $manualTool.Trim()
        FileVer = (Get-Item $manualTool.Trim()).VersionInfo.FileVersion
    })
}

# ----------------------------------------------------------------
# STEP 4 - Select version
# ----------------------------------------------------------------
Write-Title "Step 4: Select Revit Version for Backup"

$selectedTool = $null

if ($detectedTools.Count -eq 1) {
    $selectedTool = $detectedTools[0]
    Write-OK "Auto-selected: Revit $($selectedTool.Version)"
} else {
    Write-Host "  Multiple versions found. Pick the one matching your Revit Server:" -ForegroundColor White
    Write-Host "  (Tool version must match the Revit Server version)" -ForegroundColor DarkGray
    Write-Host ""
    for ($i = 0; $i -lt $detectedTools.Count; $i++) {
        Write-Host "    [$($i+1)]  Revit $($detectedTools[$i].Version)  -  $($detectedTools[$i].ToolExe)" -ForegroundColor Cyan
    }
    Write-Host ""
    $idx = [int](Read-Host "  Enter number") - 1
    if ($idx -lt 0 -or $idx -ge $detectedTools.Count) {
        Write-Fail "Invalid selection. Exiting."
        exit 1
    }
    $selectedTool = $detectedTools[$idx]
    Write-OK "Selected: Revit $($selectedTool.Version)"
}

$version     = $selectedTool.Version
$toolExe     = $selectedTool.ToolExe
$toolFileVer = $selectedTool.FileVer

Write-OK "Tool    : $toolExe"
Write-OK "Version : $toolFileVer"

# ----------------------------------------------------------------
# STEP 5 - Connect to REST API and discover model tree
#   No admin shares needed - uses Revit Server HTTP REST API
#   Endpoint: http://<host>/RevitServerAdminRESTService<VER>/AdminRESTService.svc/
#   Folder path separator: | (pipe), root = "|"
# ----------------------------------------------------------------
Write-Title "Step 5: Discovering Models via REST API"

$apiBase = "http://${serverHost}/RevitServerAdminRESTService${version}/AdminRESTService.svc"

Write-Host "  REST API endpoint:" -ForegroundColor White
Write-Host "  $apiBase" -ForegroundColor DarkCyan
Write-Host ""

# Test API connectivity
Write-Info "Testing REST API connection..."
$testResp = Invoke-RSNApi -BaseUrl $apiBase -ApiPath "|/contents"

if ($null -eq $testResp) {
    Write-Warn "API at /RevitServerAdminRESTService${version} did not respond."
    Write-Host ""
    Write-Host "  Trying fallback without version suffix..." -ForegroundColor Yellow
    $apiBase  = "http://${serverHost}/RevitServerAdminRESTService/AdminRESTService.svc"
    $testResp = Invoke-RSNApi -BaseUrl $apiBase -ApiPath "|/contents"
}

if ($null -eq $testResp) {
    Write-Warn "REST API not reachable. Possible causes:"
    Write-Host "  - Revit Server service not running on $serverHost" -ForegroundColor Yellow
    Write-Host "  - Firewall blocking port 80 on $serverHost" -ForegroundColor Yellow
    Write-Host "  - Version mismatch (selected $version but server may be different)" -ForegroundColor Yellow
    Write-Host ""
    $cont = Read-Host "  Manually enter the Projects folder path instead? (Y/N)"
    if ($cont.Trim().ToUpper() -ne "Y") { exit 1 }

    Write-Host ""
    Write-Host "  Enter the local Projects folder path on the SERVER machine." -ForegroundColor Yellow
    if ($isRemote) {
        Write-Host "  (UNC example: \\$serverHost\C`$\ProgramData\Autodesk\Revit Server $version\Projects)" -ForegroundColor DarkGray
    } else {
        Write-Host "  (Local example: C:\ProgramData\Autodesk\Revit Server $version\Projects)" -ForegroundColor DarkGray
    }
    Write-Host ""
    $manualProj = Read-Host "  Projects folder path"
    if (-not (Test-Path $manualProj.Trim() -ErrorAction SilentlyContinue)) {
        Write-Fail "Path not accessible. Exiting."
        exit 1
    }

    # Fall back to filesystem scan
    $projRoot     = $manualProj.Trim()
    $allItems     = Get-ChildItem -Path $projRoot -Recurse -ErrorAction SilentlyContinue
    $modelFolders = @($allItems | Where-Object { $_.PSIsContainer -and $_.Name -like "*.rvt" })
    $models       = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($f in $modelFolders) {
        $relPath = $f.FullName.Replace($projRoot, "").TrimStart("\")
        $rsnPath = $relPath.Replace("\", "/")
        $models.Add([PSCustomObject]@{ Name = $f.Name; RSNPath = $rsnPath })
    }
    Write-OK "Filesystem fallback: found $($models.Count) model(s)"
} else {
    Write-OK "REST API connected successfully."
    Write-Host ""
    Write-Host "  Crawling model tree..." -ForegroundColor White

    $rsnPaths = Get-RSNModels -BaseUrl $apiBase -FolderRSNPath ""
    $models   = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($rsnPath in $rsnPaths) {
        $name = [System.IO.Path]::GetFileName($rsnPath)
        $models.Add([PSCustomObject]@{ Name = $name; RSNPath = $rsnPath })
    }

    Write-OK "Models found via REST API: $($models.Count)"
}

$modelCount = [int]($models | Measure-Object).Count

Write-Host ""
Write-Host "  Model list:" -ForegroundColor White
foreach ($m in $models) {
    Write-Host "    RSN://$serverHost/$($m.RSNPath)" -ForegroundColor DarkCyan
}

if ($modelCount -eq 0) {
    Write-Warn "No models found."
    $cont = Read-Host "  Continue anyway? (Y/N)"
    if ($cont.Trim().ToUpper() -ne "Y") { exit 0 }
}

# ----------------------------------------------------------------
# STEP 6 - Create backup folder on Desktop
# ----------------------------------------------------------------
Write-Title "Step 6: Creating Backup Folder"

$desktopPath = [Environment]::GetFolderPath("Desktop")
$stamp       = Get-Date -Format "yyyyMMdd_HHmm"
$backupDest  = Join-Path $desktopPath "RevitServer_RVT_Backup\${stamp}_${version}_${serverHost}"

New-Item -ItemType Directory -Path $backupDest -Force | Out-Null
Write-OK "Backup folder:"
Write-Host "  $backupDest" -ForegroundColor White

# ----------------------------------------------------------------
# STEP 7 - Export RVT files via revitservertool.exe
# ----------------------------------------------------------------
Write-Title "Step 7: Exporting RVT Files"

Write-Host "  Tool   : revitservertool.exe createLocalRVT" -ForegroundColor White
Write-Host "  Server : $serverHost  (Revit Server $version)" -ForegroundColor White
if ($isRemote) { Write-Host "  Mode   : REMOTE" -ForegroundColor White } else { Write-Host "  Mode   : LOCAL" -ForegroundColor White }
Write-Host "  Locked/busy models will be SKIPPED automatically." -ForegroundColor DarkGray
Write-Host ""

$successList = [System.Collections.Generic.List[string]]::new()
$skipList    = [System.Collections.Generic.List[string]]::new()
$failList    = [System.Collections.Generic.List[string]]::new()
$current     = 0

foreach ($m in $models) {
    $current++
    Write-Host "  [$current/$modelCount] $($m.RSNPath)" -ForegroundColor White

    # Mirror folder structure in backup destination
    $relPath      = $m.RSNPath.Replace("/", "\")
    $destFilePath = Join-Path $backupDest $relPath
    $destFolder   = [System.IO.Path]::GetDirectoryName($destFilePath)

    if (-not (Test-Path $destFolder)) {
        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
    }

    Write-Info "  -> $destFilePath"

    # revitservertool.exe expects RSN path with forward slashes, no leading slash
    $toolArgs = "createLocalRVT `"$($m.RSNPath)`" -s $serverHost -d `"$destFilePath`" -o"

    try {
        $proc     = Start-Process -FilePath $toolExe -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru
        $exitCode = $proc.ExitCode

        if ($exitCode -eq 0) {
            if (Test-Path $destFilePath) {
                $sizeMB = [math]::Round((Get-Item $destFilePath).Length / 1MB, 1)
                Write-OK "Exported: $($m.Name)  ($sizeMB MB)"
                $successList.Add("$($m.RSNPath)  [$sizeMB MB]")
            } else {
                Write-Warn "Exit 0 but file not created: $destFilePath"
                $failList.Add("$($m.RSNPath)  [exit 0 but file missing]")
            }
        } elseif ($exitCode -eq 5) {
            Write-Skip "SKIPPED - locked by user: $($m.Name)"
            $skipList.Add("$($m.RSNPath)  [exit 5 - locked]")
        } elseif ($exitCode -eq 1) {
            Write-Skip "SKIPPED - model busy: $($m.Name)"
            $skipList.Add("$($m.RSNPath)  [exit 1 - busy]")
        } else {
            Write-Fail "FAILED (exit $exitCode): $($m.Name)"
            $failList.Add("$($m.RSNPath)  [exit $exitCode]")
        }
    } catch {
        Write-Fail "Exception: $($_.Exception.Message)"
        $failList.Add("$($m.RSNPath)  [exception: $($_.Exception.Message)]")
    }
    Write-Host ""
}

# ----------------------------------------------------------------
# STEP 8 - Write manifest
# ----------------------------------------------------------------
$successCount = [int]($successList | Measure-Object).Count
$skipCount    = [int]($skipList    | Measure-Object).Count
$failCount    = [int]($failList    | Measure-Object).Count

$manifestPath = Join-Path $backupDest "_BACKUP_MANIFEST.txt"
$manifest     = [System.Collections.Generic.List[string]]::new()
$manifest.Add("Revit Server RVT Backup Manifest v8")
$manifest.Add("=" * 64)
$manifest.Add("Date          : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
if ($isRemote) { $manifest.Add("Run Mode      : REMOTE (REST API)") } else { $manifest.Add("Run Mode      : LOCAL (REST API)") }
$manifest.Add("REST API Base : $apiBase")
$manifest.Add("Revit Version : $version")
$manifest.Add("Server Host   : $serverHost")
$manifest.Add("This Machine  : $($env:COMPUTERNAME)")
$manifest.Add("Tool          : $toolExe")
$manifest.Add("Tool Version  : $toolFileVer")
$manifest.Add("Destination   : $backupDest")
$manifest.Add("")
$manifest.Add("Models total  : $modelCount")
$manifest.Add("Succeeded     : $successCount")
$manifest.Add("Skipped       : $skipCount  (locked/busy - not an error)")
$manifest.Add("Failed        : $failCount")
$manifest.Add("")
if ($successList.Count -gt 0) {
    $manifest.Add("SUCCEEDED ($successCount):")
    $manifest.Add("-" * 64)
    foreach ($s in $successList) { $manifest.Add("  [OK]  $s") }
    $manifest.Add("")
}
if ($skipList.Count -gt 0) {
    $manifest.Add("SKIPPED - locked or busy ($skipCount):")
    $manifest.Add("-" * 64)
    foreach ($s in $skipList) { $manifest.Add("  [--] $s") }
    $manifest.Add("")
    $manifest.Add("  Tip: run backup after hours when all users are disconnected.")
    $manifest.Add("")
}
if ($failList.Count -gt 0) {
    $manifest.Add("FAILED ($failCount):")
    $manifest.Add("-" * 64)
    foreach ($f in $failList) { $manifest.Add("  [XX] $f") }
    $manifest.Add("")
}

$manifest | Out-File -FilePath $manifestPath -Encoding UTF8
Write-OK "Manifest: $manifestPath"

# ----------------------------------------------------------------
# DONE
# ----------------------------------------------------------------
Write-Title "Backup Complete"

Write-Host "  Server        : $serverHost  (Revit Server $version)" -ForegroundColor White
Write-Host "  Total models  : $modelCount"   -ForegroundColor White
Write-Host "  Succeeded     : $successCount" -ForegroundColor Green
if ($skipCount -gt 0) { Write-Host "  Skipped       : $skipCount" -ForegroundColor Yellow } else { Write-Host "  Skipped       : $skipCount" -ForegroundColor Green }
if ($failCount -gt 0) { Write-Host "  Failed        : $failCount" -ForegroundColor Red    } else { Write-Host "  Failed        : $failCount" -ForegroundColor Green }
Write-Host ""
Write-Host "  Backup   : $backupDest"   -ForegroundColor Cyan
Write-Host "  Manifest : $manifestPath" -ForegroundColor Cyan
Write-Host ""

$open = Read-Host "  Open backup folder in Explorer? (Y/N)"
if ($open.Trim().ToUpper() -eq "Y") { Start-Process explorer.exe $backupDest }

Write-Host ""
Write-Host "  Done. Press Enter to exit." -ForegroundColor Green
Read-Host | Out-Null
