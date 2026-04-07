#Requires -RunAsAdministrator
<#
.SYNOPSIS
    rs-tool v9 - Revit Server Backup Tool

    REMOTE mode : run from Windows 10/11 workstation with Revit installed
                  discovers server via RSN.ini scan or manual entry
                  uses REST API to fetch model tree (no admin shares needed)

    LOCAL mode  : run directly on the Windows Server 2016-2025 host
                  revitservertool.exe must be present (does NOT ship with
                  Revit Server - install Revit or copy the tool manually)
                  backup saved to configurable paath (no Desktop assumed)
#>

# ----------------------------------------------------------------
# CONFIG - adjust for your environment
# ----------------------------------------------------------------
# LOCAL mode backup destination. Leave empty to auto-resolve to
# current user Desktop (works for interactive sessions on Server).
# Set an explicit path for scheduled / headless runs:
#   e.g. "D:\RevitBackup"
$LocalBackupRoot = ""

# ----------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------
function Write-Title {
    param([string]$Text)
    $line = "-" * 64
    Write-Host ""
    Write-Host $line          -ForegroundColor DarkCyan
    Write-Host "  $Text"      -ForegroundColor Cyan
    Write-Host $line          -ForegroundColor DarkCyan
    Write-Host ""
}
function Write-OK   { param([string]$M) Write-Host "  [OK]  $M" -ForegroundColor Green      }
function Write-Info { param([string]$M) Write-Host "  [..]  $M" -ForegroundColor DarkGray   }
function Write-Warn { param([string]$M) Write-Host "  [!!]  $M" -ForegroundColor Yellow     }
function Write-Fail { param([string]$M) Write-Host "  [XX]  $M" -ForegroundColor Red        }
function Write-Skip { param([string]$M) Write-Host "  [--]  $M" -ForegroundColor DarkYellow }

# ----------------------------------------------------------------
# RESOLVE BACKUP ROOT
# Works for:
#   - interactive Desktop session on Windows Server
#   - headless / scheduled task on Windows Server (uses $LocalBackupRoot)
#   - Windows 10/11 workstation (standard Desktop)
# ----------------------------------------------------------------
function Get-BackupRoot {
    if (-not [string]::IsNullOrWhiteSpace($LocalBackupRoot)) {
        return $LocalBackupRoot
    }
    # Try current user Desktop
    $d = [Environment]::GetFolderPath("Desktop")
    if ([string]::IsNullOrWhiteSpace($d) -or -not (Test-Path (Split-Path $d -Parent) -ErrorAction SilentlyContinue)) {
        # Fallback: use USERPROFILE\Desktop
        $d = Join-Path $env:USERPROFILE "Desktop"
    }
    if ([string]::IsNullOrWhiteSpace($d)) {
        # Last resort: C:\RevitBackup (typical for SYSTEM/headless on Server)
        $d = "C:\RevitBackup"
    }
    return $d
}

# ----------------------------------------------------------------
# REST API HELPERS
# Revit Server REST API runs on HTTP (port 80), no TLS needed
# Required headers: User-Name, User-Machine-Name, Operation-GUID
# Path separator: | (pipe)   Root: |/contents
# ----------------------------------------------------------------
function New-RSNHeaders {
    return @{
        "User-Name"         = $env:USERNAME
        "User-Machine-Name" = $env:COMPUTERNAME
        "Operation-GUID"    = [guid]::NewGuid().ToString()
    }
}

function Invoke-RSNApi {
    param([string]$BaseUrl, [string]$ApiPath)
    $url = "$BaseUrl/$ApiPath"
    try {
        return Invoke-RestMethod -Uri $url -Headers (New-RSNHeaders) -Method GET -ErrorAction Stop
    } catch {
        return $null
    }
}

function Get-RSNModels {
    param([string]$BaseUrl, [string]$FolderRSNPath = "")
    $results = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrEmpty($FolderRSNPath)) {
        $apiPath = "|/contents"
    } else {
        $apiPath = "|$($FolderRSNPath.Replace('/',  '|'))/contents"
    }
    $resp = Invoke-RSNApi -BaseUrl $BaseUrl -ApiPath $apiPath
    if ($null -eq $resp) { return $results }
    if ($resp.Models) {
        foreach ($m in $resp.Models) {
            $results.Add($(if ([string]::IsNullOrEmpty($FolderRSNPath)) { $m.Name } else { "$FolderRSNPath/$($m.Name)" }))
        }
    }
    if ($resp.Folders) {
        foreach ($f in $resp.Folders) {
            $sub = $(if ([string]::IsNullOrEmpty($FolderRSNPath)) { $f.Name } else { "$FolderRSNPath/$($f.Name)" })
            foreach ($r in (Get-RSNModels -BaseUrl $BaseUrl -FolderRSNPath $sub)) { $results.Add($r) }
        }
    }
    return $results
}

# ----------------------------------------------------------------
# OS DETECTION - informs warnings shown to the user
# ----------------------------------------------------------------
$osCaption = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
$isWindowsServer = $osCaption -match "Server"

# ----------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host "   rs-tool v9  //  Revit Server Backup Tool" -ForegroundColor Cyan
Write-Host "   $($env:COMPUTERNAME)  //  $osCaption" -ForegroundColor DarkGray
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host ""

# ----------------------------------------------------------------
# STEP 1 - Mode selection
# ----------------------------------------------------------------
Write-Title "Step 1: Run Mode"

Write-Host "    [1]  LOCAL   run on the Revit Server host (Windows Server)" -ForegroundColor Cyan
Write-Host "    [2]  REMOTE  run from workstation, connect to server over network" -ForegroundColor Cyan
Write-Host ""

if ($isWindowsServer) {
    Write-Warn "Windows Server detected - LOCAL mode recommended on this machine."
    Write-Host ""
}

$modeInput = Read-Host "  Enter 1 or 2"
$isRemote  = $false
switch ($modeInput.Trim()) {
    "1" { $isRemote = $false; Write-OK "Mode: LOCAL" }
    "2" { $isRemote = $true;  Write-OK "Mode: REMOTE" }
    default { Write-Fail "Invalid choice. Exiting."; exit 1 }
}

# LOCAL mode on Windows Server: warn if revitservertool.exe is likely missing
if (-not $isRemote -and $isWindowsServer) {
    Write-Host ""
    Write-Warn "Note: revitservertool.exe does NOT ship with Revit Server."
    Write-Warn "It ships with Revit workstation. You need one of:"
    Write-Host "    a) Revit installed on this server (not recommended by Autodesk)" -ForegroundColor DarkGray
    Write-Host "    b) revitservertool.exe copied manually from a workstation" -ForegroundColor DarkGray
    Write-Host "    c) Run this script in REMOTE mode from a workstation instead" -ForegroundColor DarkGray
    Write-Host ""
    $cont = Read-Host "  Continue anyway? (Y/N)"
    if ($cont.Trim().ToUpper() -ne "Y") { exit 0 }
}

# ----------------------------------------------------------------
# STEP 2 - Server hostname
# ----------------------------------------------------------------
Write-Title "Step 2: Server Hostname"

$serverHost = $null

if (-not $isRemote) {
    $serverHost = $env:COMPUTERNAME
    Write-OK "Local machine: $serverHost"
} else {
    # Scan RSN.ini across all versions and all user profiles
    $rsnCandidates = [System.Collections.Generic.List[string]]::new()
    $userDirs = @("C:\Users\$env:USERNAME") + @(
        Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    )
    foreach ($v in 2020..2027) {
        $rsnCandidates.Add("C:\ProgramData\Autodesk\Revit Server $v\Config\RSN.ini")
        $rsnCandidates.Add("C:\ProgramData\Autodesk\Autodesk Revit Server $v\Config\RSN.ini")
        foreach ($u in $userDirs) {
            $rsnCandidates.Add("$u\AppData\Roaming\Autodesk\Revit\Autodesk Revit $v\RSN.ini")
            $rsnCandidates.Add("$u\AppData\Roaming\Autodesk\Revit Server $v\RSN.ini")
        }
    }

    Write-Host "  Scanning for RSN.ini files..." -ForegroundColor White
    Write-Host ""

    $foundServers = [System.Collections.Generic.List[PSCustomObject]]::new()
    $seenHosts    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($rsnPath in ($rsnCandidates | Select-Object -Unique)) {
        if (Test-Path $rsnPath -ErrorAction SilentlyContinue) {
            foreach ($line in (Get-Content $rsnPath -ErrorAction SilentlyContinue)) {
                $entry = $line.Trim()
                if (-not [string]::IsNullOrWhiteSpace($entry) -and $seenHosts.Add($entry)) {
                    $foundServers.Add([PSCustomObject]@{ Host = $entry; Source = $rsnPath })
                    Write-Host "  Found : $entry" -ForegroundColor Green
                    Write-Host "    from: $rsnPath" -ForegroundColor DarkGray
                    Write-Host ""
                }
            }
        }
    }

    if ($foundServers.Count -eq 0) {
        Write-Warn "No RSN.ini entries found on this machine."
        Write-Host ""
    }

    Write-Host "  Select server:" -ForegroundColor White
    Write-Host ""
    $menuIndex = 1
    foreach ($s in $foundServers) {
        Write-Host "    [$menuIndex]  $($s.Host)" -ForegroundColor Cyan
        $menuIndex++
    }
    Write-Host "    [$menuIndex]  Enter hostname / IP manually" -ForegroundColor Yellow
    Write-Host ""

    $selInt = 0
    if (-not [int]::TryParse((Read-Host "  Enter number").Trim(), [ref]$selInt)) {
        Write-Fail "Invalid input. Exiting."; exit 1
    }

    if ($selInt -ge 1 -and $selInt -le $foundServers.Count) {
        $serverHost = $foundServers[$selInt - 1].Host
        Write-OK "Selected: $serverHost"
    } elseif ($selInt -eq $menuIndex) {
        Write-Host ""
        Write-Host "  Hostname, IP, or FQDN  (e.g. REVIT-SRV / 10.0.0.50 / revit.company.com)" -ForegroundColor Yellow
        Write-Host ""
        $serverHost = (Read-Host "  Server").Trim()
        if ([string]::IsNullOrWhiteSpace($serverHost)) { Write-Fail "No input. Exiting."; exit 1 }
        Write-OK "Server: $serverHost"
    } else {
        Write-Fail "Invalid selection. Exiting."; exit 1
    }

    # Quick connectivity test - short timeout, ICMP often blocked on Server
    Write-Info "Pinging $serverHost ..."
    $ping = Test-Connection -ComputerName $serverHost -Count 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue |
            Select-Object -First 1
    if ($ping) { Write-OK "Host reachable." }
    else        { Write-Warn "No ping reply - continuing (ICMP may be blocked on Windows Server)." }
}

# ----------------------------------------------------------------
# STEP 3 - Find revitservertool.exe
# ----------------------------------------------------------------
Write-Title "Step 3: Locate revitservertool.exe"

$toolPathPatterns = @(
    "C:\Program Files\Autodesk\Revit {VER}\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Autodesk Revit {VER}\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Revit {VER}\tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Autodesk Revit {VER}\tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Revit Server {VER}\Tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files (x86)\Autodesk\Revit Server {VER}\Tools\RevitServerToolCommand\revitservertool.exe",
    "C:\RevitServerTools\{VER}\revitservertool.exe"
)

$detectedTools = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "  Scanning versions 2020-2027..." -ForegroundColor White
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
    Write-Warn "revitservertool.exe not found in any standard location."
    if ($isWindowsServer -and -not $isRemote) {
        Write-Host ""
        Write-Host "  On Windows Server, the tool is not installed by default." -ForegroundColor Yellow
        Write-Host "  Copy it from a Revit workstation to one of these paths:" -ForegroundColor Yellow
        Write-Host "    C:\RevitServerTools\2025\revitservertool.exe" -ForegroundColor DarkGray
        Write-Host "    C:\RevitServerTools\2026\revitservertool.exe" -ForegroundColor DarkGray
    }
    Write-Host ""
    $manualTool = Read-Host "  Full path to revitservertool.exe (or Enter to abort)"
    if ([string]::IsNullOrWhiteSpace($manualTool) -or -not (Test-Path $manualTool.Trim())) {
        Write-Fail "Not found. Cannot continue."
        exit 1
    }
    $manualVer = Read-Host "  Revit version for this tool (e.g. 2025)"
    $detectedTools.Add([PSCustomObject]@{
        Version = $manualVer.Trim()
        ToolExe = $manualTool.Trim()
        FileVer = (Get-Item $manualTool.Trim()).VersionInfo.FileVersion
    })
}

# ----------------------------------------------------------------
# STEP 4 - Select version
# ----------------------------------------------------------------
Write-Title "Step 4: Revit Version"

$selectedTool = $null
if ($detectedTools.Count -eq 1) {
    $selectedTool = $detectedTools[0]
    Write-OK "Auto-selected: Revit $($selectedTool.Version)"
} else {
    Write-Host "  Pick the version matching your Revit Server:" -ForegroundColor White
    Write-Host "  (tool version must match server version)" -ForegroundColor DarkGray
    Write-Host ""
    for ($i = 0; $i -lt $detectedTools.Count; $i++) {
        Write-Host "    [$($i+1)]  Revit $($detectedTools[$i].Version)  -  $($detectedTools[$i].ToolExe)" -ForegroundColor Cyan
    }
    Write-Host ""
    $idx = [int](Read-Host "  Enter number") - 1
    if ($idx -lt 0 -or $idx -ge $detectedTools.Count) { Write-Fail "Invalid selection. Exiting."; exit 1 }
    $selectedTool = $detectedTools[$idx]
    Write-OK "Selected: Revit $($selectedTool.Version)"
}

$version     = $selectedTool.Version
$toolExe     = $selectedTool.ToolExe
$toolFileVer = $selectedTool.FileVer
Write-OK "Tool    : $toolExe"
Write-OK "Version : $toolFileVer"

# ----------------------------------------------------------------
# STEP 5 - Discover models via REST API
# Revit Server REST API: HTTP port 80, no TLS, no shares
# ----------------------------------------------------------------
Write-Title "Step 5: Model Discovery (REST API)"

$apiBase  = "http://${serverHost}/RevitServerAdminRESTService${version}/AdminRESTService.svc"
$models   = $null
$apiUsed  = $false

Write-Info "Trying: $apiBase"
$testResp = Invoke-RSNApi -BaseUrl $apiBase -ApiPath "|/contents"

if ($null -eq $testResp) {
    Write-Warn "No response with version suffix - trying without..."
    $apiBase  = "http://${serverHost}/RevitServerAdminRESTService/AdminRESTService.svc"
    $testResp = Invoke-RSNApi -BaseUrl $apiBase -ApiPath "|/contents"
}

if ($null -ne $testResp) {
    Write-OK "REST API connected: $apiBase"
    Write-Host ""
    Write-Info "Crawling model tree..."
    $rsnPaths = Get-RSNModels -BaseUrl $apiBase -FolderRSNPath ""
    $models   = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($rsnPath in $rsnPaths) {
        $models.Add([PSCustomObject]@{ Name = [System.IO.Path]::GetFileName($rsnPath); RSNPath = $rsnPath })
    }
    Write-OK "Models found: $($models.Count)"
    $apiUsed = $true
} else {
    Write-Warn "REST API unreachable. Check:"
    Write-Host "    - Revit Server service running on $serverHost" -ForegroundColor DarkGray
    Write-Host "    - Port 80 open in Windows Firewall on $serverHost" -ForegroundColor DarkGray
    Write-Host "    - Version match (tool: $version)" -ForegroundColor DarkGray
    Write-Host ""

    if ($isWindowsServer -and -not $isRemote) {
        Write-Host "  On Windows Server: verify the Autodesk Revit Server service is running:" -ForegroundColor Yellow
        Write-Host "    Get-Service | Where-Object { `$_.DisplayName -like '*Revit*' }" -ForegroundColor DarkGray
        Write-Host ""
    }

    $cont = Read-Host "  Use filesystem scan of Projects folder instead? (Y/N)"
    if ($cont.Trim().ToUpper() -ne "Y") { exit 1 }

    Write-Host ""
    if ($isRemote) {
        Write-Host "  Enter UNC path to Projects folder on the server:" -ForegroundColor Yellow
        Write-Host "  e.g. \\$serverHost\C`$\ProgramData\Autodesk\Revit Server $version\Projects" -ForegroundColor DarkGray
    } else {
        Write-Host "  Enter local path to Projects folder:" -ForegroundColor Yellow
        Write-Host "  e.g. C:\ProgramData\Autodesk\Revit Server $version\Projects" -ForegroundColor DarkGray
    }
    Write-Host ""
    $manualProj = (Read-Host "  Projects path").Trim()
    if (-not (Test-Path $manualProj -ErrorAction SilentlyContinue)) {
        Write-Fail "Path not accessible. Exiting."; exit 1
    }
    $models = [System.Collections.Generic.List[PSCustomObject]]::new()
    Get-ChildItem -Path $manualProj -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.PSIsContainer -and $_.Name -like "*.rvt" } |
        ForEach-Object {
            $rsnPath = $_.FullName.Replace($manualProj, "").TrimStart("\").Replace("\", "/")
            $models.Add([PSCustomObject]@{ Name = $_.Name; RSNPath = $rsnPath })
        }
    Write-OK "Filesystem scan: found $($models.Count) model(s)"
}

$modelCount = [int]($models | Measure-Object).Count
Write-Host ""
Write-Host "  Model list:" -ForegroundColor White
foreach ($m in $models) {
    Write-Host "    RSN://$serverHost/$($m.RSNPath)" -ForegroundColor DarkCyan
}
if ($modelCount -eq 0) {
    Write-Warn "No models found."
    if ((Read-Host "  Continue anyway? (Y/N)").Trim().ToUpper() -ne "Y") { exit 0 }
}

# ----------------------------------------------------------------
# STEP 6 - Backup destination
# Windows Server safe: resolves Desktop or falls back to configured path
# ----------------------------------------------------------------
Write-Title "Step 6: Backup Destination"

$backupRoot = Get-BackupRoot
$stamp      = Get-Date -Format "yyyyMMdd_HHmm"
$backupDest = Join-Path $backupRoot "RevitServer_RVT_Backup\${stamp}_${version}_${serverHost}"

New-Item -ItemType Directory -Path $backupDest -Force | Out-Null
Write-OK "Destination:"
Write-Host "  $backupDest" -ForegroundColor White

# ----------------------------------------------------------------
# STEP 7 - Export
# ----------------------------------------------------------------
Write-Title "Step 7: Exporting Models"

Write-Host "  Server : $serverHost  (Revit Server $version)" -ForegroundColor White
Write-Host "  Mode   : $(if ($isRemote) { 'REMOTE' } else { 'LOCAL' })" -ForegroundColor White
Write-Host "  Tool   : $toolExe" -ForegroundColor DarkGray
Write-Host "  Locked or busy models are skipped automatically." -ForegroundColor DarkGray
Write-Host ""

$successList = [System.Collections.Generic.List[string]]::new()
$skipList    = [System.Collections.Generic.List[string]]::new()
$failList    = [System.Collections.Generic.List[string]]::new()
$current     = 0

foreach ($m in $models) {
    $current++
    Write-Host "  [$current/$modelCount] $($m.RSNPath)" -ForegroundColor White

    $destFilePath = Join-Path $backupDest $m.RSNPath.Replace("/", "\")
    $destFolder   = [System.IO.Path]::GetDirectoryName($destFilePath)
    if (-not (Test-Path $destFolder)) { New-Item -ItemType Directory -Path $destFolder -Force | Out-Null }

    Write-Info "-> $destFilePath"

    $toolArgs = "createLocalRVT `"$($m.RSNPath)`" -s $serverHost -d `"$destFilePath`" -o"
    try {
        $proc     = Start-Process -FilePath $toolExe -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru
        $exitCode = $proc.ExitCode
        switch ($exitCode) {
            0 {
                if (Test-Path $destFilePath) {
                    $sizeMB = [math]::Round((Get-Item $destFilePath).Length / 1MB, 1)
                    Write-OK "OK  $($m.Name)  ($sizeMB MB)"
                    $successList.Add("$($m.RSNPath)  [$sizeMB MB]")
                } else {
                    Write-Warn "Exit 0 but file missing: $destFilePath"
                    $failList.Add("$($m.RSNPath)  [exit 0 / file missing]")
                }
            }
            1 { Write-Skip "Busy    : $($m.Name)";   $skipList.Add("$($m.RSNPath)  [exit 1 - busy]") }
            5 { Write-Skip "Locked  : $($m.Name)";   $skipList.Add("$($m.RSNPath)  [exit 5 - locked]") }
            default {
                Write-Fail "Failed  : $($m.Name)  (exit $exitCode)"
                $failList.Add("$($m.RSNPath)  [exit $exitCode]")
            }
        }
    } catch {
        Write-Fail "Exception: $($_.Exception.Message)"
        $failList.Add("$($m.RSNPath)  [exception: $($_.Exception.Message)]")
    }
    Write-Host ""
}

# ----------------------------------------------------------------
# STEP 8 - Manifest
# ----------------------------------------------------------------
$successCount = [int]($successList | Measure-Object).Count
$skipCount    = [int]($skipList    | Measure-Object).Count
$failCount    = [int]($failList    | Measure-Object).Count

$manifestPath = Join-Path $backupDest "_BACKUP_MANIFEST.txt"
$mf = [System.Collections.Generic.List[string]]::new()
$mf.Add("rs-tool v9  //  Revit Server RVT Backup Manifest")
$mf.Add("=" * 64)
$mf.Add("Date          : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$mf.Add("Mode          : $(if ($isRemote) { 'REMOTE' } else { 'LOCAL' })")
$mf.Add("OS            : $osCaption")
$mf.Add("This machine  : $($env:COMPUTERNAME)")
$mf.Add("Server        : $serverHost")
$mf.Add("Revit version : $version")
$mf.Add("Tool          : $toolExe")
$mf.Add("Tool version  : $toolFileVer")
$mf.Add("Discovery     : $(if ($apiUsed) { "REST API - $apiBase" } else { 'Filesystem scan' })")
$mf.Add("Destination   : $backupDest")
$mf.Add("")
$mf.Add("Total models  : $modelCount")
$mf.Add("Succeeded     : $successCount")
$mf.Add("Skipped       : $skipCount  (locked/busy)")
$mf.Add("Failed        : $failCount")
$mf.Add("")
if ($successList.Count -gt 0) {
    $mf.Add("SUCCEEDED ($successCount):"); $mf.Add("-" * 64)
    foreach ($s in $successList) { $mf.Add("  [OK]  $s") }
    $mf.Add("")
}
if ($skipList.Count -gt 0) {
    $mf.Add("SKIPPED ($skipCount):"); $mf.Add("-" * 64)
    foreach ($s in $skipList) { $mf.Add("  [--]  $s") }
    $mf.Add("")
    $mf.Add("  Tip: run after hours when all users are disconnected.")
    $mf.Add("")
}
if ($failList.Count -gt 0) {
    $mf.Add("FAILED ($failCount):"); $mf.Add("-" * 64)
    foreach ($f in $failList) { $mf.Add("  [XX]  $f") }
    $mf.Add("")
}
$mf | Out-File -FilePath $manifestPath -Encoding UTF8
Write-OK "Manifest: $manifestPath"

# ----------------------------------------------------------------
# DONE
# ----------------------------------------------------------------
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host "   Done" -ForegroundColor Green
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "  Server    : $serverHost  (Revit Server $version)" -ForegroundColor White
Write-Host "  Succeeded : $successCount" -ForegroundColor Green
Write-Host "  Skipped   : $skipCount$(if ($skipCount -gt 0) { '  (locked/busy)' })" -ForegroundColor $(if ($skipCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "  Failed    : $failCount" -ForegroundColor $(if ($failCount -gt 0) { 'Red' } else { 'Green' })
Write-Host ""
Write-Host "  $backupDest" -ForegroundColor Cyan
Write-Host ""

if (-not $isWindowsServer) {
    $open = Read-Host "  Open folder in Explorer? (Y/N)"
    if ($open.Trim().ToUpper() -eq "Y") { Start-Process explorer.exe $backupDest }
}

Write-Host ""
Write-Host "  Press Enter to exit." -ForegroundColor DarkGray
Read-Host | Out-Null
