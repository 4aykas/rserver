#Requires -RunAsAdministrator
<#
.SYNOPSIS
    rs-host v1 - Revit Server host names + RSN.ini writer

.DESCRIPTION
    Makes Revit Servers addressable by a readable name instead of an
    IP address, and registers those names for the right Revit versions:

      1. Writes friendly names into the Windows hosts file
         (C:\Windows\System32\drivers\etc\hosts) inside a managed block,
         so re-runs update rather than duplicate.
      2. Writes each server into RSN.ini for the Revit versions it serves
         (C:\ProgramData\Autodesk\Revit\Autodesk Revit <year>\RSN.ini),
         which is the list Revit shows in "Open > Revit Server".

    You choose what goes where:

      -Target    Rsn  (default) | Hosts | Both
                 which file(s) to touch at all. The default is the
                 minimal one: RSN.ini only, no hosts entry needed.
      -RsnEntry  Ip  (default) | Name | Both
                 what each RSN.ini line contains - the raw IP (works
                 with no hosts entry at all), the readable name (needs
                 the hosts entry, or real DNS), or both lines, so
                 Revit's list shows the name and the IP as a fallback.

    Run without them interactively and the script asks - both for the
    files to write and for which servers of the table to use.

    RSN.ini is never rewritten: an existing file keeps its content and
    the missing servers are appended on new lines. A missing file is
    created. Only -Remove rewrites, and only to drop our own lines.

    One IP may carry several names (e.g. the same machine hosting two
    Revit Server versions) - that is expected and supported: list one
    entry per name.

    The server list lives in the SERVERS table below - edit it there.
    Both files are backed up before the first change of a run, and
    -WhatIf previews every change without touching anything.

.PARAMETER Only
    Process only these entries (match by name or by IP). Default: all,
    or - interactively - whatever is picked at the prompt.

.PARAMETER Target
    Which files to write: Rsn (default), Hosts, or Both.

.PARAMETER RsnEntry
    What to put in RSN.ini: Ip (default), Name, or Both.

.PARAMETER Remove
    Revert: delete the managed hosts block and remove the managed names
    from every RSN.ini. Other entries in those files are left alone.

.PARAMETER Force
    Comment out conflicting hosts entries (same name, different IP,
    outside the managed block) without asking.

.PARAMETER SkipVerify
    Do not resolve the names or probe the Revit Server admin page.

.PARAMETER NonInteractive
    Never prompt and never pause at the end.

.EXAMPLE
    .\rs-host.ps1
    Interactive: asks what to write and for which servers, then verifies.
    Accepting both defaults appends the IPs to RSN.ini only.

.EXAMPLE
    .\rs-host.ps1 -WhatIf
    Dry run - shows every line that would be written.

.EXAMPLE
    .\rs-host.ps1 -Only AVRORA
    Only the AVRORA entry (and only the Revit versions it declares).

.EXAMPLE
    .\rs-host.ps1 -NonInteractive
    Unattended default: RSN.ini only, plain IP addresses, hosts untouched.

.EXAMPLE
    .\rs-host.ps1 -Target Both -RsnEntry Both
    hosts + RSN.ini, and RSN.ini lists both the name and the IP.

.EXAMPLE
    .\rs-host.ps1 -Remove -NonInteractive
    Unattended cleanup of everything this script manages.

.NOTES
    Run on the WORKSTATION (the machine that opens the models), elevated.
    Exit codes: 0 = success, 1 = fatal / bad input,
                2 = written, but verification failed for some server.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]]$Only,

    # No ValidateSet: under `irm | iex` validation attributes fire on the
    # default value before the script runs. Checked by hand below.
    [string]$Target = "",

    [string]$RsnEntry = "",

    [switch]$Remove,

    [switch]$Force,

    [switch]$SkipVerify,

    [switch]$NonInteractive
)

# ----------------------------------------------------------------
# SERVERS - the only table you normally edit
#
#   Ip     : address of the Revit Server host
#   Name   : the readable name typed into Revit ("Revit Server Name")
#   Revit  : Revit versions this name serves - one RSN.ini per version
#   Note   : free text, shown in the plan only
#
# Names must be valid host names: letters, digits, hyphen. No spaces,
# no underscores. Several entries may share one Ip.
# ----------------------------------------------------------------
$SERVERS = @(
    @{ Ip = "10.253.9.2";   Name = "AVRORA";     Revit = @(2023); Note = "Avrora" }
    @{ Ip = "10.253.11.26"; Name = "TEBIN-lab";  Revit = @(2025); Note = "TEBIN lab" }
    @{ Ip = "10.253.9.34";  Name = "CERSANIT";   Revit = @(2025); Note = "Cersanit" }
    # Same machine as TEBIN-lab, second Revit Server version - own name on purpose.
    @{ Ip = "10.253.11.26"; Name = "DRAGON";     Revit = @(2026); Note = "Dragon" }
)

$script:Interactive = -not $NonInteractive

# Validated by hand (see the param block); empty means "ask, or default".
$TARGET_CHOICES = @("Both", "Hosts", "Rsn")
$RSN_CHOICES    = @("Name", "Ip", "Both")

$HOSTS_PATH   = Join-Path $env:SystemRoot "System32\drivers\etc\hosts"
$BLOCK_BEGIN  = "# >>> rs-host: Revit Server names (managed - do not edit by hand) >>>"
$BLOCK_END    = "# <<< rs-host <<<"

# #Requires -RunAsAdministrator only guards file runs; via `irm | iex`
# it is silently ignored, so check explicitly.
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "  [XX]  Run this in an elevated (Administrator) PowerShell." -ForegroundColor Red
    exit 1
}

# ----------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------
function Write-Title {
    param([string]$Text)
    Write-Host ""
    Write-Host ("  " + "-" * 62) -ForegroundColor DarkCyan
    Write-Host "  $Text"          -ForegroundColor Cyan
    Write-Host ("  " + "-" * 62) -ForegroundColor DarkCyan
    Write-Host ""
}
function Write-OK   { param([string]$M) Write-Host "  [OK]  $M" -ForegroundColor Green      }
function Write-Info { param([string]$M) Write-Host "  [..]  $M" -ForegroundColor DarkGray   }
function Write-Warn { param([string]$M) Write-Host "  [!!]  $M" -ForegroundColor Yellow     }
function Write-Fail { param([string]$M) Write-Host "  [XX]  $M" -ForegroundColor Red        }

function Test-Approved {
    # ShouldProcess gate that also works under `irm | iex`. There the script
    # body is not an advanced function, so $PSCmdlet is $null and every
    # `$PSCmdlet.ShouldProcess(...)` threw InvokeMethodOnNull - a non-
    # terminating error, so the run continued and silently wrote nothing.
    # -Caller is the script's own $PSCmdlet ($null under iex).
    [CmdletBinding(SupportsShouldProcess)]
    param([object]$Caller, [string]$Target, [string]$Action)
    if ($null -ne $Caller) { return $Caller.ShouldProcess($Target, $Action) }
    if ($WhatIfPreference) {
        Write-Host "  What if: $Action -> $Target" -ForegroundColor DarkGray
        return $false
    }
    return $true
}

function Backup-Once {
    # One backup per file per run, kept next to the original.
    param([string]$Path)
    if (-not $script:BackedUp) { $script:BackedUp = @{} }
    if ($script:BackedUp.ContainsKey($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $stamp  = Get-Date -Format "yyyyMMdd-HHmmss"
    $backup = "$Path.rs-host-$stamp.bak"
    if (Test-Approved -Caller $PSCmdlet -Target $backup -Action "Create backup") {
        Copy-Item -LiteralPath $Path -Destination $backup -Force -ErrorAction Stop
        Write-Info "Backup: $backup"
    }
    $script:BackedUp[$Path] = $true
}

function Write-TextFile {
    # ASCII, CRLF, no BOM - what both hosts and RSN.ini expect.
    param([string]$Path, [string[]]$Lines)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, (($Lines -join "`r`n") + "`r`n"), [System.Text.Encoding]::ASCII)
}

function Add-TextLine {
    # Append only - the existing file keeps its bytes. Creates the file
    # (and its folder) when missing. Adds the missing newline first if the
    # file does not end with one, so we never glue onto the last entry.
    param([string]$Path, [string[]]$Lines)
    if (-not $Lines -or $Lines.Count -eq 0) { return }
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
    }
    $prefix = ""
    if (Test-Path -LiteralPath $Path) {
        $current = [System.IO.File]::ReadAllText($Path)
        if ($current.Length -gt 0 -and $current[$current.Length - 1] -ne "`n") { $prefix = "`r`n" }
    }
    [System.IO.File]::AppendAllText($Path, ($prefix + ($Lines -join "`r`n") + "`r`n"), [System.Text.Encoding]::ASCII)
}

function Get-FileLine {
    # The leading comma matters: without it PowerShell unrolls a one-element
    # array on return, so a single-line file came back as a bare string and
    # $lines[0] indexed its first CHARACTER instead of the line.
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return ,@() }
    $content = Get-Content -LiteralPath $Path -ErrorAction Stop
    if ($null -eq $content) { return ,@() }
    return ,@($content)
}

function Test-HostNameValid {
    param([string]$Name)
    return ($Name -match '^[A-Za-z0-9]([A-Za-z0-9\-\.]*[A-Za-z0-9])?$')
}

function Get-RsnPath {
    # The client-side list Revit reads. Revit Server's own config lives
    # elsewhere and is deliberately not touched.
    param([int]$Version)
    return "C:\ProgramData\Autodesk\Revit\Autodesk Revit $Version\RSN.ini"
}

# ----------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------
if ($script:Interactive) { Clear-Host }
$mode = if ($Remove) { "REMOVE" } else { "APPLY" }
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host "   rs-host v1  //  Revit Server names + RSN.ini"                     -ForegroundColor Cyan
Write-Host "   $($env:COMPUTERNAME)  //  mode: $mode"                            -ForegroundColor DarkGray
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host ""

# ----------------------------------------------------------------
# STEP 1 - Validate and select the entries
# ----------------------------------------------------------------
Write-Title "Step 1: Plan"

$entries = [System.Collections.Generic.List[object]]::new()
$seenName = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($s in $SERVERS) {
    $name = "$($s.Name)".Trim()
    $ip   = "$($s.Ip)".Trim()

    if (-not (Test-HostNameValid $name)) {
        Write-Fail "Bad host name in SERVERS table: '$name' (letters, digits, hyphen only)."
        exit 1
    }
    $parsed = [ref]([System.Net.IPAddress]::Any)
    if (-not [System.Net.IPAddress]::TryParse($ip, $parsed)) {
        Write-Fail "Bad IP address in SERVERS table: '$ip' (entry $name)."
        exit 1
    }
    if (-not $seenName.Add($name)) {
        Write-Fail "Duplicate name in SERVERS table: '$name'."
        exit 1
    }
    $versions = @($s.Revit | Where-Object { $_ -as [int] } | ForEach-Object { [int]$_ })
    if ($versions.Count -eq 0) {
        Write-Fail "Entry '$name' declares no Revit version."
        exit 1
    }

    if ($Only -and -not ($Only | Where-Object { $_ -eq $name -or $_ -eq $ip })) { continue }

    $entries.Add([PSCustomObject]@{
        Name     = $name
        Ip       = $ip
        Versions = $versions
        Note     = "$($s.Note)"
    })
}

if ($entries.Count -eq 0) {
    Write-Fail "Nothing selected. -Only matched no entry in the SERVERS table."
    exit 1
}

Write-Host ("  {0,-4} {1,-16} {2,-16} {3,-16} {4}" -f "#", "IP", "NAME", "REVIT", "NOTE") -ForegroundColor White
Write-Host ("  " + "-" * 62) -ForegroundColor DarkGray
for ($i = 0; $i -lt $entries.Count; $i++) {
    $e = $entries[$i]
    Write-Host ("  {0,-4} {1,-16} {2,-16} {3,-16} {4}" -f ($i + 1), $e.Ip, $e.Name, ($e.Versions -join ", "), $e.Note)
}
Write-Host ""

# Which of them. -Only already filtered above; without it, ask.
if ($script:Interactive -and -not $Only -and $entries.Count -gt 1) {
    $picked = (Read-Host "  Which servers? (numbers like 1,3 - Enter = all)").Trim()
    if ($picked -ne "" -and $picked -notmatch '^(?i)a(ll)?$') {
        $chosen = [System.Collections.Generic.List[object]]::new()
        $takenIdx = [System.Collections.Generic.HashSet[int]]::new()
        foreach ($tok in @($picked -split '[,;\s]+' | Where-Object { $_ })) {
            $n = 0
            if (-not [int]::TryParse($tok, [ref]$n) -or $n -lt 1 -or $n -gt $entries.Count) {
                Write-Fail "Not a valid server number: '$tok' (expected 1-$($entries.Count))."
                exit 1
            }
            if ($takenIdx.Add($n)) { $chosen.Add($entries[$n - 1]) }
        }
        $entries = $chosen
        Write-Info "Selected: $(($entries | ForEach-Object { $_.Name }) -join ', ')"
    }
    Write-Host ""
}

$allNames    = @($entries | ForEach-Object { $_.Name })
$allIps      = @($entries | ForEach-Object { $_.Ip } | Sort-Object -Unique)
$allVersions = @($entries | ForEach-Object { $_.Versions } | Sort-Object -Unique)

# ----------------------------------------------------------------
# What to write, and where. -Target / -RsnEntry win; otherwise ask
# interactively; otherwise the defaults (Rsn / Ip).
# ----------------------------------------------------------------
function Resolve-Choice {
    param([string]$Value, [string[]]$Allowed, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    $match = @($Allowed | Where-Object { $_ -eq $Value.Trim() })
    if ($match.Count -eq 1) { return $match[0] }
    Write-Fail "Bad -$Label value '$Value'. Allowed: $($Allowed -join ', ')."
    exit 1
}

$Target   = Resolve-Choice -Value $Target   -Allowed $TARGET_CHOICES -Label "Target"
$RsnEntry = Resolve-Choice -Value $RsnEntry -Allowed $RSN_CHOICES    -Label "RsnEntry"

if ($script:Interactive -and ($Target -eq "" -or $RsnEntry -eq "")) {
    Write-Host "  What should this run write?" -ForegroundColor White
    Write-Host "    1  RSN.ini only (IP)  - no hosts entry needed  (default)" -ForegroundColor Gray
    Write-Host "    2  hosts + RSN.ini, RSN.ini gets the NAME" -ForegroundColor Gray
    Write-Host "    3  hosts + RSN.ini, RSN.ini gets NAME and IP" -ForegroundColor Gray
    Write-Host "    4  hosts + RSN.ini, RSN.ini gets the IP" -ForegroundColor Gray
    Write-Host "    5  hosts only         - readable names, RSN.ini untouched" -ForegroundColor Gray
    Write-Host ""
    $pick = (Read-Host "  Choice (1-5, Enter = 1)").Trim()
    if ($pick -eq "") { $pick = "1" }
    switch ($pick) {
        "1"     { $Target = "Rsn";   $RsnEntry = "Ip"   }
        "2"     { $Target = "Both";  $RsnEntry = "Name" }
        "3"     { $Target = "Both";  $RsnEntry = "Both" }
        "4"     { $Target = "Both";  $RsnEntry = "Ip"   }
        "5"     { $Target = "Hosts"; $RsnEntry = "Name" }
        default { Write-Fail "Not a valid choice: '$pick'."; exit 1 }
    }
    Write-Host ""
}

if ($Target   -eq "") { $Target   = "Rsn" }
if ($RsnEntry -eq "") { $RsnEntry = "Ip"  }

$doHosts = ($Target -eq "Both" -or $Target -eq "Hosts")
$doRsn   = ($Target -eq "Both" -or $Target -eq "Rsn")

$rsnDesc = switch ($RsnEntry) {
    "Name"  { "server names" }
    "Ip"    { "IP addresses" }
    default { "names and IPs" }
}
if ($doHosts -and $doRsn) { Write-Info "Writing: hosts + RSN.ini ($rsnDesc)." }
elseif ($doHosts)         { Write-Info "Writing: hosts only." }
else                      { Write-Info "Writing: RSN.ini only ($rsnDesc)." }

if ($doRsn -and -not $doHosts -and $RsnEntry -ne "Ip") {
    Write-Warn "RSN.ini will list names, but hosts is not being touched -"
    Write-Warn "those names must already resolve (real DNS, or an earlier run)."
}
Write-Host ""

if ($script:Interactive -and -not $WhatIfPreference) {
    $what = if ($doHosts -and $doRsn) { "hosts and RSN.ini" } elseif ($doHosts) { "hosts" } else { "RSN.ini" }
    $verb = if ($Remove) { "REMOVE these entries from $what" } else { "write these to $what" }
    if ((Read-Host "  $verb ? (Y/N)").Trim().ToUpper() -ne "Y") {
        Write-Info "Cancelled."
        exit 0
    }
}

# ----------------------------------------------------------------
# STEP 2 - hosts file
# ----------------------------------------------------------------
Write-Title "Step 2: hosts file"

if (-not $doHosts) {
    Write-Info "Skipped - RSN.ini only."
}
else {

Write-Info $HOSTS_PATH

if (-not (Test-Path -LiteralPath $HOSTS_PATH)) {
    Write-Fail "hosts file not found: $HOSTS_PATH"
    exit 1
}

$hostLines = Get-FileLine $HOSTS_PATH

# Split into: everything before the managed block, the block, everything after.
$beginIdx = -1
$endIdx   = -1
for ($i = 0; $i -lt $hostLines.Count; $i++) {
    if ($beginIdx -lt 0 -and $hostLines[$i].Trim() -eq $BLOCK_BEGIN) { $beginIdx = $i }
    elseif ($beginIdx -ge 0 -and $hostLines[$i].Trim() -eq $BLOCK_END) { $endIdx = $i; break }
}
if ($beginIdx -ge 0 -and $endIdx -lt 0) {
    Write-Fail "hosts file has an unterminated rs-host block (missing '$BLOCK_END'). Fix it by hand first."
    exit 1
}

$outside = [System.Collections.Generic.List[string]]::new()
for ($i = 0; $i -lt $hostLines.Count; $i++) {
    if ($beginIdx -ge 0 -and $i -ge $beginIdx -and $i -le $endIdx) { continue }
    $outside.Add($hostLines[$i])
}

# Conflicts: the same name mapped elsewhere outside our block. Windows
# resolves by first match, so such a line would silently win over ours.
$conflicts = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $outside.Count; $i++) {
    $line = $outside[$i]
    if ($line -match '^\s*#') { continue }
    $parts = @($line -split '\s+' | Where-Object { $_ })
    if ($parts.Count -lt 2) { continue }
    foreach ($n in $allNames) {
        if ($parts[1..($parts.Count - 1)] -contains $n) {
            $conflicts.Add([PSCustomObject]@{ Index = $i; Line = $line; Name = $n })
        }
    }
}

if ($conflicts.Count -gt 0 -and -not $Remove) {
    Write-Warn "Existing hosts entries use the same name(s) and would win over ours:"
    foreach ($c in $conflicts) { Write-Host "        $($c.Line.Trim())" -ForegroundColor DarkYellow }
    $doComment = $Force
    if (-not $doComment -and $script:Interactive) {
        $doComment = ((Read-Host "  Comment them out? (Y/N)").Trim().ToUpper() -eq "Y")
    }
    if ($doComment) {
        foreach ($c in $conflicts) {
            $outside[$c.Index] = "# rs-host: superseded  " + $outside[$c.Index]
        }
        Write-OK "Commented out $($conflicts.Count) conflicting line(s)."
    } else {
        Write-Warn "Left in place - '$($conflicts[0].Name)' may still resolve to the old address."
    }
}

# Trim trailing blank lines so the block always sits flush at the end.
while ($outside.Count -gt 0 -and [string]::IsNullOrWhiteSpace($outside[$outside.Count - 1])) {
    $outside.RemoveAt($outside.Count - 1)
}

$newHostLines = [System.Collections.Generic.List[string]]::new()
foreach ($l in $outside) { $newHostLines.Add($l) }

if (-not $Remove) {
    $newHostLines.Add("")
    $newHostLines.Add($BLOCK_BEGIN)
    foreach ($e in $entries) {
        $newHostLines.Add(("{0,-16}{1,-24}# Revit {2}" -f $e.Ip, $e.Name, ($e.Versions -join ", ")))
    }
    $newHostLines.Add($BLOCK_END)
}

$before = ($hostLines -join "`r`n")
$after  = ($newHostLines -join "`r`n")

if ($before -eq $after) {
    Write-OK "hosts already correct - no change."
} else {
    foreach ($e in $entries) {
        if ($Remove) { Write-Info "remove : $($e.Name)" }
        else         { Write-Info "$($e.Ip)  ->  $($e.Name)" }
    }
    if (Test-Approved -Caller $PSCmdlet -Target $HOSTS_PATH -Action "Update managed rs-host block") {
        Backup-Once $HOSTS_PATH
        try {
            Write-TextFile -Path $HOSTS_PATH -Lines $newHostLines
            Write-OK $(if ($Remove) { "Managed block removed." } else { "Managed block written ($($entries.Count) name(s))." })
        } catch {
            Write-Fail "Could not write hosts: $($_.Exception.Message)"
            exit 1
        }
    }
}

}   # end: hosts file

# ----------------------------------------------------------------
# STEP 3 - RSN.ini per Revit version
# ----------------------------------------------------------------
Write-Title "Step 3: RSN.ini"

if (-not $doRsn) { Write-Info "Skipped - hosts only." }

foreach ($v in $(if ($doRsn) { $allVersions } else { @() })) {
    $rsn      = Get-RsnPath $v
    $forThis  = @($entries | Where-Object { $_.Versions -contains $v })
    # One line per server, or two when both forms are asked for.
    $wanted   = [System.Collections.Generic.List[string]]::new()
    $seenWanted = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $forThis) {
        # Two servers of the same version may share an IP - list it once.
        if ($RsnEntry -ne "Ip"   -and $seenWanted.Add($e.Name)) { $wanted.Add($e.Name) }
        if ($RsnEntry -ne "Name" -and $seenWanted.Add($e.Ip))   { $wanted.Add($e.Ip)   }
    }
    $revitExe = "C:\Program Files\Autodesk\Revit $v\Revit.exe"

    Write-Host "  Revit $v" -ForegroundColor White
    Write-Info $rsn
    if (-not (Test-Path -LiteralPath $revitExe)) {
        Write-Warn "Revit $v is not installed here - writing the file anyway."
    }

    $rsnLines = Get-FileLine $rsn
    $exists   = Test-Path -LiteralPath $rsn
    $listed   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($line in $rsnLines) {
        $t = $line.Trim()
        if ($t -ne "") { $listed.Add($t) | Out-Null }
    }

    if ($Remove) {
        # The only case that rewrites the file: drop our own lines, in
        # either form (name or IP), and leave everything else untouched.
        $kept       = [System.Collections.Generic.List[string]]::new()
        $removedAny = $false
        foreach ($line in $rsnLines) {
            $t = $line.Trim()
            if ($t -eq "") { continue }
            if (($allNames -contains $t) -or ($allIps -contains $t)) {
                Write-Info "remove : $t"
                $removedAny = $true
                continue
            }
            $kept.Add($t)
        }
        if (-not $removedAny) {
            Write-OK "Nothing of ours listed - no change."
        } elseif (Test-Approved -Caller $PSCmdlet -Target $rsn -Action "Rewrite RSN.ini without managed entries") {
            Backup-Once $rsn
            try {
                if ($kept.Count -eq 0) {
                    Write-TextFile -Path $rsn -Lines @()
                    Write-OK "Emptied (no other servers were listed)."
                } else {
                    Write-TextFile -Path $rsn -Lines $kept
                    Write-OK "Kept: $(($kept) -join ', ')"
                }
            } catch {
                Write-Fail "Could not write RSN.ini: $($_.Exception.Message)"
                exit 1
            }
        }
    }
    else {
        # Append-only. An existing file keeps every line it already has -
        # ours and other people's - and only what is missing is added.
        $toAdd = [System.Collections.Generic.List[string]]::new()
        foreach ($n in $wanted) {
            if ($listed.Contains($n)) { Write-Info "keep   : $n" }
            else { Write-Info "add    : $n"; $toAdd.Add($n) }
        }

        if ($toAdd.Count -eq 0 -and $exists) {
            Write-OK "Already listed: $(($wanted) -join ', ')"
        } elseif (Test-Approved -Caller $PSCmdlet -Target $rsn -Action $(if ($exists) { "Append to RSN.ini" } else { "Create RSN.ini" })) {
            Backup-Once $rsn
            try {
                Add-TextLine -Path $rsn -Lines $toAdd
                if ($exists) { Write-OK "Appended: $(($toAdd) -join ', ')" }
                else         { Write-OK "Created with: $(($toAdd) -join ', ')" }
            } catch {
                Write-Fail "Could not write RSN.ini: $($_.Exception.Message)"
                exit 1
            }
        }
    }
    Write-Host ""
}

# ----------------------------------------------------------------
# STEP 4 - Verify
# ----------------------------------------------------------------
$failedChecks = [System.Collections.Generic.List[string]]::new()

if ($Remove -or $SkipVerify -or $WhatIfPreference) {
    if ($SkipVerify) { Write-Info "Verification skipped (-SkipVerify)." }
} else {
    Write-Title "Step 4: Verify"

    if (Test-Approved -Caller $PSCmdlet -Target "DNS client cache" -Action "Flush") {
        try { ipconfig /flushdns | Out-Null } catch { Write-Verbose "flushdns: $($_.Exception.Message)" }
    }

    # Only meaningful when a name is supposed to work: hosts was written,
    # or RSN.ini lists names and expects them to resolve.
    $checkNames = $doHosts -or ($RsnEntry -ne "Ip")
    if (-not $checkNames) { Write-Info "RSN.ini lists IPs only - no name resolution to check." }

    foreach ($e in $entries) {
        if ($checkNames) {
            # Name resolution comes from hosts, so this is the real check.
            $resolved = $null
            try {
                $resolved = @([System.Net.Dns]::GetHostAddresses($e.Name) | ForEach-Object { $_.IPAddressToString })
            } catch { Write-Verbose "resolve $($e.Name): $($_.Exception.Message)" }

            if (-not $resolved) {
                Write-Fail "$($e.Name) does not resolve."
                $failedChecks.Add("$($e.Name): no resolution")
                continue
            }
            if ($resolved -contains $e.Ip) {
                Write-OK "$($e.Name) -> $($e.Ip)"
            } else {
                Write-Warn "$($e.Name) resolves to $($resolved -join ', ') - expected $($e.Ip)"
                $failedChecks.Add("$($e.Name): resolves to $($resolved -join ', ')")
            }
        }

        # Revit Server answers its admin page over TCP 80.
        $reachable = $false
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $async = $tcp.BeginConnect($e.Ip, 80, $null, $null)
            $reachable = $async.AsyncWaitHandle.WaitOne(3000, $false) -and $tcp.Connected
            $tcp.Close()
        } catch { Write-Verbose "tcp $($e.Ip):80 - $($_.Exception.Message)" }

        $addr = if ($checkNames) { $e.Name } else { $e.Ip }
        if ($reachable) {
            if (-not $checkNames) { Write-OK "$($e.Ip) answers on TCP 80" }
            foreach ($v in $e.Versions) {
                Write-Host "        http://$addr/RevitServerAdmin$v" -ForegroundColor DarkGray
            }
        } else {
            Write-Warn "TCP 80 on $($e.Ip) is not answering (server down, or firewall)."
            $failedChecks.Add("$($e.Name): TCP 80 unreachable")
        }
    }
}

# ----------------------------------------------------------------
# DONE
# ----------------------------------------------------------------
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host "   Done" -ForegroundColor Green
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host ""

if ($failedChecks.Count -gt 0) {
    Write-Warn "Files are written, but some checks failed:"
    foreach ($f in $failedChecks) { Write-Host "        $f" -ForegroundColor DarkYellow }
    Write-Host ""
} elseif (-not $Remove -and -not $WhatIfPreference) {
    if ($doRsn) {
        Write-OK "Restart Revit, then: Open > Revit Server - the servers appear in the list."
    } else {
        Write-OK "Names are resolvable now. RSN.ini was not touched (-Target Hosts)."
    }
    Write-Host ""
}

if ($script:Interactive) {
    Write-Host "  Press Enter to exit." -ForegroundColor DarkGray
    Read-Host | Out-Null
}

if ($failedChecks.Count -gt 0) { exit 2 }
exit 0
