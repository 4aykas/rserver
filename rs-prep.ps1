#Requires -RunAsAdministrator
<#
.SYNOPSIS
    rs-prep v1 - Revit Server Prerequisites Installer

.DESCRIPTION
    Installs the Windows Server roles and features Autodesk requires
    before installing Revit Server (2020-2028), per the official
    prerequisites for Windows Server 2022/2025:

      Role          Web Server (IIS)
      Features      ASP.NET 4.8, WCF HTTP Activation, WCF TCP Activation
      Role services ASP, CGI, Server Side Includes,
                    IIS 6 Management Compatibility (metabase, console,
                    scripting tools, WMI compatibility)

    Optionally opens the firewall for Revit Server traffic:
    TCP 80 (REST API / admin console) and inbound ICMPv4 echo
    (required between hosts, accelerators, and workstations).

    Supports -WhatIf to preview every change without applying it.
    A reboot may be required after feature installation.

.PARAMETER OpenFirewall
    Create inbound firewall rules for TCP 80 and ICMPv4 echo without
    asking. Interactive runs are prompted instead.

.PARAMETER NonInteractive
    Never prompt. Firewall rules are only touched with -OpenFirewall.

.EXAMPLE
    .\rs-prep.ps1
    Interactive: shows feature states, installs what is missing,
    asks about firewall rules.

.EXAMPLE
    .\rs-prep.ps1 -OpenFirewall -NonInteractive
    Unattended full preparation, e.g. via remoting or a deploy pipeline.

.EXAMPLE
    .\rs-prep.ps1 -WhatIf
    Dry run - lists what would be installed or changed.

.NOTES
    Windows Server only (needs Install-WindowsFeature).
    Exit codes: 0 = ready (reboot may still be pending),
    1 = fatal error / unsupported OS, 2 = one or more features failed.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$OpenFirewall,

    [switch]$NonInteractive
)

$script:Interactive = -not $NonInteractive

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

# ----------------------------------------------------------------
# FEATURE LIST
# Autodesk "Install Server System Prerequisites" (Windows Server 2022,
# unchanged for 2025). Install-WindowsFeature resolves dependencies,
# so implied services (.NET Extensibility, ISAPI, static content, ...)
# come in automatically with the role and these selections.
# ----------------------------------------------------------------
$requiredFeatures = @(
    @{ Name = "Web-Server";               Label = "Web Server (IIS) role" }
    @{ Name = "NET-Framework-45-ASPNET";  Label = ".NET Framework 4.8 - ASP.NET" }
    @{ Name = "NET-WCF-HTTP-Activation45";Label = ".NET Framework 4.8 - WCF HTTP Activation" }
    @{ Name = "NET-WCF-TCP-Activation45"; Label = ".NET Framework 4.8 - WCF TCP Activation" }
    @{ Name = "Web-Asp-Net45";            Label = "IIS - ASP.NET 4.8" }
    @{ Name = "Web-ASP";                  Label = "IIS - ASP" }
    @{ Name = "Web-CGI";                  Label = "IIS - CGI" }
    @{ Name = "Web-Includes";             Label = "IIS - Server Side Includes" }
    # Web-Lgcy-Mgmt-Console (IIS 6 Management Console) is deliberately absent:
    # removed from Windows Server 2025, and not in Autodesk's checkbox list.
    @{ Name = "Web-Metabase";             Label = "IIS 6 - Metabase Compatibility" }
    @{ Name = "Web-Lgcy-Scripting";       Label = "IIS 6 - Scripting Tools" }
    @{ Name = "Web-WMI";                  Label = "IIS 6 - WMI Compatibility" }
    @{ Name = "Web-Mgmt-Console";         Label = "IIS - Management Console" }
)

# ----------------------------------------------------------------
# OS INFO
# ----------------------------------------------------------------
$osCaption       = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
$isWindowsServer = $osCaption -match "Server"

# ----------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------
if ($script:Interactive) { Clear-Host }
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host "   rs-prep v1  //  Revit Server Prerequisites Installer"            -ForegroundColor Cyan
Write-Host "   $($env:COMPUTERNAME)  //  $osCaption"                            -ForegroundColor DarkGray
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host ""

# ----------------------------------------------------------------
# STEP 1 - OS check
# ----------------------------------------------------------------
Write-Title "Step 1: Operating System"

if (-not $isWindowsServer) {
    Write-Fail "This is not Windows Server. Revit Server requires Windows Server 2016-2025."
    Write-Host "  (Workstations only need Revit installed - see rs-tool.ps1 REMOTE mode.)" -ForegroundColor DarkGray
    exit 1
}
if ($osCaption -match "2022|2025") {
    Write-OK "Supported: $osCaption"
} else {
    Write-Warn "Untested on this version: $osCaption"
    Write-Warn "Script targets Windows Server 2022/2025 - feature names should still match."
    if ($script:Interactive) {
        if ((Read-Host "  Continue anyway? (Y/N)").Trim().ToUpper() -ne "Y") { exit 1 }
    }
}

if (-not (Get-Command Install-WindowsFeature -ErrorAction SilentlyContinue)) {
    Write-Fail "Install-WindowsFeature not available. Cannot continue."
    exit 1
}

# ----------------------------------------------------------------
# STEP 2 - .NET Framework 4.8
# Ships in-box on Server 2022/2025; verified, not installed.
# ----------------------------------------------------------------
Write-Title "Step 2: .NET Framework 4.8"

$netRelease = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release
if ($netRelease -ge 528040) {
    Write-OK ".NET Framework 4.8+ present (release $netRelease)."
} else {
    Write-Fail ".NET Framework 4.8 not found (release: $netRelease)."
    Write-Host "  It ships with Windows Server 2022/2025. On older servers install it" -ForegroundColor DarkGray
    Write-Host "  from https://dotnet.microsoft.com/download/dotnet-framework/net48 first." -ForegroundColor DarkGray
    exit 1
}

# ----------------------------------------------------------------
# STEP 3 - Roles and features
# ----------------------------------------------------------------
Write-Title "Step 3: Roles and Features"

$missing = [System.Collections.Generic.List[string]]::new()
$failed  = [System.Collections.Generic.List[string]]::new()

foreach ($f in $requiredFeatures) {
    $state = Get-WindowsFeature -Name $f.Name -ErrorAction SilentlyContinue
    if ($null -eq $state) {
        Write-Fail "Unknown feature name: $($f.Name)"
        $failed.Add($f.Name)
    } elseif ($state.Installed) {
        Write-OK "$($f.Label)"
    } else {
        Write-Info "$($f.Label)  - missing"
        $missing.Add($f.Name)
    }
}

$rebootNeeded = $false

if ($missing.Count -eq 0) {
    Write-Host ""
    Write-OK "All required features already installed."
} else {
    Write-Host ""
    Write-Host "  Installing $($missing.Count) feature(s)..." -ForegroundColor White
    Write-Host ""
    if ($PSCmdlet.ShouldProcess(($missing -join ", "), "Install-WindowsFeature")) {
        try {
            $result = Install-WindowsFeature -Name $missing -IncludeManagementTools -ErrorAction Stop
            if ($result.Success) {
                Write-OK "Installed: $(($result.FeatureResult | Select-Object -ExpandProperty DisplayName) -join ', ')"
                if ($result.RestartNeeded -eq 'Yes') { $rebootNeeded = $true }
            } else {
                Write-Fail "Install-WindowsFeature reported failure (exit code: $($result.ExitCode))."
                foreach ($m in $missing) { $failed.Add($m) }
            }
        } catch {
            Write-Fail "Install failed: $($_.Exception.Message)"
            foreach ($m in $missing) { $failed.Add($m) }
        }
    }
}

# ----------------------------------------------------------------
# STEP 4 - Firewall (TCP 80 + ICMPv4 echo)
# ----------------------------------------------------------------
Write-Title "Step 4: Firewall"

$doFirewall = $OpenFirewall
if (-not $doFirewall -and $script:Interactive) {
    Write-Host "  Revit Server needs inbound TCP 80 (REST API / admin console)" -ForegroundColor White
    Write-Host "  and ICMPv4 echo (ping) from workstations and accelerators."   -ForegroundColor White
    Write-Host ""
    $doFirewall = ((Read-Host "  Create these inbound firewall rules? (Y/N)").Trim().ToUpper() -eq "Y")
}

if ($doFirewall) {
    $rules = @(
        @{ Name = "Revit Server - HTTP (TCP 80)"
           Params = @{ Protocol = "TCP"; LocalPort = 80 } }
        @{ Name = "Revit Server - ICMPv4 Echo"
           Params = @{ Protocol = "ICMPv4"; IcmpType = 8 } }
    )
    foreach ($r in $rules) {
        if (Get-NetFirewallRule -DisplayName $r.Name -ErrorAction SilentlyContinue) {
            Write-OK "Rule exists: $($r.Name)"
        } elseif ($PSCmdlet.ShouldProcess($r.Name, "New-NetFirewallRule")) {
            try {
                $ruleParams = $r.Params
                New-NetFirewallRule -DisplayName $r.Name -Direction Inbound -Action Allow -Profile Domain,Private @ruleParams -ErrorAction Stop | Out-Null
                Write-OK "Rule created: $($r.Name)  (Domain + Private profiles)"
            } catch {
                Write-Warn "Could not create rule '$($r.Name)': $($_.Exception.Message)"
            }
        }
    }
} else {
    Write-Info "Skipped. Re-run with -OpenFirewall, or open TCP 80 + ICMPv4 manually."
}

# ----------------------------------------------------------------
# DONE
# ----------------------------------------------------------------
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host "   Done" -ForegroundColor Green
Write-Host "  ================================================================" -ForegroundColor DarkCyan
Write-Host ""
if ($failed.Count -gt 0) {
    Write-Fail "Failed features: $($failed -join ', ')"
    Write-Host ""
}
if ($rebootNeeded) {
    Write-Warn "A RESTART IS REQUIRED before installing Revit Server."
} elseif ($failed.Count -eq 0) {
    Write-OK "Server is ready for the Revit Server installer."
}
Write-Host ""
Write-Host "  Next: run the Revit Server installer (matching your Revit version)," -ForegroundColor White
Write-Host "  then verify with:  http://localhost/RevitServerAdmin<version>" -ForegroundColor White
Write-Host ""

if ($script:Interactive) {
    Write-Host "  Press Enter to exit." -ForegroundColor DarkGray
    Read-Host | Out-Null
}

if ($failed.Count -gt 0) { exit 2 }
exit 0
