#Requires -RunAsAdministrator
<#
.SYNOPSIS
    rs-prep v1 - Revit Server Prerequisites Installer

.DESCRIPTION
    Installs the Windows roles and features Autodesk requires before
    installing Revit Server (2020-2028), per the official prerequisites
    for Windows Server 2022/2025:

      Role          Web Server (IIS)
      Features      ASP.NET 4.8, WCF HTTP Activation, WCF TCP Activation
      Role services ASP, CGI, Server Side Includes,
                    IIS 6 Management Compatibility (metabase, console,
                    scripting tools, WMI compatibility)

    Optionally opens the firewall for Revit Server traffic:
    TCP 80 (REST API / admin console) and inbound ICMPv4 echo
    (required between hosts, accelerators, and workstations).

    Runs on Windows Server AND on Windows 10/11. The same components
    exist on both, but behind different cmdlets and under different
    names, so the script picks the stack by what the host actually has:
    Install-WindowsFeature on Server, Enable-WindowsOptionalFeature
    (DISM) on client Windows. Autodesk supports Revit Server on Windows
    Server only - on a client the script says so and asks before doing
    anything.

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
    Windows Server or Windows 10/11 - needs Install-WindowsFeature or
    Enable-WindowsOptionalFeature, whichever the host provides.
    Exit codes: 0 = ready (reboot may still be pending), 1 = fatal error /
    no feature cmdlets at all, 2 = one or more features failed.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$OpenFirewall,

    [switch]$NonInteractive
)

$script:Interactive = -not $NonInteractive

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
    # terminating error, so the run continued having changed nothing.
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

# ----------------------------------------------------------------
# FEATURE LIST
# Autodesk "Install Server System Prerequisites" (Windows Server 2022,
# unchanged for 2025). The same components exist on Windows 10/11, but
# under DISM's own names and behind a different cmdlet, so every entry
# carries both - see the OS INFO block below.
#
# Server : Install-WindowsFeature resolves dependencies, so implied
#          services (.NET Extensibility, ISAPI, static content, ...)
#          come in with the role and these selections.
# Client : Enable-WindowsOptionalFeature -All does the same job, pulling
#          in each feature's parents (IIS-WebServerRole and friends).
# ----------------------------------------------------------------
$requiredFeatures = @(
    @{ Server = "Web-Server";                Client = "IIS-WebServer";              Label = "Web Server (IIS) role" }
    @{ Server = "NET-Framework-45-ASPNET";   Client = "NetFx4Extended-ASPNET45";    Label = ".NET Framework 4.8 - ASP.NET" }
    @{ Server = "NET-WCF-HTTP-Activation45"; Client = "WCF-HTTP-Activation45";      Label = ".NET Framework 4.8 - WCF HTTP Activation" }
    @{ Server = "NET-WCF-TCP-Activation45";  Client = "WCF-TCP-Activation45";       Label = ".NET Framework 4.8 - WCF TCP Activation" }
    @{ Server = "Web-Asp-Net45";             Client = "IIS-ASPNET45";               Label = "IIS - ASP.NET 4.8" }
    @{ Server = "Web-ASP";                   Client = "IIS-ASP";                    Label = "IIS - ASP" }
    @{ Server = "Web-CGI";                   Client = "IIS-CGI";                    Label = "IIS - CGI" }
    @{ Server = "Web-Includes";              Client = "IIS-ServerSideIncludes";     Label = "IIS - Server Side Includes" }
    # Web-Lgcy-Mgmt-Console (IIS 6 Management Console) is deliberately absent:
    # removed from Windows Server 2025, and not in Autodesk's checkbox list.
    @{ Server = "Web-Metabase";              Client = "IIS-Metabase";               Label = "IIS 6 - Metabase Compatibility" }
    @{ Server = "Web-Lgcy-Scripting";        Client = "IIS-LegacyScripts";          Label = "IIS 6 - Scripting Tools" }
    @{ Server = "Web-WMI";                   Client = "IIS-WMICompatibility";       Label = "IIS 6 - WMI Compatibility" }
    @{ Server = "Web-Mgmt-Console";          Client = "IIS-ManagementConsole";      Label = "IIS - Management Console" }
)

# ----------------------------------------------------------------
# OS INFO
# ----------------------------------------------------------------
$osCaption       = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
$isWindowsServer = $osCaption -match "Server"

# Which feature stack this host speaks. Decided by the cmdlet that is
# actually present, not by the OS caption - a Server Core install or a
# stripped image can disagree with its own name.
$hasServerCmdlets = [bool](Get-Command Install-WindowsFeature -ErrorAction SilentlyContinue)
$hasClientCmdlets = [bool](Get-Command Enable-WindowsOptionalFeature -ErrorAction SilentlyContinue)
$featureStack     = if ($hasServerCmdlets) { "Server" } elseif ($hasClientCmdlets) { "Client" } else { "None" }

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

if ($featureStack -eq "None") {
    Write-Fail "Neither Install-WindowsFeature nor Enable-WindowsOptionalFeature is available."
    Write-Fail "Cannot manage Windows features on this host."
    exit 1
}

if ($isWindowsServer) {
    if ($osCaption -match "2022|2025") {
        Write-OK "Supported: $osCaption"
    } else {
        Write-Warn "Untested on this version: $osCaption"
        Write-Warn "Script targets Windows Server 2022/2025 - feature names should still match."
        if ($script:Interactive) {
            if ((Read-Host "  Continue anyway? (Y/N)").Trim().ToUpper() -ne "Y") { exit 1 }
        }
    }
}
else {
    # Client Windows. Autodesk supports Revit Server on Windows Server
    # only, but the prerequisites themselves exist on 10/11 and install
    # cleanly, which is enough for a lab, a test host, or a workstation
    # that also serves models.
    Write-Warn "Client Windows detected: $osCaption"
    Write-Warn "Autodesk supports Revit Server on Windows Server only - this is unsupported"
    Write-Warn "by them, but the same IIS / .NET prerequisites do exist here and will be"
    Write-Warn "installed under their Windows 10/11 names."
    Write-Host "  (A workstation that only OPENS models needs none of this -" -ForegroundColor DarkGray
    Write-Host "   see rs-tool.ps1 REMOTE mode, and rs-host.ps1 for the server list.)" -ForegroundColor DarkGray
    if ($script:Interactive) {
        Write-Host ""
        if ((Read-Host "  Install the prerequisites here anyway? (Y/N)").Trim().ToUpper() -ne "Y") {
            Write-Info "Cancelled."
            exit 0
        }
    }
}

Write-Info "Feature stack: $featureStack ($(if ($featureStack -eq 'Server') { 'Install-WindowsFeature' } else { 'Enable-WindowsOptionalFeature' }))"

# ----------------------------------------------------------------
# STEP 2 - .NET Framework 4.8
# Ships in-box on Server 2022/2025 and on Windows 10 1903+ / 11;
# verified, not installed.
# ----------------------------------------------------------------
Write-Title "Step 2: .NET Framework 4.8"

$netRelease = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release
if ($netRelease -ge 528040) {
    Write-OK ".NET Framework 4.8+ present (release $netRelease)."
} else {
    Write-Fail ".NET Framework 4.8 not found (release: $netRelease)."
    Write-Host "  It ships with Server 2022/2025 and Windows 10 1903+ / 11. On anything" -ForegroundColor DarkGray
    Write-Host "  older install it from" -ForegroundColor DarkGray
    Write-Host "  https://dotnet.microsoft.com/download/dotnet-framework/net48 first." -ForegroundColor DarkGray
    exit 1
}

# ----------------------------------------------------------------
# STEP 3 - Roles and features
# ----------------------------------------------------------------
Write-Title "Step 3: Roles and Features"

$missing = [System.Collections.Generic.List[string]]::new()
$failed  = [System.Collections.Generic.List[string]]::new()

foreach ($f in $requiredFeatures) {
    $name = if ($featureStack -eq "Server") { $f.Server } else { $f.Client }

    if ($featureStack -eq "Server") {
        $state     = Get-WindowsFeature -Name $name -ErrorAction SilentlyContinue
        $known     = $null -ne $state
        $installed = $known -and $state.Installed
    } else {
        $state     = Get-WindowsOptionalFeature -Online -FeatureName $name -ErrorAction SilentlyContinue
        $known     = $null -ne $state
        # DISM distinguishes Enabled from EnablePending (enabled, awaiting
        # the reboot) - both mean "do not install again".
        $installed = $known -and ($state.State -eq "Enabled" -or $state.State -eq "EnablePending")
    }

    if (-not $known) {
        # One wrong name must not sink the other eleven: report it and
        # carry on, so the run still installs everything it can.
        Write-Fail "Unknown feature name on this OS: $name  ($($f.Label))"
        $failed.Add($name)
    } elseif ($installed) {
        Write-OK "$($f.Label)"
    } else {
        Write-Info "$($f.Label)  - missing"
        $missing.Add($name)
    }
}

$rebootNeeded = $false

if ($missing.Count -eq 0) {
    Write-Host ""
    Write-OK "All required features already installed."
} elseif (Test-Approved -Caller $PSCmdlet -Target ($missing -join ", ") -Action "Install Windows features") {
    Write-Host ""
    Write-Host "  Installing $($missing.Count) feature(s)..." -ForegroundColor White
    Write-Host ""
    if ($featureStack -eq "Server") {
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
    } else {
        # No bulk form here - DISM takes one feature at a time, so a single
        # failure costs one feature instead of the whole batch. -All pulls
        # in the parents, -NoRestart keeps the reboot our decision.
        foreach ($m in $missing) {
            try {
                $r = Enable-WindowsOptionalFeature -Online -FeatureName $m -All -NoRestart -ErrorAction Stop
                Write-OK "Installed: $m"
                if ($r.RestartNeeded) { $rebootNeeded = $true }
            } catch {
                Write-Fail "Install failed for $m : $($_.Exception.Message)"
                $failed.Add($m)
            }
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
        } elseif (Test-Approved -Caller $PSCmdlet -Target $r.Name -Action "New-NetFirewallRule") {
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
    Write-OK "$(if ($isWindowsServer) { 'Server' } else { 'This host' }) is ready for the Revit Server installer."
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
