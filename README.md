# rs-tool

Revit Server backup tool. Exports all models from a Revit Server to local `.rvt` files.

Companion scripts:
- [`rs-prep.ps1`](rs-prep.ps1) prepares a fresh host — Windows Server, or Windows 10/11 — for the Revit Server installer, see [Preparing a server](#preparing-a-server-rs-prepps1).
- [`rs-host.ps1`](rs-host.ps1) gives each Revit Server a readable name and registers it in `RSN.ini` for the right Revit versions — `irm https://tebin.pro/rs-host | iex`, see [Naming servers](#naming-servers-rs-hostps1).

```powershell
irm https://tebin.pro/rs | iex
```

> Run as **Administrator** in PowerShell 5.1+

Runs the interactive wizard. For unattended runs (e.g. Task Scheduler), pass parameters:

```powershell
# nightly on the Revit Server host, keep 14 newest backups
.\rs-tool.ps1 -Mode Local -BackupRoot D:\RevitBackup -KeepLast 14 -NonInteractive

# unattended remote backup from a workstation
.\rs-tool.ps1 -Mode Remote -Server REVIT-SRV-01 -RevitVersion 2026 -NonInteractive

# one-liner with parameters
& ([scriptblock]::Create((irm https://tebin.pro/rs))) -Mode Local -NonInteractive
```

| Parameter | Purpose |
|---|---|
| `-Mode Local\|Remote` | Skip the mode prompt |
| `-Server <host>` | Revit Server hostname / IP / FQDN (Remote mode) |
| `-RevitVersion <year>` | Pin the Revit version, e.g. `2026` |
| `-BackupRoot <path>` | Backup destination root (default: Desktop, or `C:\RevitBackup`) |
| `-KeepLast <n>` | After a run, keep only the newest N backup folders |
| `-NonInteractive` | Never prompt — abort with exit 1 on missing input |

Script exit codes: `0` all exported (or skipped busy/locked), `1` fatal error / bad input, `2` one or more exports failed — useful for Task Scheduler monitoring.

---

## Modes

**LOCAL** — run directly on the Revit Server host
- Windows Server 2016 / 2019 / 2022 / 2025
- `revitservertool.exe` is part of Revit Server installation

**REMOTE** — run from a Revit workstation
- Windows 10 / 11
- connects to the server over the network

---

## How it works

```
REMOTE mode
|
+-- scan RSN.ini on this machine (all versions 2020-2028, all users)
|
+-- select server from list or enter hostname / IP / FQDN
|
+-- REST API http://<server>/RevitServerAdminRESTService/
|   walks full model tree - no admin shares needed
|
+-- revitservertool.exe createLocalRVT -> export each model
|
+-- Desktop\RevitServer_RVT_Backup  _BACKUP_MANIFEST.txt

LOCAL mode
|
+-- server = this machine
|
+-- same REST API + revitservertool.exe flow
|
+-- backup saved to Desktop or C:\RevitBackup if no Desktop
     (configurable via $BackupRoot at top of script)
```

---

## Steps

| # | Step | Notes |
|---|---|---|
| 1 | Mode | LOCAL or REMOTE |
| 2 | Server | RSN.ini scan (REMOTE) or localhost (LOCAL) |
| 3 | Tool scan | Finds `revitservertool.exe` for versions 2020-2028 |
| 4 | Version | Pick version matching the Revit Server |
| 5 | Discovery | REST API crawls model tree, fallback to filesystem scan |
| 6 | Destination | Desktop or `C:\RevitBackup` (configurable) |
| 7 | Export | `createLocalRVT` per model, locked files auto-skipped |
| 8 | Manifest | `_BACKUP_MANIFEST.txt` with full per-model results |
| 9 | Retention | With `-KeepLast N`, prunes old backup folders |

---

## Requirements

| | LOCAL | REMOTE |
|---|---|---|
| OS | Windows Server 2016-2025 | Windows 10 / 11 |
| PowerShell | 5.1+ | 5.1+ |
| Privileges | Administrator | Administrator |
| Software | Revit Server installed | Revit installed |
| Network | - | Port 80 open to Revit Server |

---

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Exported successfully |
| `1` | Model busy - skipped |
| `5` | Model locked by a user - skipped |
| other | Export failed - logged to manifest |

---

## Backup output

```text
Desktop (or C:\RevitBackup on Windows Server)
  RevitServer_RVT_Backup
    20260407_0300_2026_REVIT-SRV-01
      ProjectA Building.rvt
      ProjectB Site.rvt
      _BACKUP_MANIFEST.txt
```

Manifest records: date, OS, mode, server, Revit version, tool version, discovery method (REST API or filesystem), and per-model result with file sizes.

---

## Preparing a server (`rs-prep.ps1`)

Installs everything a fresh **Windows Server 2022 / 2025** — or a **Windows 10 / 11** host — needs before running the Revit Server installer, per Autodesk's official prerequisites:

- **Web Server (IIS)** role with ASP, CGI, Server Side Includes, ASP.NET 4.8, and the IIS 6 Management Compatibility set (metabase, console, scripting tools, WMI)
- **.NET Framework 4.8** ASP.NET + WCF HTTP/TCP Activation (presence of 4.8 itself is verified — it ships in-box)
- optional inbound firewall rules: **TCP 80** (REST API / admin console) and **ICMPv4 echo** (Autodesk requires ping between hosts, accelerators, and workstations) — Domain + Private profiles

```powershell
# interactive
.\rs-prep.ps1

# unattended, including firewall rules
.\rs-prep.ps1 -OpenFirewall -NonInteractive

# dry run - show what would change
.\rs-prep.ps1 -WhatIf
```

Run as **Administrator**. Reports whether a restart is required before installing Revit Server. Exit codes: `0` ready, `1` fatal / no feature cmdlets on the host, `2` feature install failed.

**Windows 10 / 11.** The same components exist on client Windows, but under DISM's own names and behind a different cmdlet, so the script picks the stack from what the host actually has rather than from its name — `Install-WindowsFeature` on Server, `Enable-WindowsOptionalFeature -All` on a client. Autodesk supports Revit Server on Windows Server only; on a client the script says so up front and asks before touching anything. On the client stack features install one at a time (DISM has no bulk form), so one bad name costs one feature, not the batch.

| Autodesk / Server feature | Windows 10 / 11 name |
|---|---|
| `Web-Server` | `IIS-WebServer` |
| `NET-Framework-45-ASPNET` | `NetFx4Extended-ASPNET45` |
| `NET-WCF-HTTP-Activation45` | `WCF-HTTP-Activation45` |
| `NET-WCF-TCP-Activation45` | `WCF-TCP-Activation45` |
| `Web-Asp-Net45` | `IIS-ASPNET45` |
| `Web-ASP` | `IIS-ASP` |
| `Web-CGI` | `IIS-CGI` |
| `Web-Includes` | `IIS-ServerSideIncludes` |
| `Web-Metabase` | `IIS-Metabase` |
| `Web-Lgcy-Scripting` | `IIS-LegacyScripts` |
| `Web-WMI` | `IIS-WMICompatibility` |
| `Web-Mgmt-Console` | `IIS-ManagementConsole` |

---

## Naming servers (`rs-host.ps1`)

Run on the **workstation**. Turns bare IPs into names people can read, and puts those names where Revit looks for them:

1. writes `IP  NAME` lines into `C:\Windows\System32\drivers\etc\hosts`, inside a managed block (`# >>> rs-host ... # <<< rs-host <<<`) so re-runs update instead of duplicating;
2. writes each server into `C:\ProgramData\Autodesk\Revit Server <release>\Config\RSN.ini` for the versions it serves — this is the list shown in Revit's *Open → Revit Server*. [Autodesk documents this one path for every role](https://help.autodesk.com/cloudhelp/2025/ENU/Revit-Installation/files/GUID-00163A5A-1379-4743-87B7-DBBBBF00FC93.htm) — workstation, Accelerator, Admin server — and it applies even on a workstation with no Revit Server installed. It is **not** under `…\Autodesk\Revit\Autodesk Revit <release>\`; that is Revit's own application data (`Revit.ini`), and an `RSN.ini` placed there is never read;
3. verifies: flushes the DNS cache, resolves each name, probes TCP 80, prints the admin-console URL.

**You choose what goes where, and for which servers.** Run it plain and it asks both:

```text
  #    IP               NAME             REVIT            NOTE
  1    10.253.9.2       AVRORA           2023             Avrora
  2    10.253.11.26     TEBIN-lab        2025             TEBIN lab
  ...
  Which servers? (numbers like 1,3 - Enter = all):

  What should this run write?
    1  RSN.ini only (IP)  - no hosts entry needed  (default)
    2  hosts + RSN.ini, RSN.ini gets the NAME
    3  hosts + RSN.ini, RSN.ini gets NAME and IP
    4  hosts + RSN.ini, RSN.ini gets the IP
    5  hosts only         - readable names, RSN.ini untouched
```

or pass it non-interactively: `-Only` (which servers), `-Target Rsn|Hosts|Both` (which files) and `-RsnEntry Ip|Name|Both` (what each `RSN.ini` line contains). The default is the minimal one — plain IPs into `RSN.ini`, nothing else touched.

`RSN.ini` is **never rewritten**: an existing file keeps every line it has and only the missing servers are appended; a missing file is created. Only `-Remove` rewrites, and only to drop the script's own lines.

The server list is the `$SERVERS` table at the top of the script — the only part you edit:

```powershell
$SERVERS = @(
    @{ Ip = "10.253.9.2";   Name = "AVRORA";    Revit = @(2023); Note = "Avrora" }
    @{ Ip = "10.253.11.26"; Name = "TEBIN-lab"; Revit = @(2025); Note = "TEBIN lab" }
    @{ Ip = "10.253.9.34";  Name = "CERSANIT";  Revit = @(2025); Note = "Cersanit" }
    @{ Ip = "10.253.11.26"; Name = "DRAGON";    Revit = @(2026); Note = "Dragon" }
)
```

One machine running two Revit Server versions gets **one entry per name** — the same IP twice is expected, not a mistake (`TEBIN-lab` and `DRAGON` above). Names must be valid host names: letters, digits, hyphen.

```powershell
# interactive - asks for servers and mode, straight from the repo (elevated PowerShell)
irm https://tebin.pro/rs-host | iex

# same, but with parameters - `irm | iex` cannot pass them, a scriptblock can
& ([scriptblock]::Create((irm https://tebin.pro/rs-host))) -WhatIf
```

Running from memory this way sidesteps both traps of a downloaded copy: the execution policy never applies (there is no file), and there is no Mark-of-the-Web to `Unblock-File`. If you do save the script, take it from `raw.githubusercontent.com` — the GitHub *page* URL saves the HTML page, which fails to parse.

```powershell
# interactive - asks for servers and mode
.\rs-host.ps1

# dry run - show every line that would be written
.\rs-host.ps1 -WhatIf

# only one server
.\rs-host.ps1 -Only AVRORA

# unattended default - RSN.ini gets plain IPs, hosts untouched
.\rs-host.ps1 -NonInteractive

# names in hosts, and RSN.ini lists name + IP
.\rs-host.ps1 -Target Both -RsnEntry Both -NonInteractive

# unattended cleanup of everything the script manages
.\rs-host.ps1 -Remove -NonInteractive
```

| Parameter | Purpose |
|---|---|
| `-Only <name\|ip>` | Process only these entries (skips the "which servers?" prompt) |
| `-Target Rsn\|Hosts\|Both` | Which files to write (default `Rsn`) |
| `-RsnEntry Ip\|Name\|Both` | What each `RSN.ini` line contains (default `Ip`) |
| `-Remove` | Delete the managed hosts block and the managed names from `RSN.ini` |
| `-Force` | Comment out conflicting hosts lines (same name, other IP) without asking |
| `-SkipVerify` | No DNS / TCP checks |
| `-NonInteractive` | Never prompt |

Both files are backed up (`*.rs-host-<timestamp>.bak`) before the first change of a run, and entries in `RSN.ini` that the script does not manage are preserved.

> Fixed in this version: under `irm | iex` the script body is not an advanced function, so `$PSCmdlet` is `$null` and every `$PSCmdlet.ShouldProcess(...)` gate threw `InvokeMethodOnNull` — a *non-terminating* error, so the run looked successful while writing nothing at all. All gates now go through a helper that falls back to `$WhatIfPreference`. Exit codes: `0` success, `1` fatal / bad input, `2` written but a check failed.
