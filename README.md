# rs-tool

Revit Server backup tool. Exports all models from a Revit Server to local `.rvt` files.

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
