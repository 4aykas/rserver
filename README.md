# rs-tool

Revit Server backup tool. Exports all models from a Revit Server to local `.rvt` files.

```powershell
irm https://tebin.pro/rs | iex
```

> Run as **Administrator** in PowerShell 5.1+

---

## Modes

---

## How it works
REMOTE mode
| +-- scan RSN.ini files on this machine
| C:\ProgramData\Autodesk\Revit Server <VER>\Config\RSN.ini
| %APPDATA%\Autodesk\Revit\Autodesk Revit <VER>\RSN.ini
| (all versions 2020-2027, all user profiles)
| +-- select server from list or enter hostname / IP / FQDN
| +-- REST API http://<server>/RevitServerAdminRESTService<VER>/
| walks the full model tree
| no admin shares, no UNC access required
| +-- revitservertool.exe createLocalRVT -> export each model
| +-- Desktop\RevitServer_RVT_Backup<date><ver><host>
_BACKUP_MANIFEST.txt

LOCAL mode
| +-- server = this machine ($env:COMPUTERNAME)
| +-- same REST API + revitservertool.exe flow
| +-- backup saved to Desktop or C:\RevitBackup if no Desktop

text

---

## Steps

| # | Step | Notes |
|---|---|---|
| 1 | Mode | LOCAL or REMOTE |
| 2 | Server | RSN.ini scan (REMOTE) or localhost (LOCAL) |
| 3 | Tool scan | Finds `revitservertool.exe` for versions 2020-2027 |
| 4 | Version | Pick version matching the Revit Server |
| 5 | Discovery | REST API crawls model tree, fallback to filesystem scan |
| 6 | Destination | Desktop or `C:\RevitBackup` on Windows Server |
| 7 | Export | `createLocalRVT` per model, locked files auto-skipped |
| 8 | Manifest | `_BACKUP_MANIFEST.txt` with per-model results |

---

## revitservertool.exe on Windows Server

`revitservertool.exe` ships with **Revit workstation**, not Revit Server.

Options for LOCAL mode on Windows Server:

```text
  a) Copy the tool from a Revit workstation:
       source : C:\Program Files\Autodesk\Revit 2026\RevitServerToolCommand\
       place  : C:\RevitServerTools\2026\revitservertool.exe

  b) Install Revit on the server
     (not recommended by Autodesk for production servers)

  c) Use REMOTE mode from a Windows 10/11 workstation instead
```

---

## Requirements

| | LOCAL | REMOTE |
|---|---|---|
| OS | Windows Server 2016-2025 | Windows 10 / 11 |
| PowerShell | 5.1+ | 5.1+ |
| Privileges | Administrator | Administrator |
| Revit | `revitservertool.exe` present (see above) | Revit installed |
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
      ProjectA
        Building.rvt
      ProjectB
        Site.rvt
      _BACKUP_MANIFEST.txt
```

Manifest records: date, OS, mode, server, Revit version, tool path, discovery method (REST API or filesystem), and per-model result with file size.

---

## Source

[`rs-tool.ps1`](./rs-tool.ps1) - [`github.com/4aykas/rserver`](https://github.com/4aykas/rserver)
Small wording suggestion
I would slightly rename this section in the final GitHub page:

## revitservertool.exe on Windows Server
to

## Local mode on Windows Server

It reads a bit cleaner and feels less technical as a heading, while keeping the same meaning.

Next, I can make this README a bit more GitHub-beautiful — same content, but tighter spacing, cleaner pseudo-graphics, and more polished wording.

Prepared using Claude Sonnet 4.6
