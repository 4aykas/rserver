# rs-tool

Revit Server backup tool for exporting all server models to local `.rvt` files.

```powershell
irm https://tebin.pro/rs | iex
```

Run in **PowerShell as Administrator**.

------------------------------------------------------------
  rs-tool
------------------------------------------------------------

  [1] LOCAL   run on the Revit Server machine
  [2] REMOTE  run from any PC with Revit installed

  REMOTE mode:
    - scans RSN.ini files for known servers
    - allows manual hostname / IP / FQDN entry
    - uses Revit Server REST API to fetch the model tree
    - exports each model with revitservertool.exe

------------------------------------------------------------

## What it does

| Step | Action |
|---|---|
| 1 | Choose LOCAL or REMOTE mode |
| 2 | Detect Revit Server from `RSN.ini` or enter it manually |
| 3 | Find `revitservertool.exe` for Revit 2020-2027 |
| 4 | Select the Revit version matching the server |
| 5 | Read the full model tree through REST API |
| 6 | Export every model to `Desktop\\RevitServer_RVT_Backup\\...` |
| 7 | Write `_BACKUP_MANIFEST.txt` with results |

## Notes

- No admin shares are required in normal operation.
- Model discovery uses the Revit Server REST API.
- Locked or busy models are skipped and written to the manifest.
- Folder structure is preserved in the backup output.

## Requirements

| Item | Requirement |
|---|---|
| OS | Windows 10 / 11 |
| PowerShell | 5.1+ |
| Rights | Administrator |
| Revit | Revit 2020-2027 installed on the machine running the script |
| Network | Access to the Revit Server REST API |

## Output

```text
Desktop
  RevitServer_RVT_Backup
    20260407_1700_2026_SERVER01
      ProjectA
        Model1.rvt
      ProjectB
        Model2.rvt
      _BACKUP_MANIFEST.txt
```

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Exported |
| `1` | Busy, skipped |
| `5` | Locked, skipped |
| other | Failed |

## Source

- Script: [`rs-tool.ps1`](./rs-tool.ps1)
- Shortcut: `irm https://tebin.pro/rs | iex`
