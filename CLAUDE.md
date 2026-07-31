# rs-tool

Single-file PowerShell tool (`rs-tool.ps1`) that backs up Revit Server models
to local `.rvt` files. Interactive wizard by default; parameter-driven
non-interactive mode for scheduled runs.

## Working on this repo

- Use the `powershell-expert` skill (in `.claude/skills/`) for any PowerShell
  work: naming, parameters, error handling, and live verification of modules
  and cmdlet syntax.
- Target **Windows PowerShell 5.1 compatibility** — no PS7-only syntax
  (no ternary, no `??`). The script runs via `irm https://tebin.pro/rs | iex`
  on customer servers, so `main` is effectively production.
- Validate before committing:
  - parse with both PS 5.1 and PS 7 parsers
  - `Invoke-ScriptAnalyzer -Severity Warning,Error` — `PSAvoidUsingWriteHost`
    is accepted (deliberate console UI), everything else should be clean.
- Keep the console output style (Write-Title / Write-OK / Write-Warn helpers,
  two-space indent, aligned columns) consistent with existing steps.
- Script exit codes are part of the contract: 0 = success, 1 = fatal/bad
  input, 2 = some model exports failed.
