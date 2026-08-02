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
- **`irm | iex` constraints** (the primary distribution path, so they are
  hard rules): no validation attributes (`ValidateSet`, `ValidateRange`) in
  param blocks — under `iex` they are applied to the default values and an
  empty/zero default throws `ValidationMetadataException` before the script
  runs; validate manually in the body. And `#Requires` lines are silently
  ignored under `iex` — enforce admin with an explicit
  `WindowsPrincipal.IsInRole` check.
- Validate before committing:
  - parse with both PS 5.1 and PS 7 parsers
  - `Invoke-ScriptAnalyzer -Severity Warning,Error` — `PSAvoidUsingWriteHost`
    is accepted (deliberate console UI), everything else should be clean.
- Keep the console output style (Write-Title / Write-OK / Write-Warn helpers,
  two-space indent, aligned columns) consistent with existing steps.
- Script exit codes are part of the contract: 0 = success, 1 = fatal/bad
  input, 2 = some model exports failed.

## Agent skills

### Issue tracker

Issues are tracked as GitHub Issues on `4aykas/rserver` (via the `gh` CLI). See `docs/agents/issue-tracker.md`.

### Triage labels

Default label vocabulary — label strings equal the canonical role names (`needs-triage`, `needs-info`, `ready-for-agent`, `ready-for-human`, `wontfix`). See `docs/agents/triage-labels.md`.

### Domain docs

Single-context layout — one `CONTEXT.md` and `docs/adr/` at the repo root. See `docs/agents/domain.md`.
