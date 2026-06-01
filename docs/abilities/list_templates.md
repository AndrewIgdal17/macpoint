# Ability spec: `list_templates`

## Purpose

Enumerate **template file paths** on disk (built-in and user-scope) so agents can pass a POSIX `template_path` into **[`manage_presentation`](manage_presentation.md)** (`action="create"`) without manual Finder browsing. Upstream intent: **[powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp)** — Mac paths and packaging differ from Windows.

## Parameters (contract)

The MCP tool is registered with **no parameters** in [`macpoint/server.py`](../../macpoint/server.py) (`list_templates() -> str`). If the Windows reference later adds filters (scope, theme vs slide template), document deltas here and in [mac-deltas.md](../mac-deltas.md) when MacPoint adopts them.

## MacPoint today


No filesystem walk and no AppleScript query runs.

## Reference parity audit

- [ ] Read upstream `list_templates` in **`repos/powerpoint-mcp/`** when available: return format (plain lines, JSON, table), sorted order, max entries.
- [ ] Which extensions are included (`.potx` only vs `.pot` / `.pptx` / themes)?
- [ ] Whether “installed” vs “custom” is distinguished in the return value.
- [ ] Update [mac-deltas.md](../mac-deltas.md) when Mac behavior and return shape are fixed.

## Implementation plan

1. **Path inventory (#verify)** — On target macOS + Microsoft 365 / Office builds, confirm where PowerPoint stores **user** and **system** templates (e.g. `~/Library/Group Containers/…`, `Application Support/Microsoft/Office/…`). Do not treat web recipes as canonical until verified on your machine.
2. **Suffix allow-list** — Align discovered files with [`template_instantiate._ALLOWED_TEMPLATE_SUFFIXES`](../../macpoint/backends/template_instantiate.py): `.potx`, `.pptx`, `.ppt`, `.pot`, `.potm`.
3. **Walk vs AppleScript** — Prefer a deterministic **filesystem** scan under verified roots; optional AppleScript spike only if it exposes a stable “list templates” API on Mac.
4. **Return shape** — Choose agent-friendly output (e.g. one POSIX path per line vs JSON array); document in mac-deltas and keep stable across minor releases.
5. **Safety** — Avoid following symlinks into sensitive dirs; cap result size; handle permission errors gracefully in tool return text.
6. **Verification** — Add rows to [verification-checklist.md](../verification-checklist.md) when implemented (known template file present in output).
7. **Rollup** — Update [reverse engineering checklist](../../reverse_engineering_checklist.md) row status when the stub is replaced.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- [manage_presentation.md](manage_presentation.md) (consumer of `template_path`)
- Code: [`macpoint/server.py`](../../macpoint/server.py), [`macpoint/backends/template_instantiate.py`](../../macpoint/backends/template_instantiate.py)
