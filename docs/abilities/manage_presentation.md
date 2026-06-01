# Ability spec: `manage_presentation`

## Purpose

Open, create (blank or from template), save, save-as, and close the **active** presentation in **Microsoft PowerPoint for Mac**, using POSIX paths. Agents call this tool to establish which `.pptx` is “current” for follow-on tools (for example `populate_placeholder`, which uses the last path set here when PowerPoint does not report an active path).

Behavioral reference (Windows / COM): upstream **[powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp)** — same tool name and parameter shape for cross-platform agent prompts; Mac implementation is **AppleScript** + **template file copy**, not COM.

## Parameters (contract)

| Parameter | Type | Required | Notes |
|-----------|------|----------|--------|
| `action` | string | Yes | Normalized with `strip().lower()`. See **MacPoint today** for allowed values. |
| `file_path` | string | For `open`; for `create` when `template_path` is set | POSIX path to `.pptx` (or template source handled separately). |
| `save_path` | string | For `save_as` | Destination `.pptx`; parent directories are created. |
| `template_path` | string | Optional | Used only with `action="create"`. Allowed suffixes per `template_instantiate`: `.potx`, `.pptx`, `.ppt`, `.pot`, `.potm`. Requires `file_path` as the new `.pptx` destination. |
| `presentation_name` | string | Optional | **Reserved for API parity** with the Windows reference; **not used** on Mac today. See **Reference parity audit**. |

## MacPoint today

Implementation: [`macpoint/server.py`](../../macpoint/server.py) (`manage_presentation`, ~lines 28–103).

| `action` | Behavior | Backend |
|----------|----------|---------|
| `open` | Error if `file_path` missing; opens file; sets [`state.set_last_active`](../../macpoint/state.py). | `applescript_ppt.presentation_open` |
| `create` (no `template_path`) | New blank presentation; if AppleScript returns a path, set last active. Message reminds caller to `save_as` when ready. | `applescript_ppt.presentation_create`, `active_presentation_path` |
| `create` + `template_path` | Error if `file_path` missing; `copy_template_to_new_presentation`; open result; set last active. | [`template_instantiate.copy_template_to_new_presentation`](../../macpoint/backends/template_instantiate.py), then `presentation_open` |
| `save` | Save active presentation; refresh last active from `active_presentation_path` if available. | `applescript_ppt.presentation_save` |
| `save_as` | Error if `save_path` missing; `mkdir` parents; save as; set last active to `save_path`. | `applescript_ppt.presentation_save_as` |
| `close` | Close with save; clear last active. | `applescript_ppt.presentation_close(saving="yes")` |
| `close_discard` | Close without saving; clear last active. | `applescript_ppt.presentation_close(saving="no")` |
| (other) | Error listing allowed actions. | — |

AppleScript primitives live in [`macpoint/backends/applescript_ppt.py`](../../macpoint/backends/applescript_ppt.py): `presentation_open`, `presentation_create`, `presentation_save`, `presentation_save_as`, `presentation_close`, `active_presentation_path`.

**Operational notes**

- **Automation:** The process that starts the MCP server must be allowed to control PowerPoint (macOS **Privacy & Security → Automation**).
- **Template copy failures:** If PowerPoint rejects an OOXML copy, operator workaround: open template in app, **Save As** `.pptx`, then `open` — see [mac-deltas.md](../mac-deltas.md#create-from-template-manage_presentation).

## Reference parity audit

- [ ] Read upstream `manage_presentation` (or equivalent) in **`repos/powerpoint-mcp/`** when the clone is available locally, or on GitHub: list every `action` value, defaults, and error messages.
- [ ] **`presentation_name`:** Document upstream semantics (e.g. pick presentation by window title vs path). Decide Mac behavior: implement via AppleScript (if feasible) or document **intentional non-support** in [mac-deltas.md](../mac-deltas.md).
- [ ] Confirm whether reference supports **`close`** vs **`close_discard`** naming parity only, or additional close modes (`ask`, etc.).
- [x] **Unknown-action error string:** `server.py` hint lists `open|create|save|save_as|close|close_discard` (no duplicate token).

Items marked unchecked depend on upstream audit (**#verify** if clone missing).

## Implementation plan

1. **Upstream diff** — Map each reference branch to Mac behavior; note gaps in [mac-deltas.md](../mac-deltas.md).
2. **`presentation_name` (if required)** — If reference uses it to select among multiple open documents: design AppleScript targeting (e.g. by window / document name); update `server.py`, docstrings, and verification checklist; otherwise document N/A on Mac.
3. **Hardening** — Clear, stable error strings for missing paths, bad extensions, AppleScript failures (already surfaced as `Error: …`).
4. **Verification** — Extend [verification-checklist.md](../verification-checklist.md) for any new branch; manual runs on a pinned Office + macOS version.

**Acceptance criteria (when extending behavior)**

- New `action` or parameter semantics match reference intent **or** are explicitly documented as Mac-only in mac-deltas.
- `docs/abilities/manage_presentation.md` and mac-deltas stay aligned with `server.py`.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- Code: [`macpoint/server.py`](../../macpoint/server.py), [`macpoint/backends/applescript_ppt.py`](../../macpoint/backends/applescript_ppt.py), [`macpoint/backends/template_instantiate.py`](../../macpoint/backends/template_instantiate.py)
