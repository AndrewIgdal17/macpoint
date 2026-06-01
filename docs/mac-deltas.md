# MacPoint vs Windows reference (powerpoint-mcp)

Reference: [Ayushmaniar/powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp) (MIT). Local mirror: `repos/powerpoint-mcp/`.

## Platform

- Reference: **Windows**, `pywin32` COM.
- MacPoint: **macOS**, AppleScript (`osascript`) + **`python-pptx`** for some file-level edits.

## Create from template (`manage_presentation`)

- **`action="create"`** + **`template_path`** (``.potx``, ``.pptx``, ``.ppt``, ``.pot``, ``.potm``) **requires** **`file_path`**: destination **``.pptx``** (POSIX). Parent directories are created; template is copied with ``shutil.copy2``, then PowerPoint opens the new file.
- **Blank deck:** ``action="create"`` **without** ``template_path`` — empty presentation (AppleScript), same as before.
- **Limitations:** Copying OOXML is usually fine; if PowerPoint warns or will not open the file, open the template in the app once and **Save As** the ``.pptx`` manually, then use ``action="open"``. If the new file is open in PowerPoint, ``python-pptx`` edits in ``populate_placeholder`` may fail until you **close** the deck or save and retry.

## Tool parity (summary)

| Tool | MacPoint v0 |
|------|-------------|
| `manage_presentation` | Partial — `open`, `create` (blank or **from template** with `.potx` content-type fix), `save`, `save_as`, `close`, `close_discard`. Paths are **POSIX**. |
| `switch_slide` | Partial — best-effort AppleScript; verify on your Office build. |
| `populate_placeholder` | Partial — **plain text** via `python-pptx`; crude tag strip. **Close** the deck in PowerPoint if you hit file locks. `content_type` image/plot not supported. |
| `add_slide_with_layout` | Partial — appends slide with named layout via `python-pptx`. `template_name` ignored (single slide master). `after_slide` accepted for API parity but ignored (always appends). |
| `slide_snapshot`, `add_speaker_notes`, `list_templates`, `analyze_template`, `manage_slide`, `add_animation` | Stubs — return not-implemented message. |
| `evaluate` | **Not arbitrary execution.** Returns guidance to use explicit tools. Per-tool: [evaluate.md](abilities/evaluate.md). |

## Security

Upstream `evaluate` runs Python in a COM context. **MacPoint does not.** Any future “batch” facility must be **allow-listed** operations only. MCP contract: [evaluate.md](abilities/evaluate.md).

## Related

