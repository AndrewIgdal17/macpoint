# Changelog

## 0.2.0 (2026-06-21)

Five new tools implemented — MacPoint now has 10/11 tools working.

- `slide_snapshot` — text-only snapshot of shapes, placeholders, text content, and speaker notes per slide
- `add_speaker_notes` — set speaker notes on any slide (overwrites existing)
- `list_templates` — scan macOS template directories for `.potx`/`.pptx`/`.pot`/`.potm` files
- `analyze_template` — enumerate layouts, placeholders, types (with optional dimensions)
- `manage_slide` — delete, move (reorder), and duplicate slides
- `add_animation` remains a documented stub (OOXML animation schema not supported by python-pptx; AppleScript cannot access animation pane)
- Added test suite (13 tests)

## 0.1.0 (2026-05-31)

Initial public release.

- 11 MCP tools aligned with [powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp) reference
- `manage_presentation` — open, create (blank or from `.potx` template), save, save_as, close
- `populate_placeholder` — plain text population via python-pptx
- `add_slide_with_layout` — append slide with named layout
- `switch_slide` — AppleScript slide navigation
- `evaluate` — safe guidance-only (no arbitrary code execution)
- `.potx` template support with automatic content-type fix for python-pptx compatibility
- Stub tools for future implementation: `slide_snapshot`, `add_speaker_notes`, `list_templates`, `analyze_template`, `manage_slide`, `add_animation`
