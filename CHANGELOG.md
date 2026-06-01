# Changelog

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
