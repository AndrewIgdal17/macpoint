# MacPoint — manual verification (v0)

Run from vault root after `cd Tools/MacPoint && uv sync`.

1. **MCP starts:** `uv run macpoint` — process stays up (stdio MCP). Cancel with Ctrl+C.
2. **PowerPoint installed:** Microsoft PowerPoint for Mac launches for `manage_presentation` actions.
3. **Automation:** macOS may prompt to allow **Terminal** or **Cursor** to control **Microsoft PowerPoint** — grant in **System Settings → Privacy & Security → Automation**.
4. **manage_presentation open:** Open a copy of a test `.pptx` via tool with POSIX path; confirm deck opens.
4b. **manage_presentation create from template:** `create` with `template_path` pointing at a `.potx` and `file_path` a new `/tmp/macpoint_test.pptx`; confirm PowerPoint opens the new file and layouts exist.
5. **switch_slide:** From MCP client, call `switch_slide` with `2`; confirm slide changes or read stderr-style error in tool result.
6. **populate_placeholder:** With deck **closed** in PowerPoint (avoid lock), after `open` + `close` or `save_as` path known, call `populate_placeholder` for a known shape name on slide 1; reopen deck and confirm text.

Record Office version and macOS version in WORKLOG when validating.
