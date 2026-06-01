# MacPoint

[![PyPI version](https://img.shields.io/pypi/v/macpoint.svg)](https://pypi.org/project/macpoint/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)

**macOS MCP server for Microsoft PowerPoint for Mac** — automate slide decks from [Cursor](https://cursor.sh), Claude, and other MCP clients.

MacPoint gives AI agents the ability to create, edit, and manage PowerPoint presentations on your Mac using the [Model Context Protocol](https://modelcontextprotocol.io). Tool names and parameters are aligned with [powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp) (Windows/COM) for cross-platform agent compatibility.

## What it does

- **11 MCP tools** for PowerPoint automation (open, save, create from template, add slides, populate text, navigate, and more)
- **AppleScript backend** drives the live PowerPoint app (open, save, close, navigate slides)
- **python-pptx backend** handles on-disk edits (populate placeholders, add slides) without needing the app open

## Install

```bash
pip install macpoint
```

Or with [uv](https://docs.astral.sh/uv/):

```bash
uv add macpoint
```

## Configure your MCP client

MacPoint works with **any MCP client** — not just Cursor.

<details>
<summary><strong>Cursor</strong></summary>

Add to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "macpoint": {
      "command": "macpoint",
      "args": []
    }
  }
}
```

Restart Cursor after editing.
</details>

<details>
<summary><strong>Claude Desktop</strong></summary>

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "macpoint": {
      "command": "macpoint",
      "args": []
    }
  }
}
```

Restart Claude Desktop after editing.
</details>

<details>
<summary><strong>Other MCP clients</strong></summary>

Point your client at the `macpoint` command (stdio transport, JSON-RPC). Any client that speaks the [Model Context Protocol](https://modelcontextprotocol.io) will work.
</details>

<details>
<summary><strong>Python (no MCP client)</strong></summary>

You can use the backends directly without an MCP client:

```python
from macpoint.backends.template_instantiate import copy_template_to_new_presentation
from macpoint.backends.pptx_backend import add_slide, populate_plain_text
from pathlib import Path

copy_template_to_new_presentation(Path("template.potx"), Path("deck.pptx"))
add_slide(Path("deck.pptx"), "Title and Content")
populate_plain_text(Path("deck.pptx"), 1, "Title", "Hello World")
```
</details>

On first use, macOS will ask you to grant **Automation** permission (your app → Microsoft PowerPoint) in **System Settings → Privacy & Security → Automation**.

## Tools

| Tool | Status | What it does |
|------|--------|--------------|
| `manage_presentation` | Working | Open, create (blank or from template), save, save_as, close |
| `populate_placeholder` | Working | Set plain text on named placeholders (python-pptx) |
| `add_slide_with_layout` | Working | Append a slide using a named layout |
| `switch_slide` | Working | Navigate to a slide by number (AppleScript) |
| `evaluate` | Safe | Returns guidance (does NOT execute arbitrary code) |
| `slide_snapshot` | Stub | Not yet implemented |
| `add_speaker_notes` | Stub | Not yet implemented |
| `list_templates` | Stub | Not yet implemented |
| `analyze_template` | Stub | Not yet implemented |
| `manage_slide` | Stub | Not yet implemented |
| `add_animation` | Stub | Not yet implemented |

## Example: Create a deck from a template

```
manage_presentation(action="create", template_path="/path/to/template.potx", file_path="/path/to/new-deck.pptx")
add_slide_with_layout(template_name="", layout_name="Title and Content", after_slide=0)
populate_placeholder(placeholder_name="Title", content="Hello World", slide_number=1)
manage_presentation(action="save")
```

## Requirements

- **macOS** (AppleScript is Mac-only)
- **Microsoft PowerPoint for Mac** (installed and launchable)
- **Python 3.10+**

## Development

```bash
git clone https://github.com/AndrewIgdal17/macpoint.git
cd macpoint
python3 -m venv .venv
.venv/bin/pip install -e .
.venv/bin/macpoint
```

## Support

If MacPoint saves you time, consider [sponsoring](https://github.com/sponsors/AndrewIgdal17) the project.

## Attribution

- API reference: [Ayushmaniar/powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp) (MIT). MacPoint is independent code for macOS.
- Microsoft PowerPoint is a trademark of Microsoft Corporation.

## License

[MIT](LICENSE)
