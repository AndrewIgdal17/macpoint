"""
MacPoint MCP server — PowerPoint for Mac.

Tool names and parameters mirror the Windows reference server where practical:
https://github.com/Ayushmaniar/powerpoint-mcp (MIT; reference only).
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Union

from mcp.server.fastmcp import FastMCP

from macpoint import state
from macpoint.backends import applescript_ppt, pptx_backend, template_instantiate

mcp = FastMCP("MacPoint")


def _not_impl(name: str) -> str:
    return (
        f"Not implemented on Mac v0: {name}. "
        "See Tools/MacPoint/docs/mac-deltas.md and Projects/MacPoint/Capability matrix."
    )


@mcp.tool()
def manage_presentation(
    action: str,
    file_path: Optional[str] = None,
    save_path: Optional[str] = None,
    template_path: Optional[str] = None,
    presentation_name: Optional[str] = None,
) -> str:
    """
    Open, close, create, save, or save_as the active presentation (PowerPoint for Mac via AppleScript).

    Paths should be POSIX (e.g. /Users/you/talk.pptx).

    **Create from template:** ``action="create"`` with ``template_path`` (``.potx`` / ``.pptx`` / etc.) requires
    ``file_path`` — destination ``.pptx`` (parent directories are created). Copies the template then opens it.

    **Blank create:** ``action="create"`` without ``template_path`` — new empty presentation (AppleScript).

    ``presentation_name`` is reserved for API parity with the Windows reference.
    """
    _ = presentation_name  # reserved for parity with reference API
    action = (action or "").strip().lower()
    try:
        if action == "open":
            if not file_path:
                return "Error: file_path required for open"
            p = Path(file_path).expanduser()
            applescript_ppt.presentation_open(p)
            state.set_last_active(p)
            return f"Opened presentation: {p}"
        if action == "create":
            if template_path:
                if not file_path:
                    return (
                        "Error: when using template_path, file_path is required "
                        "(destination .pptx, e.g. /Users/you/decks/MyDeck.pptx)."
                    )
                tp = Path(template_path).expanduser()
                fp = Path(file_path).expanduser()
                try:
                    template_instantiate.copy_template_to_new_presentation(tp, fp)
                except (OSError, ValueError, FileNotFoundError) as exc:
                    return f"Error: could not instantiate template: {exc}"
                applescript_ppt.presentation_open(fp)
                state.set_last_active(fp)
                return f"Created presentation from template {tp} → {fp} and opened it."
            applescript_ppt.presentation_create()
            ap = applescript_ppt.active_presentation_path()
            if ap:
                state.set_last_active(ap)
            return "Created new presentation (unsaved). Use save_as with save_path when ready."
        if action == "save":
            applescript_ppt.presentation_save()
            ap = applescript_ppt.active_presentation_path()
            if ap:
                state.set_last_active(ap)
            return "Saved active presentation."
        if action == "save_as":
            if not save_path:
                return "Error: save_path required for save_as"
            p = Path(save_path).expanduser()
            p.parent.mkdir(parents=True, exist_ok=True)
            applescript_ppt.presentation_save_as(p)
            state.set_last_active(p)
            return f"Saved active presentation to: {p}"
        if action == "close":
            applescript_ppt.presentation_close(saving="yes")
            state.set_last_active(None)
            return "Closed active presentation (with save)."
        if action == "close_discard":
            applescript_ppt.presentation_close(saving="no")
            state.set_last_active(None)
            return "Closed active presentation without saving."
        return f"Error: unknown action {action!r}. Use open|create|save|save_as|close|close_discard."
    except Exception as exc:  # noqa: BLE001 — surface AppleScript errors to the agent
        return f"Error: {exc}"


@mcp.tool()
def slide_snapshot(
    slide_number: Optional[Union[str, int]] = None,
    include_screenshot: Optional[bool] = False,
    screenshot_filename: Optional[str] = None,
) -> str:
    _ = slide_number, include_screenshot, screenshot_filename
    return _not_impl("slide_snapshot")


@mcp.tool()
def switch_slide(slide_number: Union[str, int]) -> str:
    """Switch to slide (1-based). Best-effort AppleScript; may fail depending on Office version."""
    try:
        n = int(slide_number)
    except (TypeError, ValueError):
        return f"Error: slide_number must be integer, got {slide_number!r}"
    try:
        applescript_ppt.switch_slide(n)
        return f"Switched to slide {n} (if AppleScript succeeded; verify in PowerPoint)."
    except Exception as exc:  # noqa: BLE001
        return f"Error: {exc}"


@mcp.tool()
def add_speaker_notes(slide_number: Union[str, int], notes_text: str) -> str:
    _ = slide_number, notes_text
    return _not_impl("add_speaker_notes")


@mcp.tool()
def list_templates() -> str:
    return _not_impl("list_templates")


@mcp.tool()
def analyze_template(source: str = "current", detailed: bool = False) -> str:
    _ = source, detailed
    return _not_impl("analyze_template")


@mcp.tool()
def add_slide_with_layout(template_name: str, layout_name: str, after_slide: int) -> str:
    """Add a new slide (appended to end) using a named layout from the deck's slide master."""
    _ = template_name, after_slide  # parity params; MacPoint v0 always appends
    path = state.get_last_active()
    if path is None:
        path = applescript_ppt.active_presentation_path()
    if path is None or not path.exists():
        return "Error: no .pptx path known. Open or create a presentation first."
    try:
        return pptx_backend.add_slide(path, layout_name)
    except Exception as exc:  # noqa: BLE001
        return f"Error: {exc}"


@mcp.tool()
def populate_placeholder(
    placeholder_name: str,
    content: str,
    content_type: str = "auto",
    slide_number: Optional[Union[str, int]] = None,
) -> str:
    """
    Populate placeholder text on disk using python-pptx.

    v0: plain text only (crude tag strip). content_type image/plot not supported.
    Requires a known on-disk path: last deck opened via manage_presentation open/save_as,
    or active presentation path if PowerPoint reports it.

    If PowerPoint has the file open with a write lock, this may fail — save, close, then retry.
    """
    if content_type not in ("auto", "text"):
        return f"Error: content_type {content_type!r} not supported on Mac v0 (use text or auto)."
    path = state.get_last_active()
    if path is None:
        path = applescript_ppt.active_presentation_path()
    if path is None or not path.exists():
        return (
            "Error: no .pptx path known. Open a file with manage_presentation action=open "
            "or save_as first, then close the deck in PowerPoint if save fails due to file lock."
        )
    try:
        sn = int(slide_number) if slide_number is not None else 1
    except (TypeError, ValueError):
        return f"Error: slide_number must be integer or null, got {slide_number!r}"
    try:
        msg = pptx_backend.populate_plain_text(path, sn, placeholder_name, content)
        return msg
    except Exception as exc:  # noqa: BLE001
        return f"Error: {exc}"


@mcp.tool()
def manage_slide(
    operation: str,
    slide_number: Union[str, int],
    target_position: Optional[int] = None,
) -> str:
    _ = operation, slide_number, target_position
    return _not_impl("manage_slide")


@mcp.tool()
def evaluate(
    code: str,
    slide_number: Optional[Union[str, int]] = None,
    shape_ref: Optional[str] = None,
    description: Optional[str] = None,
) -> str:
    _ = code, slide_number, shape_ref, description
    return (
        "MacPoint does not execute arbitrary code against PowerPoint. "
        "Use the explicit tools (populate_placeholder, manage_presentation, …). "
        "See docs/mac-deltas.md."
    )


@mcp.tool()
def add_animation(
    shape_name: str,
    effect: str = "fade",
    animate_text: str = "all_at_once",
    slide_number: Optional[Union[str, int]] = None,
) -> str:
    _ = shape_name, effect, animate_text, slide_number
    return _not_impl("add_animation")


def main() -> None:
    mcp.run()
