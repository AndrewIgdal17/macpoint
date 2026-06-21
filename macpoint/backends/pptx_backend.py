"""Edit .pptx on disk with python-pptx (close file in PowerPoint first if locked)."""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

from pptx import Presentation

_TEMPLATE_EXTENSIONS = frozenset({".potx", ".pptx", ".ppt", ".pot", ".potm"})

_TEMPLATE_SEARCH_DIRS = [
    Path.home() / "Library/Group Containers/UBF8T346G9.Office/User Content/Templates",
    Path.home() / "Library/Application Support/Microsoft/Office/User Templates",
    Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/Templates"),
]


def list_templates() -> str:
    """Scan macOS template directories and return a formatted list of templates."""
    found: list[dict] = []
    scanned: list[tuple[Path, int]] = []

    for d in _TEMPLATE_SEARCH_DIRS:
        if not d.is_dir():
            scanned.append((d, -1))
            continue
        count = 0
        try:
            for f in sorted(d.rglob("*")):
                if f.is_file() and f.suffix.lower() in _TEMPLATE_EXTENSIONS:
                    stat = f.stat()
                    found.append({
                        "name": f.stem,
                        "path": str(f),
                        "extension": f.suffix.lower(),
                        "size_kb": stat.st_size // 1024,
                        "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d"),
                    })
                    count += 1
        except OSError:
            pass
        scanned.append((d, count))

    lines: list[str] = []
    if found:
        lines.append(f"Found {len(found)} template(s):\n")
        for i, t in enumerate(found, 1):
            lines.append(f"{i}. {t['name']} ({t['extension']})")
            lines.append(f"   {t['path']}")
            lines.append(f"   {t['size_kb']} KB | Modified: {t['modified']}\n")
    else:
        lines.append("No templates found.\n")
        lines.append("To use a template, place a .potx file in one of the directories below,")
        lines.append('or pass a direct path to manage_presentation(action="create", template_path="...").\n')

    lines.append("Directories scanned:")
    for d, count in scanned:
        if count == -1:
            lines.append(f"- {d} (not found)")
        else:
            lines.append(f"- {d} ({count} found)")

    return "\n".join(lines)


def analyze_template(pptx_path: Path, detailed: bool = False) -> str:
    """Analyze a deck's layouts and placeholders. Returns formatted text."""
    prs = Presentation(str(pptx_path))
    lines: list[str] = []
    lines.append(f"Template: {pptx_path.name}")
    lines.append(f"Slides: {len(prs.slides)} | Layouts: {len(prs.slide_layouts)} | Masters: {len(prs.slide_masters)}")
    lines.append("")

    for i, layout in enumerate(prs.slide_layouts):
        ph_count = len(list(layout.placeholders))
        lines.append(f"Layout {i}: \"{layout.name}\" ({ph_count} placeholders)")
        for ph in layout.placeholders:
            idx = ph.placeholder_format.idx
            name = ph.name
            ph_type = str(ph.placeholder_format.type).split(".")[-1].strip("()")
            if detailed:
                left = ph.left / 914400 if ph.left else 0
                top = ph.top / 914400 if ph.top else 0
                width = ph.width / 914400 if ph.width else 0
                height = ph.height / 914400 if ph.height else 0
                lines.append(f"  [{idx}] {name} — {ph_type} ({left:.1f}\", {top:.1f}\", {width:.1f}\" × {height:.1f}\")")
            else:
                lines.append(f"  [{idx}] {name} — {ph_type}")
        lines.append("")

    return "\n".join(lines)


def add_slide(pptx_path: Path, layout_name: str) -> str:
    """Append a new slide using a named layout from the deck's slide master."""
    prs = Presentation(str(pptx_path))
    layout = None
    for sl in prs.slide_layouts:
        if sl.name.lower() == layout_name.lower():
            layout = sl
            break
    if layout is None:
        available = [sl.name for sl in prs.slide_layouts]
        raise ValueError(f"Layout {layout_name!r} not found. Available: {available}")
    prs.slides.add_slide(layout)
    prs.save(str(pptx_path))
    return f"Added slide with layout '{layout.name}' (now {len(prs.slides)} slides total)."


def add_speaker_notes(pptx_path: Path, slide_number: int, notes_text: str) -> str:
    """Set speaker notes on a slide (overwrites existing notes)."""
    prs = Presentation(str(pptx_path))
    if slide_number < 1 or slide_number > len(prs.slides):
        raise ValueError(f"slide_number out of range: {slide_number} (deck has {len(prs.slides)} slides)")
    slide = prs.slides[slide_number - 1]
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = notes_text
    prs.save(str(pptx_path))
    return f"Set speaker notes on slide {slide_number} ({len(notes_text)} chars)."


def slide_snapshot(pptx_path: Path, slide_number: int | None = None) -> str:
    """Return a text snapshot of a slide's shapes, placeholders, and text content."""
    prs = Presentation(str(pptx_path))
    if slide_number is None:
        slide_number = 1
    if slide_number < 1 or slide_number > len(prs.slides):
        raise ValueError(f"slide_number out of range: {slide_number} (deck has {len(prs.slides)} slides)")
    slide = prs.slides[slide_number - 1]
    layout_name = slide.slide_layout.name

    lines: list[str] = []
    lines.append(f"Slide {slide_number}/{len(prs.slides)} — Layout: \"{layout_name}\"")
    lines.append("")

    # Shapes
    lines.append(f"Shapes ({len(slide.shapes)}):")
    for shape in slide.shapes:
        name = getattr(shape, "name", "") or "(unnamed)"
        shape_type = str(type(shape).__name__)
        text = ""
        if getattr(shape, "has_text_frame", False) and shape.text_frame is not None:
            text = shape.text_frame.text[:120]
            if len(shape.text_frame.text) > 120:
                text += "…"
        if text:
            lines.append(f"  • {name} [{shape_type}]: \"{text}\"")
        else:
            lines.append(f"  • {name} [{shape_type}]")

    # Placeholders
    phs = list(slide.placeholders)
    if phs:
        lines.append("")
        lines.append(f"Placeholders ({len(phs)}):")
        for ph in phs:
            idx = ph.placeholder_format.idx
            ph_name = ph.name
            ph_type = str(ph.placeholder_format.type).split(".")[-1].strip("()")
            text = ""
            if ph.has_text_frame:
                text = ph.text_frame.text[:120]
                if len(ph.text_frame.text) > 120:
                    text += "…"
            if text:
                lines.append(f"  [{idx}] {ph_name} ({ph_type}): \"{text}\"")
            else:
                lines.append(f"  [{idx}] {ph_name} ({ph_type})")

    # Speaker notes
    try:
        notes_slide = slide.notes_slide
        notes_text = notes_slide.notes_text_frame.text
        if notes_text.strip():
            lines.append("")
            lines.append(f"Speaker notes: \"{notes_text[:200]}\"")
    except Exception:
        pass

    return "\n".join(lines)


def manage_slide(pptx_path: Path, operation: str, slide_number: int, target_position: int | None = None) -> str:
    """Manage slides: delete, duplicate, or move."""
    prs = Presentation(str(pptx_path))
    slide_count = len(prs.slides)

    if slide_number < 1 or slide_number > slide_count:
        raise ValueError(f"slide_number out of range: {slide_number} (deck has {slide_count} slides)")

    operation = operation.strip().lower()

    if operation == "delete":
        if slide_count == 1:
            raise ValueError("Cannot delete the last remaining slide.")
        rId = prs.slides._sldIdLst[slide_number - 1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[slide_number - 1]
        prs.save(str(pptx_path))
        return f"Deleted slide {slide_number} (now {slide_count - 1} slides)."

    if operation == "move":
        if target_position is None:
            raise ValueError("target_position is required for move operation.")
        if target_position < 1 or target_position > slide_count:
            raise ValueError(f"target_position out of range: {target_position} (deck has {slide_count} slides)")
        if target_position == slide_number:
            return f"Slide {slide_number} is already at position {target_position}."
        sldIdLst = prs.slides._sldIdLst
        elem = sldIdLst[slide_number - 1]
        sldIdLst.remove(elem)
        sldIdLst.insert(target_position - 1, elem)
        prs.save(str(pptx_path))
        return f"Moved slide {slide_number} to position {target_position}."

    if operation == "duplicate":
        import copy
        from lxml import etree
        slide = prs.slides[slide_number - 1]
        layout = slide.slide_layout
        new_slide = prs.slides.add_slide(layout)
        # Copy all shapes from source to new slide by cloning the spTree
        for shape in slide.shapes:
            el = copy.deepcopy(shape._element)
            new_slide.shapes._spTree.append(el)
        # Remove the default placeholder shapes that add_slide created
        # (they'd be duplicates of what we just cloned)
        default_shapes = [s for s in new_slide.shapes if s not in slide.shapes]
        # Actually, simplest: replace the new slide's spTree contents entirely
        new_spTree = new_slide.shapes._spTree
        # Remove all children except the nvGrpSpPr and grpSpPr (first two)
        children = list(new_spTree)
        for child in children[2:]:
            new_spTree.remove(child)
        # Clone source shapes
        src_spTree = slide.shapes._spTree
        for child in list(src_spTree)[2:]:
            new_spTree.append(copy.deepcopy(child))
        prs.save(str(pptx_path))
        new_count = len(prs.slides)
        return f"Duplicated slide {slide_number} (new slide at position {new_count}, deck now has {new_count} slides)."

    raise ValueError(f"Unknown operation {operation!r}. Use: delete, move, duplicate.")


def populate_plain_text(
    pptx_path: Path,
    slide_number: int,
    placeholder_name: str,
    content: str,
) -> str:
    """
    Set plain text on the first shape whose name contains placeholder_name (case-insensitive).

    v0 does not parse HTML/LaTeX; strips tags crudely for display text only.
    """
    prs = Presentation(str(pptx_path))
    if slide_number < 1 or slide_number > len(prs.slides):
        raise ValueError(f"slide_number out of range: {slide_number} (deck has {len(prs.slides)} slides)")
    slide = prs.slides[slide_number - 1]
    needle = placeholder_name.lower()
    updated = 0
    for shape in slide.shapes:
        name = (getattr(shape, "name", "") or "").lower()
        if needle in name and getattr(shape, "has_text_frame", False) and shape.text_frame is not None:
            shape.text_frame.text = _strip_simple_tags(content)
            updated += 1
    if updated == 0:
        for ph in slide.placeholders:
            name = (getattr(ph, "name", "") or "").lower()
            if needle in name and ph.has_text_frame:
                ph.text_frame.text = _strip_simple_tags(content)
                updated += 1
    if updated == 0:
        raise ValueError(
            f"No placeholder/shape matched {placeholder_name!r} on slide {slide_number}. "
            "Use slide_snapshot (when available) or inspect shape names in PowerPoint."
        )
    prs.save(str(pptx_path))
    return f"Updated {updated} shape(s) on slide {slide_number}."


def _strip_simple_tags(s: str) -> str:
    """Very small v0 stripper; not HTML-safe."""
    out = s
    for tag in ("<b>", "</b>", "<i>", "</i>", "<u>", "</u>"):
        out = out.replace(tag, "")
    return out
