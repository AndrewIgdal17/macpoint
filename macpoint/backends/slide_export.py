"""Export one rendered slide to PNG through PowerPoint."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from macpoint.backends.applescript_ppt import run_applescript

DIALOG_MESSAGE = "PowerPoint did not return within 20s. A dialog may be open."
EXPORT_MODE = "blocked"


def _quote(path: Path) -> str:
    return str(path).replace("\\", "\\\\").replace('"', '\\"')


def _png_script(out_dir: Path) -> str:
    quoted = _quote(out_dir)
    return f'''
-- OUT={out_dir}
tell application "Microsoft PowerPoint"
  set hfsPath to POSIX file "{quoted}" as text
  save active presentation in hfsPath as save as PNG
  return "ok"
end tell
'''


def _pdf_script(pdf_path: Path) -> str:
    quoted = _quote(pdf_path)
    return f'''
-- OUT={pdf_path.parent}
tell application "Microsoft PowerPoint"
  set hfsPath to POSIX file "{quoted}" as text
  save active presentation in hfsPath as save as PDF
  return "ok"
end tell
'''


def _pick_slide_png(produced: Path, slide_number: int) -> Path:
    if produced.is_file() and produced.suffix.lower() == ".png":
        return produced
    if produced.is_dir():
        for name in (f"Slide{slide_number}.PNG", f"Slide{slide_number}.png"):
            candidate = produced / name
            if candidate.exists():
                return candidate
        pngs = sorted(produced.glob("*.PNG")) + sorted(produced.glob("*.png"))
        if len(pngs) == 1:
            return pngs[0]
        if 1 <= slide_number <= len(pngs):
            return pngs[slide_number - 1]
    raise FileNotFoundError(f"No PNG for slide {slide_number} under {produced}")


def _rasterize_pdf(pdf: Path, dest: Path) -> None:
    if shutil.which("pdftoppm") is None:
        raise RuntimeError("pdftoppm is missing. Install it with: brew install poppler")
    prefix = dest.with_suffix("")
    subprocess.run(
        ["pdftoppm", "-png", "-f", "1", "-l", "1", "-r", "144", str(pdf), str(prefix)],
        check=True,
    )
    produced = Path(str(prefix) + "-1.png")
    if not produced.exists():
        raise FileNotFoundError(f"pdftoppm wrote no PNG next to {prefix}")
    produced.replace(dest)


def export_slide_png(pptx_path: Path, slide_number: int, dest: Path, *, run=None, mode: str | None = None) -> Path:
    _ = pptx_path
    chosen = EXPORT_MODE if mode is None else mode
    if chosen == "blocked":
        raise RuntimeError(
            "slide export is blocked. The PowerPoint probe did not find a headless save that writes an image."
        )
    if chosen not in ("png", "pdf"):
        raise RuntimeError(f"Unknown EXPORT_MODE {chosen!r}")
    runner = run_applescript if run is None else run
    dest.parent.mkdir(parents=True, exist_ok=True)
    work = dest.parent / f".export-{dest.stem}"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir()
    try:
        script = _png_script(work) if chosen == "png" else _pdf_script(work / "deck.pdf")
        try:
            runner(script, timeout=20)
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(DIALOG_MESSAGE) from exc
        if chosen == "png":
            src = _pick_slide_png(work, slide_number)
            shutil.copyfile(src, dest)
            return dest
        _rasterize_pdf(work / "deck.pdf", dest)
        return dest
    finally:
        if work.exists():
            shutil.rmtree(work, ignore_errors=True)
