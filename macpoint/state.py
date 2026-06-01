"""Process-local state for last active deck path (set by manage_presentation)."""

from pathlib import Path

last_active_pptx: Path | None = None


def set_last_active(path: Path | None) -> None:
    global last_active_pptx
    last_active_pptx = path.resolve() if path is not None else None


def get_last_active() -> Path | None:
    return last_active_pptx
