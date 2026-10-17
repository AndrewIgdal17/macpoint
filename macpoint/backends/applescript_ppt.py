"""Drive Microsoft PowerPoint for Mac via osascript (AppleScript)."""

from __future__ import annotations

import subprocess
from pathlib import Path


def _escape_posix_path(p: Path) -> str:
    s = str(p.resolve())
    return s.replace("\\", "\\\\").replace('"', '\\"')


def run_applescript(source: str, timeout: float = 120.0) -> str:
    proc = subprocess.run(
        ["osascript", "-e", source],
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip() or "osascript failed"
        raise RuntimeError(err)
    return (proc.stdout or "").strip()


def presentation_open(path: Path) -> str:
    ps = _escape_posix_path(path)
    script = f"""
tell application "Microsoft PowerPoint"
    activate
    open POSIX file "{ps}"
end tell
"""
    return run_applescript(script)


def _open_without_activate(path: Path) -> None:
    """Open a presentation without bringing PowerPoint to the foreground.

    Uses Launch Services (macOS `open -g`) to avoid stealing focus
    during the edit-while-open transaction.
    """
    import subprocess as _sp
    _sp.run(
        ["open", "-g", "-a", "Microsoft PowerPoint", str(path.resolve())],
        check=True,
        capture_output=True,
        text=True,
    )


def presentation_create() -> str:
    script = """
tell application "Microsoft PowerPoint"
    activate
    make new presentation
end tell
"""
    return run_applescript(script)


def presentation_save() -> str:
    script = """
tell application "Microsoft PowerPoint"
    save active presentation of active window
end tell
"""
    return run_applescript(script)


def presentation_save_as(path: Path) -> str:
    ps = _escape_posix_path(path)
    script = f"""
tell application "Microsoft PowerPoint"
    save active presentation of active window in POSIX file "{ps}"
end tell
"""
    return run_applescript(script)


_CLOSE_SAVING_MAP = {
    "yes": "saving yes",
    "no": "without saving",
    "ask": "saving ask",
}


def presentation_close(saving: str = "yes") -> str:
    """Close active presentation. saving: 'yes' | 'no' | 'ask'."""
    clause = _CLOSE_SAVING_MAP.get(saving)
    if clause is None:
        raise ValueError(
            f"Invalid saving={saving!r}. Must be one of: {', '.join(_CLOSE_SAVING_MAP)}"
        )
    script = f"""
tell application "Microsoft PowerPoint"
    close active presentation of active window {clause}
end tell
"""
    return run_applescript(script)


def switch_slide(slide_number: int) -> str:
    """Navigate to 1-based slide index.

    Uses 'go to slide (view) number N' which targets the view object
    per the PowerPoint AppleScript dictionary (command code sPPTDWGs).
    """
    n = int(slide_number)
    script = f"""
tell application "Microsoft PowerPoint"
    go to slide (view of active window) number {n}
end tell
"""
    return run_applescript(script)


def active_presentation_path() -> Path | None:
    """Return POSIX path of active presentation, or None."""
    script = """
tell application "Microsoft PowerPoint"
    try
        set p to full name of active presentation of active window
        return POSIX path of (p as alias)
    on error
        return ""
    end try
end tell
"""
    out = run_applescript(script)
    if not out:
        return None
    return Path(out)


def current_slide_index() -> int | None:
    script = """
tell application "Microsoft PowerPoint"
    try
        return slide index of slide of view of active window
    on error
        return ""
    end try
end tell
"""
    out = run_applescript(script)
    if not out:
        return None
    try:
        return int(out)
    except ValueError:
        return None


def active_presentation_state() -> str:
    """Return 'none', 'unsaved', or 'saved'."""
    script = """
tell application "Microsoft PowerPoint"
    if (count of presentations) is 0 then return "none"
    try
        set fn to full name of active presentation of active window
        if fn is missing value or fn is "" then return "unsaved"
        return "saved"
    on error
        return "unsaved"
    end try
end tell
"""
    out = run_applescript(script).strip()
    if out in ("none", "unsaved", "saved"):
        return out
    return "none"
