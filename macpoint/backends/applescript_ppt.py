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
    tell active presentation of active window
        save
    end tell
end tell
"""
    return run_applescript(script)


def presentation_save_as(path: Path) -> str:
    ps = _escape_posix_path(path)
    script = f"""
tell application "Microsoft PowerPoint"
    tell active presentation of active window
        save as in POSIX file "{ps}"
    end tell
end tell
"""
    return run_applescript(script)


def presentation_close(saving: str = "yes") -> str:
    """Close active presentation. saving: 'yes' | 'no' | 'ask'."""
    script = f"""
tell application "Microsoft PowerPoint"
    close active presentation of active window saving {saving}
end tell
"""
    return run_applescript(script)


def switch_slide(slide_number: int) -> str:
    """Navigate to 1-based slide index (best-effort AppleScript)."""
    script = f"""
tell application "Microsoft PowerPoint"
    activate
    tell active window
        set view type of active pane to slide view
        go to slide {int(slide_number)} of active presentation
    end tell
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
