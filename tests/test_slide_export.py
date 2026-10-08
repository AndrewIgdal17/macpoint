import re
import subprocess
from pathlib import Path

import pytest

from macpoint.backends.slide_export import DIALOG_MESSAGE, export_slide_png


def _run_writing_slide2(script: str, timeout: float = 120):
    assert timeout == 20
    match = re.search(r"-- OUT=(\S+)", script)
    assert match
    out = Path(match.group(1))
    out.mkdir(parents=True, exist_ok=True)
    (out / "Slide2.PNG").write_bytes(b"png-bytes")
    return "ok"


def test_png_mode_copies_slide_file(tmp_path):
    dest = tmp_path / "shot.png"
    result = export_slide_png(tmp_path / "deck.pptx", 2, dest, run=_run_writing_slide2, mode="png")
    assert result == dest
    assert dest.read_bytes() == b"png-bytes"


def test_timeout_raises_dialog_message():
    def run(script: str, timeout: float = 120):
        raise subprocess.TimeoutExpired(cmd="osascript", timeout=timeout)

    with pytest.raises(RuntimeError, match=DIALOG_MESSAGE):
        export_slide_png(Path("/tmp/deck.pptx"), 1, Path("/tmp/shot.png"), run=run, mode="png")


def test_blocked_mode_raises():
    with pytest.raises(RuntimeError, match="blocked"):
        export_slide_png(Path("/tmp/deck.pptx"), 1, Path("/tmp/shot.png"), mode="blocked")
