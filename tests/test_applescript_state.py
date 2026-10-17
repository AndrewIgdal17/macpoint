from macpoint.backends import applescript_ppt


def test_current_slide_index_parses_integer(monkeypatch):
    monkeypatch.setattr(applescript_ppt, "run_applescript", lambda source, timeout=120: "4")
    assert applescript_ppt.current_slide_index() == 4


def test_current_slide_index_empty_is_none(monkeypatch):
    monkeypatch.setattr(applescript_ppt, "run_applescript", lambda source, timeout=120: "")
    assert applescript_ppt.current_slide_index() is None


def test_active_presentation_state_values(monkeypatch):
    monkeypatch.setattr(applescript_ppt, "run_applescript", lambda source, timeout=120: "unsaved")
    assert applescript_ppt.active_presentation_state() == "unsaved"


def test_presentation_save_script(monkeypatch):
    """save must use app-level 'save active presentation of active window', not nested tell."""
    scripts = []
    monkeypatch.setattr(
        "macpoint.backends.applescript_ppt.run_applescript",
        lambda src, timeout=120.0: scripts.append(src) or "",
    )
    from macpoint.backends.applescript_ppt import presentation_save
    presentation_save()
    assert len(scripts) == 1
    src = scripts[0]
    assert "save active presentation of active window" in src
    assert "tell active presentation" not in src


def test_presentation_save_as_script(monkeypatch):
    """save_as must use 'save ... in POSIX file', not 'save as in'."""
    from pathlib import Path
    scripts = []
    monkeypatch.setattr(
        "macpoint.backends.applescript_ppt.run_applescript",
        lambda src, timeout=120.0: scripts.append(src) or "",
    )
    from macpoint.backends.applescript_ppt import presentation_save_as
    presentation_save_as(Path("/tmp/deck.pptx"))
    assert len(scripts) == 1
    src = scripts[0]
    assert "save active presentation of active window in POSIX file" in src
    assert "save as in" not in src
