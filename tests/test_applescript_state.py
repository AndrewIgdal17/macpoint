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
    """save must use 'save presentation 1', not nested tell or 'active presentation of active window'."""
    scripts = []
    monkeypatch.setattr(
        "macpoint.backends.applescript_ppt.run_applescript",
        lambda src, timeout=120.0: scripts.append(src) or "",
    )
    from macpoint.backends.applescript_ppt import presentation_save
    presentation_save()
    assert len(scripts) == 1
    src = scripts[0]
    assert "save presentation 1" in src
    assert "tell active presentation" not in src


def test_presentation_save_as_script(monkeypatch):
    """save_as must use 'save presentation 1 in POSIX file', not 'save as in'."""
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
    assert "save presentation 1 in POSIX file" in src
    assert "save as in" not in src


def test_presentation_close_saving_no_uses_without_saving(monkeypatch):
    """close with saving='no' must use 'without saving', not 'saving no'."""
    scripts = []
    monkeypatch.setattr(
        "macpoint.backends.applescript_ppt.run_applescript",
        lambda src, timeout=120.0: scripts.append(src) or "",
    )
    from macpoint.backends.applescript_ppt import presentation_close
    presentation_close(saving="no")
    src = scripts[0]
    assert "without saving" in src
    assert "saving no" not in src


def test_presentation_close_saving_yes(monkeypatch):
    """close with saving='yes' keeps 'saving yes'."""
    scripts = []
    monkeypatch.setattr(
        "macpoint.backends.applescript_ppt.run_applescript",
        lambda src, timeout=120.0: scripts.append(src) or "",
    )
    from macpoint.backends.applescript_ppt import presentation_close
    presentation_close(saving="yes")
    src = scripts[0]
    assert "saving yes" in src


def test_presentation_close_saving_ask(monkeypatch):
    """close with saving='ask' uses 'saving ask'."""
    scripts = []
    monkeypatch.setattr(
        "macpoint.backends.applescript_ppt.run_applescript",
        lambda src, timeout=120.0: scripts.append(src) or "",
    )
    from macpoint.backends.applescript_ppt import presentation_close
    presentation_close(saving="ask")
    src = scripts[0]
    assert "saving ask" in src


def test_presentation_close_rejects_unknown_saving():
    """Unknown saving values must raise before reaching osascript (injection prevention)."""
    import pytest
    from macpoint.backends.applescript_ppt import presentation_close
    with pytest.raises(ValueError, match="saving"):
        presentation_close(saving="no\nbeep")


def test_switch_slide_targets_view_with_number(monkeypatch):
    """switch_slide must target a view object with a number parameter."""
    scripts = []
    monkeypatch.setattr(
        "macpoint.backends.applescript_ppt.run_applescript",
        lambda src, timeout=120.0: scripts.append(src) or "",
    )
    from macpoint.backends.applescript_ppt import switch_slide
    switch_slide(3)
    src = scripts[0]
    assert "go to slide" in src.lower()
    assert "number 3" in src
    assert "activate" not in src
    assert "of active presentation" not in src.split("go to slide")[1]
