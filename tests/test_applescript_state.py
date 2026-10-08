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
