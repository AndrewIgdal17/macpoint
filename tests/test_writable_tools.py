from pathlib import Path

from macpoint import server
from macpoint.backends import applescript_ppt


def test_notes_go_through_run_writable(monkeypatch, tmp_path):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    seen = {}

    monkeypatch.setattr(server.state, "get_last_active", lambda: deck)
    monkeypatch.setattr(applescript_ppt, "active_presentation_state", lambda: "saved")

    def fake_run(path, write, session=None):
        seen["path"] = path
        return write()

    monkeypatch.setattr(server, "run_writable", fake_run)
    monkeypatch.setattr(
        "macpoint.backends.pptx_backend.add_speaker_notes",
        lambda path, sn, text: f"notes {sn} {text}",
    )
    assert server.add_speaker_notes(2, "hello") == "notes 2 hello"
    assert seen["path"] == deck


def test_unsaved_front_deck_asks_for_save_as(monkeypatch):
    monkeypatch.setattr(server.state, "get_last_active", lambda: None)
    monkeypatch.setattr(applescript_ppt, "active_presentation_path", lambda: None)
    monkeypatch.setattr(applescript_ppt, "active_presentation_state", lambda: "unsaved")
    message = server.populate_placeholder("Title 1", "Hello")
    assert message == (
        "Error: the open deck has no path. Save it with manage_presentation action=save_as first."
    )


def test_state_check_failure_returns_no_path(monkeypatch):
    monkeypatch.setattr(server.state, "get_last_active", lambda: None)
    monkeypatch.setattr(applescript_ppt, "active_presentation_path", lambda: None)

    def raise_state_error():
        raise RuntimeError("PowerPoint not running")

    monkeypatch.setattr(applescript_ppt, "active_presentation_state", raise_state_error)
    assert server.add_speaker_notes(1, "x") == server._NO_PATH


def test_populate_still_rejects_image(monkeypatch, tmp_path):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    monkeypatch.setattr(server.state, "get_last_active", lambda: deck)
    assert "not supported" in server.populate_placeholder("Title 1", "x", content_type="image")


from mcp.server.fastmcp import Image

from macpoint.backends.slide_export import DIALOG_MESSAGE


def test_snapshot_without_screenshot_is_text(monkeypatch, tmp_path):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    monkeypatch.setattr(server.state, "get_last_active", lambda: deck)
    monkeypatch.setattr("macpoint.backends.pptx_backend.slide_snapshot", lambda path, sn: "SLIDE TEXT")
    assert server.slide_snapshot(1, include_screenshot=False) == "SLIDE TEXT"


def test_snapshot_with_screenshot_returns_image(monkeypatch, tmp_path):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    png = tmp_path / "shot.png"
    png.write_bytes(b"png")
    monkeypatch.setattr(server.state, "get_last_active", lambda: deck)
    monkeypatch.setattr(applescript_ppt, "active_presentation_path", lambda: deck)
    monkeypatch.setattr("macpoint.backends.pptx_backend.slide_snapshot", lambda path, sn: "SLIDE TEXT")
    monkeypatch.setattr(server, "export_slide_png", lambda path, sn, dest: png)
    result = server.slide_snapshot(1, include_screenshot=True, screenshot_filename=str(png))
    assert isinstance(result, list)
    assert result[0] == "SLIDE TEXT"
    assert isinstance(result[1], Image)


def test_snapshot_timeout_keeps_text_and_omits_image(monkeypatch, tmp_path):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    monkeypatch.setattr(server.state, "get_last_active", lambda: deck)
    monkeypatch.setattr(applescript_ppt, "active_presentation_path", lambda: deck)
    monkeypatch.setattr("macpoint.backends.pptx_backend.slide_snapshot", lambda path, sn: "SLIDE TEXT")

    def boom(path, sn, dest):
        raise RuntimeError(DIALOG_MESSAGE)

    monkeypatch.setattr(server, "export_slide_png", boom)
    result = server.slide_snapshot(1, include_screenshot=True)
    assert isinstance(result, str)
    assert "SLIDE TEXT" in result
    assert DIALOG_MESSAGE in result


def test_snapshot_opens_deck_when_closed(monkeypatch, tmp_path):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    opened = []
    monkeypatch.setattr(server.state, "get_last_active", lambda: deck)
    monkeypatch.setattr(applescript_ppt, "active_presentation_path", lambda: None)
    monkeypatch.setattr(applescript_ppt, "presentation_open", lambda path: opened.append(path) or "opened")
    monkeypatch.setattr("macpoint.backends.pptx_backend.slide_snapshot", lambda path, sn: "SLIDE TEXT")
    monkeypatch.setattr(server, "export_slide_png", lambda path, sn, dest: dest)
    server.slide_snapshot(1, include_screenshot=True, screenshot_filename=str(tmp_path / "s.png"))
    assert opened == [deck]
