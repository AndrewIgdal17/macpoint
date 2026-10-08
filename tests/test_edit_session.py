from pathlib import Path

import pytest

from macpoint.edit_session import run_writable


class FakeSession:
    def __init__(self, active: Path | None, slide: int | None = 2):
        self.active = active
        self.slide = slide
        self.calls: list[tuple] = []

    def active_path(self):
        return self.active

    def current_slide(self):
        self.calls.append(("current_slide",))
        return self.slide

    def save(self):
        self.calls.append(("save",))
        return "saved"

    def close(self):
        self.calls.append(("close",))
        self.active = None
        return "closed"

    def open(self, path: Path):
        self.calls.append(("open", path))
        self.active = path
        return "opened"

    def switch(self, slide_number: int):
        self.calls.append(("switch", slide_number))
        return "switched"


def test_closed_deck_writes_only(tmp_path):
    deck = tmp_path / "deck.pptx"
    session = FakeSession(active=None)
    assert run_writable(deck, lambda: "ok", session) == "ok"
    assert session.calls == []


def test_open_deck_saves_closes_writes_reopens(tmp_path):
    deck = tmp_path / "deck.pptx"
    session = FakeSession(active=deck, slide=2)
    during = []

    def write():
        during.append(session.active)
        return "wrote"

    assert run_writable(deck, write, session) == "wrote"
    assert [c[0] for c in session.calls] == ["current_slide", "save", "close", "open", "switch"]
    assert session.calls[-1] == ("switch", 2)
    assert during == [None]


def test_write_error_still_reopens(tmp_path):
    deck = tmp_path / "deck.pptx"
    session = FakeSession(active=deck, slide=1)

    def write():
        raise RuntimeError("boom")

    with pytest.raises(RuntimeError, match="boom"):
        run_writable(deck, write, session)
    assert ("open", deck) in session.calls
    assert ("switch", 1) in session.calls


def test_missing_slide_index_skips_switch(tmp_path):
    deck = tmp_path / "deck.pptx"
    session = FakeSession(active=deck, slide=None)
    assert run_writable(deck, lambda: "ok", session) == "ok"
    assert not any(c[0] == "switch" for c in session.calls)


def test_lock_error_retries_once_when_path_matches(tmp_path):
    deck = tmp_path / "deck.pptx"
    session = FakeSession(active=None, slide=3)
    n = {"count": 0}

    def write():
        n["count"] += 1
        if n["count"] == 1:
            session.active = deck
            raise PermissionError("locked")
        return "ok"

    assert run_writable(deck, write, session) == "ok"
    assert n["count"] == 2
    assert any(c[0] == "close" for c in session.calls)
