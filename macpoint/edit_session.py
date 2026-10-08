"""Save, close, write, reopen when PowerPoint holds the deck."""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Protocol


class PptSession(Protocol):
    def active_path(self) -> Path | None: ...
    def current_slide(self) -> int | None: ...
    def save(self) -> str: ...
    def close(self) -> str: ...
    def open(self, path: Path) -> str: ...
    def switch(self, slide_number: int) -> str: ...


def _same(a: Path | None, b: Path) -> bool:
    if a is None:
        return False
    return a.expanduser().resolve() == b.expanduser().resolve()


def _is_lock_error(exc: BaseException) -> bool:
    if isinstance(exc, PermissionError):
        return True
    msg = str(exc).lower()
    return "permission" in msg or "locked" in msg or "being used" in msg


def _transact(path: Path, write: Callable[[], str], session: PptSession) -> str:
    slide = session.current_slide()
    session.save()
    session.close()
    try:
        return write()
    finally:
        session.open(path)
        if isinstance(slide, int) and slide >= 1:
            session.switch(slide)


def run_writable(path: Path, write: Callable[[], str], session: PptSession | None = None) -> str:
    if session is None:
        session = _default_session()
    target = path.expanduser()
    if _same(session.active_path(), target):
        return _transact(target, write, session)
    try:
        return write()
    except Exception as exc:
        if _is_lock_error(exc) and _same(session.active_path(), target):
            return _transact(target, write, session)
        raise


def _default_session() -> PptSession:
    from macpoint.backends import applescript_ppt

    class _Live:
        def active_path(self) -> Path | None:
            return applescript_ppt.active_presentation_path()

        def current_slide(self) -> int | None:
            return applescript_ppt.current_slide_index()

        def save(self) -> str:
            return applescript_ppt.presentation_save()

        def close(self) -> str:
            return applescript_ppt.presentation_close(saving="yes")

        def open(self, path: Path) -> str:
            return applescript_ppt.presentation_open(path)

        def switch(self, slide_number: int) -> str:
            return applescript_ppt.switch_slide(slide_number)

    return _Live()
