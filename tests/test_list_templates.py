"""Tests for list_templates."""

from pathlib import Path
from unittest.mock import patch

from macpoint.backends.pptx_backend import list_templates


def test_list_templates_finds_files_in_scanned_dirs(tmp_path):
    """list_templates should find .potx files in the scanned directories."""
    templates_dir = tmp_path / "templates"
    templates_dir.mkdir()
    (templates_dir / "MyTemplate.potx").write_bytes(b"fake")
    (templates_dir / "Another.pptx").write_bytes(b"fake")
    (templates_dir / "not_a_template.txt").write_bytes(b"fake")

    with patch("macpoint.backends.pptx_backend._TEMPLATE_SEARCH_DIRS", [templates_dir]):
        result = list_templates()

    assert "MyTemplate" in result
    assert "Another" in result
    assert "not_a_template" not in result
    assert "2 template(s)" in result


def test_list_templates_handles_missing_dirs():
    """list_templates should not crash when directories don't exist."""
    fake_dirs = [Path("/nonexistent/dir1"), Path("/nonexistent/dir2")]
    with patch("macpoint.backends.pptx_backend._TEMPLATE_SEARCH_DIRS", fake_dirs):
        result = list_templates()

    assert "No templates found" in result
    assert "nonexistent" in result
