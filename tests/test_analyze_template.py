"""Tests for analyze_template."""

from macpoint.backends.pptx_backend import analyze_template


def test_analyze_template_basic(tmp_pptx):
    """analyze_template should report layouts and placeholders."""
    result = analyze_template(tmp_pptx, detailed=False)
    assert "Layout" in result
    assert "Title" in result


def test_analyze_template_detailed(tmp_pptx):
    """detailed=True should include dimensions."""
    result = analyze_template(tmp_pptx, detailed=True)
    assert '"' in result  # inch markers in dimensions


def test_analyze_template_weg(weg_pptx):
    """Should enumerate all 11 WEG layouts."""
    result = analyze_template(weg_pptx, detailed=False)
    assert "Title Slide WEG" in result
    assert "Section Header" in result
    assert "Layouts: 11" in result
