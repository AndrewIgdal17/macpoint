"""Create a new .pptx on disk by copying a template (.potx, .pptx, etc.)."""

from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path

_ALLOWED_TEMPLATE_SUFFIXES = frozenset({".potx", ".pptx", ".ppt", ".pot", ".potm"})

_TEMPLATE_CT = b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
_PRESENTATION_CT = b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"

_NEEDS_CT_REWRITE = frozenset({".potx", ".potm", ".pot"})


def copy_template_to_new_presentation(template: Path, dest_pptx: Path) -> None:
    """
    Copy OOXML template file to dest_pptx (must be ``.pptx``).

    For ``.potx``/``.potm``/``.pot`` sources, rewrites ``[Content_Types].xml``
    so python-pptx can open the result without a content-type mismatch error.
    """
    template = template.expanduser().resolve()
    dest_pptx = dest_pptx.expanduser().resolve()
    if not template.is_file():
        raise FileNotFoundError(f"Template not found: {template}")
    suf = template.suffix.lower()
    if suf not in _ALLOWED_TEMPLATE_SUFFIXES:
        raise ValueError(
            f"Unsupported template extension {template.suffix!r}; "
            f"allowed: {', '.join(sorted(_ALLOWED_TEMPLATE_SUFFIXES))}"
        )
    if dest_pptx.suffix.lower() != ".pptx":
        raise ValueError("dest_pptx must end with .pptx")
    dest_pptx.parent.mkdir(parents=True, exist_ok=True)

    if suf in _NEEDS_CT_REWRITE:
        tmp = Path(tempfile.mktemp(suffix=".pptx", dir=str(dest_pptx.parent)))
        try:
            with zipfile.ZipFile(template, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "[Content_Types].xml":
                        data = data.replace(_TEMPLATE_CT, _PRESENTATION_CT)
                    zout.writestr(item, data)
            tmp.replace(dest_pptx)
        except BaseException:
            tmp.unlink(missing_ok=True)
            raise
    else:
        shutil.copy2(template, dest_pptx)
