"""Renders the converted FCI template end to end with real docxtpl + images."""
from pathlib import Path

import pytest
from docx import Document
from PIL import Image

docxtpl = pytest.importorskip("docxtpl")

from docgen.context import build_context  # noqa: E402
from docgen.render import image_factory_for, render  # noqa: E402
from fakes import COLUMNS, ROWS, make_sheet  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "FCI_TEMPLATE_docxtpl.docx"


def test_render_fci_template_with_photos(tmp_path):
    from docgen import settings
    mappings = settings.load_mappings()
    img = tmp_path / "p.jpg"
    Image.new("RGB", (400, 300), "red").save(img)

    out = render(
        TEMPLATE,
        lambda tpl: build_context(
            make_sheet(COLUMNS, ROWS), "Sunny Acres", mappings,
            photo_provider=lambda row: (img, img if row.id == 1000 else None, []),
            image_factory=image_factory_for(tpl, 3.0),
        )[0],
        tmp_path / "out.docx",
    )
    doc = Document(out)
    xml = doc.element.xml
    assert "{{" not in xml and "{%" not in xml
    assert "Item #1" in xml and "Item #2" in xml and "Sunny Acres Solar" in xml
    assert len(doc.inline_shapes) >= 3  # 2 photos on item 1, 1 on item 2
    assert "No photo provided" in xml
