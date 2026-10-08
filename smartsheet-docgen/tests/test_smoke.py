"""Offline smoke test: renders a generated template with text, a loop, and an image."""
import sys
from pathlib import Path

from docx import Document
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from docgen.render import inline_image, render  # noqa: E402


def test_render_text_loop_and_image(tmp_path):
    tpl_path = tmp_path / "t.docx"
    d = Document()
    d.add_paragraph("{{ PROJECT_NAME }}")
    d.add_paragraph("{% for row in rows %}")
    d.add_paragraph("Item {{ row.ISSUE_ID }} {{ row.PHOTO }}")
    d.add_paragraph("{% endfor %}")
    d.save(tpl_path)

    img = tmp_path / "p.png"
    Image.new("RGB", (40, 30), "red").save(img)

    out = render(
        tpl_path,
        lambda tpl: {
            "PROJECT_NAME": "Test Project",
            "rows": [{"ISSUE_ID": i, "PHOTO": inline_image(tpl, img, 1)} for i in (1, 2)],
        },
        tmp_path / "out.docx",
    )
    doc = Document(out)
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Test Project" in text and "Item 1" in text and "Item 2" in text
    assert len(doc.inline_shapes) == 2
