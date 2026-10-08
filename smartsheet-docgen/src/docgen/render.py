"""Word rendering via docxtpl. Images go in as InlineImage objects in the
template context (e.g. {{ row.PHOTO_1 }}), not via run-level XML edits."""
from pathlib import Path

from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage


def inline_image(tpl: DocxTemplate, path: Path, width_in: float = 3.0):
    return InlineImage(tpl, str(path), width=Inches(width_in))


def render(template_path: Path, context: dict, output_path: Path) -> Path:
    tpl = DocxTemplate(str(template_path))
    # Callables let context builders create InlineImages bound to this template.
    if callable(context):
        context = context(tpl)
    tpl.render(context)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tpl.save(str(output_path))
    return output_path
