"""Word rendering via docxtpl. Photos are InlineImage objects placed by the
{{ row.PHOTO_BEFORE }} / {{ row.PHOTO_AFTER }} tags in the template's photo table."""
from pathlib import Path

from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage


def render(template_path: Path, context, output_path: Path) -> Path:
    """context may be a dict, or a callable(tpl) -> dict when InlineImages are needed."""
    tpl = DocxTemplate(str(template_path))
    if callable(context):
        context = context(tpl)
    tpl.render(context)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tpl.save(str(output_path))
    return output_path


def image_factory_for(tpl: DocxTemplate, width_in: float = 3.0):
    return lambda path: InlineImage(tpl, str(path), width=Inches(width_in))
