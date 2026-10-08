"""Convert the legacy bracket-token FCI template into a docxtpl template.

Usage:
    python tools/convert_template.py "templates/FCI AUTOMATION TEMPLATE.docx" \
        "templates/FCI_TEMPLATE_docxtpl.docx"

What it does:
  * [TOKEN]               -> {{ TOKEN }}        (static fields)
  * [TOKEN] inside section -> {{ row.TOKEN }}   (per-issue fields)
  * Fixes typos: [CONTRATOR_RESPONSE] -> CONTRACTOR_RESPONSE, [IX-VOLTAGE] -> IX_VOLTAGE
  * Cover-page [DATE]      -> {{ SUBMISSION_DATE }}
  * [Section_Start] / [END_OF_SECTION] -> {%p for row in rows %} / {%p endfor %}
  * Adds an "Item #.. - .." Heading 2 per issue (matches the TOC entry style)
  * Replaces the floating example photos + caption text boxes with a
    2-column borderless table holding {{ row.PHOTO_BEFORE }} / {{ row.PHOTO_AFTER }}
  * Sets updateFields so Word refreshes the Table of Contents on open
The source file is never modified.
"""
import copy
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}
W = "{%s}" % NS["w"]

TYPO_FIXES = {"CONTRATOR_RESPONSE": "CONTRACTOR_RESPONSE", "IX-VOLTAGE": "IX_VOLTAGE"}
ROW_FIELDS = {"ISSUE_ID", "LOCATION", "INSPECTION_COMMENT", "STATUS", "DATE",
              "CONTRACTOR_RESPONSE", "AGI_COMMENT"}
TOKEN_RE = re.compile(r"\[([A-Za-z][A-Za-z0-9_\-]*)\]")
CONTENT_WIDTH_DXA = 12240 - 1440 - 1440  # letter page minus 1" margins


def para_text_nodes(p):
    """Direct text nodes of a paragraph (not those of nested text boxes)."""
    return p.xpath("./w:r/w:t | ./w:hyperlink/w:r/w:t | ./w:ins/w:r/w:t", namespaces=NS)


def para_text(p):
    return "".join(t.text or "" for t in para_text_nodes(p))


XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def replace_spans(p, pattern, repl):
    """Regex replace across a paragraph's runs. Each replacement is written into
    the run where the match starts, so surrounding run formatting (e.g. a bold
    'SITE:' label next to a plain value) is preserved. Handles tokens split
    across several runs."""
    nodes = para_text_nodes(p)
    if not nodes:
        return False
    texts = [t.text or "" for t in nodes]
    full = "".join(texts)
    matches = list(pattern.finditer(full))
    if not matches:
        return False
    # offsets[i] = start index of node i in full text
    offsets, pos = [], 0
    for s in texts:
        offsets.append(pos)
        pos += len(s)

    def locate(index):
        for i in range(len(nodes) - 1, -1, -1):
            if offsets[i] <= index:
                return i
        return 0

    for m in reversed(matches):
        new = repl(m)
        a, b = m.start(), m.end()
        ia, ib = locate(a), locate(max(b - 1, a))
        if ia == ib:
            s = texts[ia]
            texts[ia] = s[: a - offsets[ia]] + new + s[b - offsets[ia]:]
        else:
            texts[ia] = texts[ia][: a - offsets[ia]] + new
            for i in range(ia + 1, ib):
                texts[i] = ""
            texts[ib] = texts[ib][b - offsets[ib]:]
    for node, s in zip(nodes, texts):
        node.text = s
        node.set(XML_SPACE, "preserve")
    return True


def replace_in_paragraph(p, func):
    """Whole-paragraph rewrite; all text goes into the first run."""
    nodes = para_text_nodes(p)
    if not nodes:
        return False
    old = "".join(t.text or "" for t in nodes)
    new = func(old)
    if new == old:
        return False
    nodes[0].text = new
    nodes[0].set(XML_SPACE, "preserve")
    for t in nodes[1:]:
        t.text = ""
    return True


def set_paragraph_text(p, text):
    replace_in_paragraph(p, lambda _old: text)


def token_repl(in_section):
    def sub(match):
        name = TYPO_FIXES.get(match.group(1), match.group(1)).replace("-", "_")
        if name in ("Section_Start", "END_OF_SECTION"):
            return match.group(0)
        if in_section and name in ROW_FIELDS:
            return "{{ row.%s }}" % name
        if not in_section and name == "DATE":
            return "{{ SUBMISSION_DATE }}"
        return "{{ %s }}" % name
    return sub


def build_photo_table(caption_ppr, caption_rpr):
    half = CONTENT_WIDTH_DXA // 2
    w = NS["w"]

    def cell(text, ppr=None, rpr=None):
        tc = etree.Element(W + "tc")
        tcpr = etree.SubElement(tc, W + "tcPr")
        etree.SubElement(tcpr, W + "tcW", {W + "w": str(half), W + "type": "dxa"})
        p = etree.SubElement(tc, W + "p")
        ppr_el = copy.deepcopy(ppr) if ppr is not None else etree.Element(W + "pPr")
        for old_jc in ppr_el.findall(W + "jc"):
            ppr_el.remove(old_jc)
        jc = etree.Element(W + "jc", {W + "val": "center"})
        # jc sits after spacing/ind/etc. and before rPr in schema order.
        rpr_in_ppr = ppr_el.find(W + "rPr")
        (rpr_in_ppr.addprevious(jc) if rpr_in_ppr is not None else ppr_el.append(jc))
        p.append(ppr_el)
        r = etree.SubElement(p, W + "r")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = etree.SubElement(r, W + "t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return tc

    tbl = etree.Element(W + "tbl", nsmap={"w": w})
    tblpr = etree.SubElement(tbl, W + "tblPr")
    etree.SubElement(tblpr, W + "tblW", {W + "w": str(CONTENT_WIDTH_DXA), W + "type": "dxa"})
    etree.SubElement(tblpr, W + "jc", {W + "val": "center"})
    borders = etree.SubElement(tblpr, W + "tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        etree.SubElement(borders, W + side, {W + "val": "nil"})
    etree.SubElement(tblpr, W + "tblLayout", {W + "type": "fixed"})
    grid = etree.SubElement(tbl, W + "tblGrid")
    for _ in range(2):
        etree.SubElement(grid, W + "gridCol", {W + "w": str(half)})

    row1 = etree.SubElement(tbl, W + "tr")
    trpr = etree.SubElement(row1, W + "trPr")
    etree.SubElement(trpr, W + "cantSplit")
    row1.append(cell("{{ row.PHOTO_BEFORE }}"))
    row1.append(cell("{{ row.PHOTO_AFTER }}"))

    row2 = etree.SubElement(tbl, W + "tr")
    row2.append(cell("FIG {{ row.FIG_BEFORE }}: EXAMPLE FIGURE", caption_ppr, caption_rpr))
    row2.append(cell("FIG {{ row.FIG_AFTER }}: CORRECTED", caption_ppr, caption_rpr))
    return tbl


def new_paragraph(text, ppr=None):
    p = etree.Element(W + "p")
    if ppr is not None:
        p.append(copy.deepcopy(ppr))
    r = etree.SubElement(p, W + "r")
    t = etree.SubElement(r, W + "t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def convert_document(xml_bytes):
    root = etree.fromstring(xml_bytes)
    all_paras = root.iter(W + "p")
    paras = list(all_paras)

    start = next(p for p in paras if para_text(p).strip() == "[Section_Start]")
    end = next(p for p in paras if "[END_OF_SECTION]" in para_text(p))
    container = start.getparent()
    siblings = list(container)
    section = set(siblings[siblings.index(start): siblings.index(end) + 1])

    def in_section(p):
        node = p
        while node is not None and node.getparent() is not container:
            node = node.getparent()
        return node in section

    # 1. TOC placeholder entry -> generic text (Word regenerates the TOC on open).
    for p in paras:
        if p.xpath("ancestor::w:sdt[.//w:docPartGallery[@w:val='Table of Contents']]", namespaces=NS):
            replace_spans(p, re.compile(r"Item #\[ISSUE_ID\]\s*–\s*\[LOCATION\]"),
                          lambda _m: "Inspection Items")

    # 2. Token replacement everywhere (body, text boxes, AlternateContent fallbacks).
    for p in paras:
        replace_spans(p, TOKEN_RE, token_repl(in_section(p)))

    # 3. Loop markers + per-item heading (style copied from an existing Heading 2).
    heading_src = next(
        p for p in paras
        if p.xpath("./w:pPr/w:pStyle[@w:val='Heading2']", namespaces=NS) and para_text(p).strip()
    )
    set_paragraph_text(start, "{%p for row in rows %}")
    heading = copy.deepcopy(heading_src)
    for bm in heading.xpath(".//w:bookmarkStart | .//w:bookmarkEnd", namespaces=NS):
        bm.getparent().remove(bm)
    set_paragraph_text(heading, "Item #{{ row.ISSUE_ID }} – {{ row.LOCATION }}")
    start.addnext(heading)

    # 4. Floating photos + caption boxes -> inline photo table.
    photo_para = next(
        p for p in section
        if p.tag == W + "p" and p.xpath(".//wp:anchor//pic:pic", namespaces=NS)
    )
    caption_p = photo_para.xpath(".//wps:txbx//w:p[w:r/w:t]", namespaces=NS)
    caption_ppr = caption_p[0].find(W + "pPr") if caption_p else None
    caption_rpr = caption_p[0].find("w:r/w:rPr", NS) if caption_p else None
    photo_para.addprevious(build_photo_table(caption_ppr, caption_rpr))
    container.remove(photo_para)

    # 5. End marker: keep the page break, drop the token, close the loop.
    replace_spans(end, re.compile(r"\[END_OF_SECTION\]"), lambda _m: "")
    end.addnext(new_paragraph("{%p endfor %}"))

    leftover = [m for p in root.iter(W + "p") for m in TOKEN_RE.findall(para_text(p))]
    if leftover:
        raise RuntimeError(f"Unconverted tokens remain: {sorted(set(leftover))}")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def enable_update_fields(xml_bytes):
    root = etree.fromstring(xml_bytes)
    if root.find(W + "updateFields") is None:
        el = etree.Element(W + "updateFields", {W + "val": "true"})
        # Schema order: updateFields precedes these elements.
        after = ("hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars",
                 "rsids", "mathPr", "attachedSchema", "themeFontLang",
                 "clrSchemeMapping", "doNotIncludeSubdocsInStats",
                 "doNotAutoCompressPictures", "forceUpgrade", "captions",
                 "readModeInkLockDown", "smartTagType", "schemaLibrary",
                 "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator")
        anchor = next((c for c in root if etree.QName(c).localname in after), None)
        (anchor.addprevious(el) if anchor is not None else root.append(el))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def convert(src, dst):
    src, dst = Path(src), Path(dst)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = convert_document(data)
            elif item.filename == "word/settings.xml":
                data = enable_update_fields(data)
            zout.writestr(item, data)
    return dst


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    print("Wrote", convert(sys.argv[1], sys.argv[2]))
