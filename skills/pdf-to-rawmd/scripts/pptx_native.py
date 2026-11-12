"""Read native slide text and tables without converting the presentation."""

from itertools import groupby
import re

from pptx.enum.shapes import MSO_SHAPE_TYPE

from content import NativePage, box_text, literal, prose_line, style_text, table_markdown


DRAWING_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
MATH_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
VISIBLE_SHAPES = {"AlternateContent", "contentPart", "relIds", "oleObj",
                  "audioFile", "videoFile", "cxnSp", "custGeom", "control", "ink"}
MATH_OPS = str.maketrans({"×": r"\times ", "÷": r"\div ", "≤": r"\le ",
                          "≥": r"\ge ", "−": "-", "±": r"\pm "})


def tag(element):
    return element.tag.rsplit("}", 1)[-1]


def math_text(element):
    return "".join(node.text or "" for node in element.iter(MATH_NS + "t"))


def latex(element):
    """Render common Office Math structures. Leave others for visual review."""
    name = tag(element)
    if name == "t":
        text = (element.text or "").translate(MATH_OPS)
        return re.sub(r"([%#&_{}$])", r"\\\1", text)
    if name.endswith("Pr"):
        return ""
    if name in {"oMath", "oMathPara", "r", "e", "num", "den", "sup", "sub", "deg"}:
        return "".join(latex(child) for child in element)
    fields = {tag(child): child for child in element}
    if name == "f":
        properties = fields.get("fPr")
        if properties is not None and len(properties):
            raise ValueError("fraction properties require visual review")
        return r"\frac{" + latex(fields["num"]) + "}{" + latex(fields["den"]) + "}"
    if name in {"sSub", "sSup", "sSubSup"}:
        text = "{" + latex(fields["e"]) + "}"
        for field, operator in (("sub", "_"), ("sup", "^")):
            if field in fields:
                text += operator + "{" + latex(fields[field]) + "}"
        return text
    if name == "rad":
        degree = latex(fields["deg"]) if "deg" in fields else ""
        degree = "[" + degree + "]" if degree else ""
        return r"\sqrt" + degree + "{" + latex(fields["e"]) + "}"
    raise ValueError("unsupported equation structure")


def paragraph_text(paragraph, warnings):
    entries = []
    style_hints = set()
    for item in paragraph._p:
        if tag(item) == "br":
            entries.append((None, "\n"))
            continue
        if tag(item) in {"r", "fld"}:
            properties = item.find(DRAWING_NS + "rPr")
            props = properties.attrib if properties is not None else {}
            color = properties.find(DRAWING_NS + "solidFill") if properties is not None else None
            if color is not None and len(color):
                style_hints.add(f"color {color[0].get('val', '?')}")
            if props.get("u") not in (None, "none"):
                style_hints.add("underline")
                warnings.add("check underline and other rich formatting")
            baseline = int(props.get("baseline", 0))
            script = "sup" if baseline > 0 else "sub" if baseline < 0 else ""
            style = (props.get("b") in ("1", "true"), props.get("i") in ("1", "true"), script)
            entries.append((style, item.findtext(DRAWING_NS + "t", "")))
            continue
        math = next(item.iter(MATH_NS + "oMath"), None)
        if math is not None:
            # A wrapper can contain more than one equation. Do not duplicate nested alternatives.
            for equation in item.iter(MATH_NS + "oMath"):
                warnings.add("check equation layout")
                try:
                    converted = latex(equation)
                    entries.append((None, f"${converted}$"))
                except (KeyError, ValueError):
                    warnings.add("equation structure requires visual transcription")
                    entries.append((None, f"[수식: {literal(math_text(equation))}]"))
    fragments = []
    for style, group in groupby(entries, key=lambda entry: entry[0]):
        text = "".join(fragment for _, fragment in group)
        fragments.append(style_text(text, *style) if style is not None else text)
    return "".join(fragments), style_hints


def bullet_prefix(paragraph, numbering, warnings):
    props = paragraph._p.find(DRAWING_NS + "pPr")
    bullet = None if props is None else next(
        (child for child in props if tag(child) in {"buChar", "buAutoNum", "buNone"}), None
    )
    level = paragraph.level
    for lower in tuple(numbering):
        if lower > level:
            del numbering[lower]
    if bullet is None or tag(bullet) != "buAutoNum":
        numbering.pop(level, None)
        if bullet is not None and tag(bullet) == "buChar":
            return literal(bullet.get("char", "")) + " "
        return ""
    scheme = bullet.get("type", "")
    start = int(bullet.get("startAt", 1))
    previous_scheme, previous = numbering.get(level, (scheme, start - 1))
    number = previous + 1 if previous_scheme == scheme else start
    numbering[level] = (scheme, number)
    schemes = {"arabicPeriod": f"{number}\\.", "arabicParenR": f"{number}\\)",
               "arabicParenBoth": f"({number})", "arabicPlain": str(number)}
    if scheme not in schemes:
        warnings.add("numbering format requires visual review")
        return ""
    return schemes[scheme] + " "


def frame_text(frame, warnings, *, title=False, prose=False):
    lines = []
    hints = set()
    numbering = {}
    for paragraph in frame.paragraphs:
        text, styles = paragraph_text(paragraph, warnings)
        hints.update(styles)
        prefix = bullet_prefix(paragraph, numbering, warnings)
        if paragraph.level:
            hints.add(f"level {paragraph.level}")
        if not text and not prefix:
            lines.append("")
            continue
        indent = "&nbsp;" * (2 * paragraph.level) if paragraph.level else ""
        text = text.replace("\n", "<br>")
        if prose:
            text = prose_line(text)
        lines.append(indent + prefix + text)
    text = ("<br>" if title else "\n").join(lines).strip("\n")
    return text, hints


def read_table(table, warnings):
    rows = [[""] * len(table.columns) for _ in table.rows]
    covered = set()
    merged = []
    for row_number, row in enumerate(table.rows):
        for col_number, cell in enumerate(row.cells):
            if cell.is_spanned:
                if (row_number, col_number) not in covered:
                    warnings.add("table merge geometry needs review")
                if cell.text.strip():
                    warnings.add("covered table cell contains text")
                continue
            text, _ = frame_text(cell.text_frame, warnings)
            if not cell.is_merge_origin:
                rows[row_number][col_number] = text
                continue
            row_end = row_number + cell.span_height
            col_end = col_number + cell.span_width
            merged.append(f"R{row_number + 1}C{col_number + 1}–R{row_end}C{col_end}")
            for row_index in range(row_number, row_end):
                for col_index in range(col_number, col_end):
                    rows[row_index][col_index] = text
                    covered.add((row_index, col_index))
    markdown = table_markdown(rows)
    if merged:
        markdown += "\n<!-- 병합: " + ", ".join(merged) + " -->"
    return markdown


def shape_box(shape):
    if any(value is None for value in (shape.left, shape.top, shape.width, shape.height)):
        return "inherited geometry"
    return box_text((shape.left / 12700, shape.top / 12700,
                     (shape.left + shape.width) / 12700, (shape.top + shape.height) / 12700))


def walk_shapes(shapes, depth=0):
    for shape in shapes:
        yield shape, depth
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(shape.shapes, depth + 1)


def read_slide(deck, number, details=True, notes=False):
    slide = deck.slides[number - 1]
    result = NativePage(number, f"{deck.slide_width.pt:g}x{deck.slide_height.pt:g} pt",
                        hidden=slide._element.get("show") in ("0", "false"))
    result.warnings.add("check layout, master, and inherited styling")
    if len(slide.shapes) > 1:
        result.warnings.add("check displayed reading order")
    if slide._element.find("{*}timing") is not None:
        result.warnings.add("animation or media timing")
    if any(relation.is_external for relation in slide.part.rels.values()):
        result.warnings.add("external relationship")
    if any(tag(element) == "bg" for element in slide._element.iter()):
        result.warnings.add("check slide background")

    title = slide.shapes.title
    for shape, depth in walk_shapes(slide.shapes):
        geometry = (("local " if depth else "") + shape_box(shape)) if details else ""
        shape_type = shape.shape_type
        if shape_type == MSO_SHAPE_TYPE.GROUP:
            result.warnings.add("check group transforms and composition")
            if details:
                result.details.append(f"[group {geometry}]")
            continue
        if shape.has_text_frame:
            is_title = title is not None and shape.shape_id == title.shape_id and depth == 0
            text, styles = frame_text(shape.text_frame, result.warnings, title=is_title, prose=not is_title)
            if text:
                if is_title:
                    result.title = text
                else:
                    result.body.append(text)
                if details:
                    style_hint = " " + ", ".join(sorted(styles)) if styles else ""
                    kind = "title" if is_title else "text"
                    result.details.append(f"[{kind} {geometry}{style_hint}] {text}")
        if shape.has_table:
            table = read_table(shape.table, result.warnings)
            result.tables.append(table)
            if details:
                result.details.extend([f"[table {geometry}]", table])
        if shape.has_chart:
            result.warnings.add("chart needs visual review")
        elif tag(shape._element) == "pic":
            result.warnings.add("picture display and labels require review")
            if details:
                crops = (shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom)
                result.details.append(f"[picture {geometry}; crop fractions {box_text(crops)}; "
                                      f"rotation {shape.rotation:g}°]")
        elif not (shape.has_text_frame or shape.has_table):
            result.warnings.add("unsupported object requires visual review")
            if details:
                result.details.append(f"[object {shape_type} {geometry}]")
        if shape_type in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.FREEFORM, MSO_SHAPE_TYPE.LINE):
            result.warnings.add("shape geometry may encode relationships")
        if any(tag(element) in VISIBLE_SHAPES for element in shape._element.iter()):
            result.warnings.add("complex or alternative content requires visual review")
    if details and notes and slide.has_notes_slide:
        frame = slide.notes_slide.notes_text_frame
        if frame is not None:
            result.notes = frame.text
    return result
