"""Read PDF text and table candidates and locate regions for visual review."""

from contextlib import redirect_stdout
from itertools import groupby
import re
from statistics import median
import sys

import pymupdf as pdf

from content import NativePage, box_text, prose_line, style_text, table_markdown


MATH_GLYPHS = re.compile(r"[∑∫√∂≈≤≥±≠×÷→⇒]")
UNMAPPED_GLYPHS = re.compile(r"[\x00\ufffd\ue000-\uf8ff]")


def text_blocks(page):
    flags = pdf.TEXTFLAGS_DICT & ~pdf.TEXT_PRESERVE_IMAGES
    return page.get_text("dict", flags=flags, sort=True)["blocks"]


def line_text(spans):
    runs = []
    plain_sizes = [span["size"] for span in spans if not span["flags"] & pdf.TEXT_FONT_SUPERSCRIPT]
    normal_size = max(plain_sizes, default=0)
    baselines = [span["origin"][1] for span in spans if span["size"] >= normal_size * 0.85]
    baseline = median(baselines) if baselines else 0
    for span in spans:
        flags = span["flags"]
        script = ""
        if flags & pdf.TEXT_FONT_SUPERSCRIPT:
            script = "sup"
        elif span["size"] < normal_size * 0.85 and span["origin"][1] - baseline > normal_size * 0.1:
            script = "sub"
        style = (flags & pdf.TEXT_FONT_BOLD, flags & pdf.TEXT_FONT_ITALIC, script)
        runs.append((style, span["text"]))
    parts = []
    for style, group in groupby(runs, key=lambda run: run[0]):
        text = "".join(fragment for _, fragment in group)
        parts.append(style_text(text, *style))
    return "".join(parts)


def contains_span(rectangles, span):
    span_rect = pdf.Rect(span["bbox"])
    return any(rect.contains(span_rect) for rect in rectangles)


def pdf_tables(page, drawings, result, details):
    if not drawings:
        return []
    try:
        # PyMuPDF may print an advisory while detecting tables. Keep stdout for the CLI result.
        with redirect_stdout(sys.stderr):
            tables = page.find_tables(paths=drawings).tables
    except (RuntimeError, ValueError, TypeError):
        result.warnings.add("table detection failed")
        return []
    rectangles = []
    for table in tables:
        cells = table.extract()
        if len(cells) < 2 or len(cells[0]) < 2 or not any(any(row) for row in cells):
            continue
        rows = [[style_text(cell or "") for cell in row] for row in cells]
        markdown = table_markdown(rows)
        if any(None in row for row in cells):
            result.warnings.add("possible merged PDF cells")
        result.tables.append(markdown)
        rectangles.append(pdf.Rect(table.bbox))
        if details:
            box = pdf.Rect(table.bbox) * page.rotation_matrix
            result.details.extend([f"[table candidate {box_text(box)}]", markdown])
    if rectangles:
        result.warnings.add("table grid, cell contents, and styles need comparison")
    return rectangles


def read_pdf_page(page, details=True):
    result = NativePage(page.number + 1,
                        f"{page.rect.width:g}x{page.rect.height:g} pt, displayed coordinates")
    drawings = page.get_drawings()
    images = page.get_image_info()
    rectangles = pdf_tables(page, drawings, result, details)
    blocks = text_blocks(page)
    all_lines = [line for block in blocks if block["type"] == 0 for line in block["lines"]]
    body_lines = []
    table_lines = []
    for line in all_lines:
        table_spans = [span for span in line["spans"] if contains_span(rectangles, span)]
        body_spans = [span for span in line["spans"] if span not in table_spans]
        if body_spans:
            body_lines.append((line, body_spans))
        if table_spans:
            table_lines.append((line, table_spans))
    if len(blocks) > 1 or page.rotation:
        result.warnings.add("check displayed reading order")
    if drawings or images:
        result.warnings.add("visual objects require review")
    if not all_lines and images:
        result.warnings.add("possible scanned or image-only page")
    if all_lines and images:
        for image in images:
            rect = pdf.Rect(image["bbox"]) * page.rotation_matrix
            if rect.get_area() > page.rect.get_area() * 0.8:
                result.warnings.add("page-sized image with text layer")
                break
    title_line = identify_title(page, body_lines)
    body_by_line = {}
    for line, spans in body_lines:
        text = line_text(spans)
        if line is title_line:
            result.title = text
            result.warnings.add("inferred title")
        else:
            body_by_line[id(line)] = prose_line(text)
        if UNMAPPED_GLYPHS.search(text):
            result.warnings.add("unmapped text glyphs")
        if MATH_GLYPHS.search(text) or "<sup>" in text or "<sub>" in text:
            result.warnings.add("check equation and script layout")
        if details:
            result.details.append(packet_line(page, line, spans, text))
    for block in blocks:
        lines = [body_by_line[id(line)] for line in block.get("lines", []) if id(line) in body_by_line]
        if lines:
            result.body.append("\n".join(lines))
    if details:
        for line, spans in table_lines:
            if any(span["flags"] & (pdf.TEXT_FONT_BOLD | pdf.TEXT_FONT_ITALIC) for span in spans):
                result.details.append(packet_line(page, line, spans, line_text(spans), "table style"))
        for image in images:
            box = pdf.Rect(image["bbox"]) * page.rotation_matrix
            result.details.append(f"[image {box_text(box)}]")
        if drawings:
            result.details.append(f"[vector paths: {len(drawings)}; inspect the displayed page]")
    return result


def identify_title(page, lines):
    if len(lines) < 2:
        return None
    first_line, spans = lines[0]
    box = pdf.Rect(first_line["bbox"]) * page.rotation_matrix
    first_size = max(span["size"] for span in spans)
    rest_size = median(span["size"] for _, other_spans in lines[1:] for span in other_spans)
    if box.y0 < page.rect.height * 0.3 and first_size >= rest_size * 1.25:
        return first_line
    return None


def packet_line(page, line, spans, text, label="text"):
    box = pdf.Rect(line["bbox"]) * page.rotation_matrix
    sizes = "/".join(f"{size:g}" for size in sorted({round(span["size"], 1) for span in spans}))
    colors = sorted({span["color"] for span in spans if span["color"]})
    color = " " + "/".join(f"#{value:06x}" for value in colors) if colors else ""
    return f"[{label} {box_text(box)} {sizes}pt{color}] {text}"


def figure_candidates(page):
    """Return likely figure regions in displayed-page points."""
    rotation = page.rotation_matrix
    rectangles = [pdf.Rect(image["bbox"]) for image in page.get_image_info()]
    drawings = page.get_drawings()
    if drawings:
        rectangles.extend(page.cluster_drawings(drawings=drawings, x_tolerance=12, y_tolerance=12))
    rectangles = [(rect * rotation) & page.rect for rect in rectangles]
    rectangles = [rect for rect in rectangles if rect.width >= 20 and rect.height >= 20]
    merged = []
    for rect in rectangles:
        index = 0
        while index < len(merged):
            if expand(rect, 12).intersects(merged[index]):
                rect |= merged.pop(index)
                index = 0
            else:
                index += 1
        merged.append(rect)

    labels = [pdf.Rect(line["bbox"]) * rotation for block in text_blocks(page)
              if block["type"] == 0 for line in block["lines"]]
    candidates = []
    for rect in merged:
        # A page-sized scan is for page review, not an automatic figure candidate.
        if rect.get_area() > page.rect.get_area() * 0.8:
            continue
        nearby = expand(rect, 18)
        grown = pdf.Rect(rect)
        for label in labels:
            if label.width < page.rect.width * 0.6 and nearby.intersects(label):
                if nearby.x0 <= label.x0 and label.x1 <= nearby.x1:
                    grown |= label
        candidates.append(expand(grown, 6) & page.rect)
    return sorted(candidates, key=lambda rect: (rect.y0, rect.x0))


def expand(rect, distance):
    return pdf.Rect(rect.x0 - distance, rect.y0 - distance, rect.x1 + distance, rect.y1 + distance)
