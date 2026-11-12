"""Shared Yama print layout and fail-closed Word pagination quality checks.

The caller records len(document.element.body) before rendering each question,
then calls add_block(start, key). save() writes the document and its bound QC JSON.
Source images are read only; panel crops live in an isolated temporary directory.
"""
import hashlib
import io
import json
import math
import re
import unicodedata
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shape import InlineShape
from docx.shared import Cm
from docx.text.paragraph import Paragraph
from PIL import Image, ImageStat, JpegImagePlugin
import docx.image
from docx.image.image import BaseImageHeader


class _JpegWithoutAppHeader(BaseImageHeader):
    """Read a valid JPEG without JFIF/Exif while embedding its original bytes."""
    content_type = "image/jpeg"
    default_ext = "jpg"

    @classmethod
    def from_stream(cls, stream):
        stream.seek(0)
        image = Image.open(stream)
        if image.format != "JPEG":
            raise ValueError("JPEG signature does not contain a valid JPEG")
        width, height = image.size
        dpi = image.info.get("dpi", (72, 72))
        image.verify()
        return cls(width, height, int(dpi[0]) if dpi[0] > 0 else 72,
                   int(dpi[1]) if dpi[1] > 0 else 72)


if not any(signature == b"\xff\xd8\xff" for _, _, signature in docx.image.SIGNATURES):
    docx.image.SIGNATURES += ((_JpegWithoutAppHeader, 0, b"\xff\xd8\xff"),)


FORMAT_SPEC = "yama-print/1.0.7"
MIN_IMAGE_CM = 8.0
PT_TO_CM = 2.54 / 72
_YEAR = r"(?:(?:건양대)?\d{1,4}|연도미상)"
LABEL = re.compile(rf"^(?:(?:\[)?(?:\*\*)?{_YEAR}(?:[·,/]{_YEAR})*(?:년)?\s*(?:첨부(?: 이미지)?|복기 이미지)(?:\*\*)?(?:\])?|공통 사진 [^\n]+|\[Figure [^\n]+\])$")


def iter_paragraphs(parent):
    """Yield body and table-cell paragraphs in XML reading order."""
    element = getattr(parent, "element", getattr(parent, "_element", parent))
    body = getattr(element, "body", element)
    for node in body.iter(qn("w:p")):
        yield Paragraph(node, parent)


def _snapshot(directories):
    files = {}
    missing = []
    for directory in directories:
        if not directory.is_dir():
            missing.append(str(directory))
            continue
        for path in sorted(directory.rglob("*")):
            if path.is_file():
                files[str(path.resolve())] = hashlib.md5(path.read_bytes()).hexdigest()
    return {"files": files, "missing_directories": missing}


def find_panel_gap(image):
    """Return a continuous white separator with >=25% image on each side."""
    gray = image.convert("L")
    if gray.height <= gray.width:
        return None
    bands, start = [], None
    for y in range(gray.height):
        row = gray.crop((0, y, gray.width, y + 1))
        white = row.getextrema()[0] > 225 and ImageStat.Stat(row).mean[0] > 248
        if white and start is None:
            start = y
        if not white and start is not None:
            bands.append((start, y))
            start = None
    if start is not None:
        bands.append((start, gray.height))
    candidates = [(a, b) for a, b in bands if b - a >= 2
                  and a >= gray.height * .25 and gray.height - b >= gray.height * .25]
    return min(candidates, key=lambda band: abs(sum(band) / 2 - gray.height / 2), default=None)


def measure_word(path, blocks, label_rows, paragraph_names, *, export_pdf=None):
    """Measure the exact saved artifact with a separate, read-only Word instance."""
    try:
        import pythoncom
        import win32com.client
        import win32gui
        import win32process
    except ImportError as exc:
        return {"available": False, "status": "UNAVAILABLE", "error": str(exc)}
    app = document = None
    owned = False
    pythoncom.CoInitialize()
    try:
        existing_pids = set(win32process.EnumProcesses())
        app = win32com.client.DispatchEx("Word.Application")
        new_windows = []

        def collect_window(hwnd, _):
            if win32gui.GetClassName(hwnd) == "OpusApp":
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid and pid not in existing_pids:
                    new_windows.append(hwnd)

        win32gui.EnumWindows(collect_window, None)
        if len(new_windows) != 1 or win32gui.IsWindowVisible(new_windows[0]) or app.Visible:
            return {"available": False, "status": "ERROR",
                    "error": "Word automation did not start one separate hidden Word window."}
        owned = True
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3
        document = app.Documents.Open(str(Path(path).resolve()), ReadOnly=True, Visible=False,
                                      AddToRecentFiles=False, ConfirmConversions=False)
        document.Repaginate()
        if export_pdf is not None:
            document.ExportAsFixedFormat(str(export_pdf), 17)

        def location(name):
            region = document.Bookmarks(name).Range
            start = document.Range(region.Start, region.Start)
            end = document.Range(max(region.Start, region.End - 1), max(region.Start, region.End - 1))
            return {"page_first": int(start.Information(3)), "page_last": int(end.Information(3)),
                    "top_pt": float(start.Information(6)), "end_top_pt": float(end.Information(6)),
                    "lines": max(1, int(region.ComputeStatistics(1)))}

        measured = [{"key": block["key"], **location(block["bookmark"])} for block in blocks]
        labels = [[location(name) for name in names] for names in label_rows]
        label_alignment = all(len({p["page_first"] for p in row}) <= 1
                              and (not row or max(p["top_pt"] for p in row) - min(p["top_pt"] for p in row) <= .1)
                              for row in labels)
        widths = [float(document.InlineShapes(i).Width) * PT_TO_CM
                  for i in range(1, document.InlineShapes.Count + 1)]
        return {"available": True, "status": "OK", "blocks": measured,
                "spans_pages": [b["key"] for b in measured if b["page_first"] != b["page_last"]],
                "label_alignment": label_alignment, "label_rows": labels,
                "paragraphs": {name: location(name) for name in paragraph_names},
                "inline_shapes": int(document.InlineShapes.Count), "image_widths_cm": widths,
                "pages": int(document.ComputeStatistics(2))}
    except Exception as exc:
        return {"available": False, "status": "ERROR", "error": f"{type(exc).__name__}: {exc}"}
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass  # A disconnected Word instance must still produce CHECK.
        if owned and app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


class YamaLayout:
    def __init__(self, document, source_dirs=()):
        self.document = document
        self.source_dirs = sorted({Path(path).resolve() for path in source_dirs}, key=str)
        self.before = _snapshot(self.source_dirs)
        section = document.sections[0]
        self.width_cm = (section.page_width - section.left_margin - section.right_margin) / 360000
        self.height_cm = (section.page_height - section.top_margin - section.bottom_margin) / 360000
        self.blocks = []
        self.group_chains = []
        self.issues = []
        self.transformations = []
        self.label_rows = []
        self.paragraph_names = {}
        ids = [int(item.get(qn("w:id"))) for item in document.element.iter(qn("w:bookmarkStart"))]
        self.bookmark_id = max(ids, default=0)

    def add_block(self, start_index, key):
        body = self.document.element.body
        has_section = len(body) and body[-1].tag == qn("w:sectPr")
        # python-docx inserts every new node before the trailing sectPr.
        start = start_index - int(has_section)
        elements = list(body)[max(start, 0):len(body) - int(has_section)]
        if not elements:
            raise ValueError(f"Empty Yama block: {key}")
        self.blocks.append({"key": str(key), "elements": elements})

    def _paragraphs(self, elements):
        return [Paragraph(p, self.document) for node in elements for p in node.iter(qn("w:p"))]

    def _shapes(self, elements):
        return [InlineShape(inline) for node in elements for inline in node.iter(qn("wp:inline"))]

    def _font_size(self, paragraph):
        sizes = [run.font.size.pt for run in paragraph.runs if run.font.size is not None]
        style = paragraph.style
        while not sizes and style is not None:
            if style.font.size is not None:
                sizes.append(style.font.size.pt)
            style = style.base_style
        return max(sizes, default=10.5)

    def _format(self, paragraph, name, default=0):
        value = getattr(paragraph.paragraph_format, name)
        style = paragraph.style
        while value is None and style is not None:
            value = getattr(style.paragraph_format, name)
            style = style.base_style
        return default if value is None else value

    def _height(self, node, width_cm, measurements=None, text_only=False):
        if node.tag == qn("w:p"):
            paragraph = Paragraph(node, self.document)
            before = self._format(paragraph, "space_before")
            after = self._format(paragraph, "space_after")
            spacing = (getattr(before, "pt", 0) + getattr(after, "pt", 0)) * PT_TO_CM
            shapes = self._shapes([node])
            if shapes:
                return 0 if text_only else max(shape.height.cm for shape in shapes) + 2 * PT_TO_CM + spacing
            size = self._font_size(paragraph)
            line_spacing = self._format(paragraph, "line_spacing", 1.15)
            line_pt = line_spacing.pt if hasattr(line_spacing, "pt") else size * 1.305 * float(line_spacing)
            measured = (measurements or {}).get(self.paragraph_names.get(node))
            if measured:
                lines = measured["lines"]
            else:
                indents = sum(getattr(self._format(paragraph, name), "cm", 0)
                              for name in ("left_indent", "right_indent"))
                usable_pt = max(1, (width_cm - indents) / PT_TO_CM)
                lines = sum(max(1, math.ceil(sum(size * (1 if unicodedata.east_asian_width(ch) in "WF" else .55)
                                               for ch in line) / usable_pt))
                            for line in paragraph.text.split("\n"))
            return lines * line_pt * PT_TO_CM + spacing
        if node.tag == qn("w:tbl"):
            height = 0
            for row in node.findall(qn("w:tr")):
                cells = row.findall(qn("w:tc"))
                heights = [sum(self._height(child, width_cm / max(1, len(cells)) - .2,
                                            measurements, text_only)
                               for child in cell if child.tag in (qn("w:p"), qn("w:tbl")))
                           for cell in cells]
                height += max(heights, default=0)
            return height
        return 0

    def _block_height(self, block, measurements=None, text_only=False):
        return sum(self._height(node, self.width_cm, measurements, text_only) for node in block["elements"])

    def _units(self, elements):
        """Only adjacent image paragraphs and their own attachment labels qualify."""
        units, i = [], 0
        while i < len(elements):
            node = elements[i]
            if node.tag != qn("w:p"):
                i += 1
                continue
            paragraph = Paragraph(node, self.document)
            if LABEL.fullmatch(paragraph.text.strip()) and i + 1 < len(elements):
                image = elements[i + 1]
                if image.tag == qn("w:p") and not Paragraph(image, self.document).text.strip() and len(self._shapes([image])) == 1:
                    units.append((i, i + 2, [node, image]))
                    i += 2
                    continue
            if not paragraph.text.strip() and len(self._shapes([node])) == 1:
                units.append((i, i + 1, [node]))
            i += 1
        return units

    def _table(self, block, units, kind="horizontal"):
        count = len(units)
        image_width = self.width_cm / count - .2
        if image_width < MIN_IMAGE_CM:
            self.issues.append(f"{block['key']}: {count}열 가로 배치는 이미지 폭 8cm 하한을 지킬 수 없음")
            return False
        table = self.document.add_table(rows=1, cols=count)
        table.autofit = False
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement(f"w:{edge}")
            border.set(qn("w:val"), "nil")
            borders.append(border)
        table._tbl.tblPr.append(borders)
        table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        first = units[0][2][0]
        first.addprevious(table._tbl)
        labels = []
        for column, (_, _, nodes) in enumerate(units):
            cell = table.cell(0, column)
            table.columns[column].width = Cm(self.width_cm / count)
            cell.width = Cm(self.width_cm / count)
            margins = OxmlElement("w:tcMar")
            for edge in ("top", "bottom", "left", "right"):
                margin = OxmlElement(f"w:{edge}")
                margin.set(qn("w:w"), str(round(.1 / PT_TO_CM * 20)) if edge in ("left", "right") else "0")
                margin.set(qn("w:type"), "dxa")
                margins.append(margin)
            cell._tc.get_or_add_tcPr().append(margins)
            for paragraph in list(cell._tc.findall(qn("w:p"))):
                cell._tc.remove(paragraph)
            for node in nodes:
                cell._tc.append(node)
                paragraph = Paragraph(node, self.document)
                paragraph.paragraph_format.space_before = Cm(0)
                if paragraph.text.strip():
                    labels.append(node)
                for shape in self._shapes([node]):
                    target = max(MIN_IMAGE_CM, min(shape.width.cm, image_width))
                    ratio = target / shape.width.cm
                    shape.height = Cm(shape.height.cm * ratio)
                    shape.width = Cm(target)
        start, end = units[0][0], units[-1][1]
        block["elements"][start:end] = [table._tbl]
        if len(labels) >= 2:
            self.label_rows.append(labels)
        self.transformations.append({"kind": kind, "block": block["key"], "columns": count,
                                     "image_width_limit_cm": round(image_width, 3)})
        return True

    def _horizontal(self, block):
        units = self._units(block["elements"])
        groups = []
        for unit in units:
            if groups and groups[-1][-1][1] == unit[0]:
                groups[-1].append(unit)
            else:
                groups.append([unit])
        for group in reversed(groups):
            if len(group) > 1:
                self._table(block, group)

    def _split(self, block):
        shapes = self._shapes(block["elements"])
        units = self._units(block["elements"])
        if len(shapes) != 1 or len(units) != 1 or self.width_cm / 2 - .2 < MIN_IMAGE_CM:
            return
        shape = shapes[0]
        relation = shape._inline.xpath(".//a:blip")[0].get(qn("r:embed"))
        blob = self.document.part.related_parts[relation].blob
        with Image.open(io.BytesIO(blob)) as image:
            gap = find_panel_gap(image)
            if gap is None:
                return
            panels = []
            for box in ((0, 0, image.width, gap[0]), (0, gap[1], image.width, image.height)):
                jpeg = image.format == "JPEG"
                panel = io.BytesIO()
                options = {"qtables": image.quantization, "subsampling": JpegImagePlugin.get_sampling(image)} if jpeg else {}
                image.crop(box).save(panel, format="JPEG" if jpeg else "PNG", **options)
                panel.seek(0)
                panels.append(panel)
        start, end, original_nodes = units[0]
        original_image = original_nodes[-1]
        retained_starts = [deepcopy(n) for n in original_image if n.tag == qn("w:bookmarkStart")]
        retained_ends = [deepcopy(n) for n in original_image if n.tag == qn("w:bookmarkEnd")]
        nodes = list(original_nodes[:-1])
        for panel in panels:
            paragraph = self.document.add_paragraph()
            if original_image.find(qn("w:pPr")) is not None:
                paragraph._p.insert(0, deepcopy(original_image.find(qn("w:pPr"))))
            paragraph.add_run().add_picture(panel, width=Cm(MIN_IMAGE_CM))
            nodes.append(paragraph._p)
        for marker in retained_starts:
            nodes[len(original_nodes) - 1].insert(0, marker)
        for marker in retained_ends:
            nodes[-1].append(marker)
        for node in nodes:
            original_image.addprevious(node)
        original_image.getparent().remove(original_image)
        block["elements"][start:end] = nodes
        first_nodes = nodes[:-1]
        new_units = [(start, start + len(first_nodes), first_nodes),
                     (start + len(first_nodes), start + len(nodes), [nodes[-1]])]
        self._table(block, new_units, kind="split")
        self.transformations[-1].update({"source_sha256": hashlib.sha256(blob).hexdigest(),
                                        "panel_sha256": [hashlib.sha256(panel.getvalue()).hexdigest() for panel in panels],
                                        "original_inline_shapes": 1, "panel_count": 2,
                                        "gap_rows": list(gap)})

    def _shrink(self, block, measurements=None):
        text_height = self._block_height(block, measurements, text_only=True)
        image_budget = self.height_cm - text_height - 1.5
        image_height = self._block_height(block, measurements) - text_height
        if image_budget <= 0:
            self.issues.append(f"{block['key']}: 텍스트와 여유 1.5cm만으로 인쇄 가능 높이 초과")
            return
        factor = min(1, image_budget / max(image_height, .001))
        shapes = self._shapes(block["elements"])
        if not shapes:
            return
        changed = False
        for shape in shapes:
            target = max(MIN_IMAGE_CM, shape.width.cm * factor)
            if target < shape.width.cm:
                shape.height = Cm(shape.height.cm * target / shape.width.cm)
                shape.width = Cm(target)
                changed = True
        if changed:
            self.transformations.append({"kind": "resize", "block": block["key"],
                                         "text_height_cm": round(text_height, 3), "reserve_cm": 1.5,
                                         "minimum_width_cm": MIN_IMAGE_CM})
        if self._block_height(block, measurements) > self.height_cm:
            self.issues.append(f"{block['key']}: 이미지 폭 8cm 하한에서 블록이 한 페이지를 초과함")

    def _image_widths(self, node, width_cm, cap=False):
        valid = True
        if node.tag == qn("w:tbl"):
            for row in node.findall(qn("w:tr")):
                cells = row.findall(qn("w:tc"))
                for cell in cells:
                    for child in cell:
                        valid = self._image_widths(child, width_cm / max(1, len(cells)) - .2, cap) and valid
        elif node.tag == qn("w:p"):
            for shape in self._shapes([node]):
                if cap and shape.width.cm > width_cm and width_cm >= MIN_IMAGE_CM:
                    shape.height = Cm(shape.height.cm * width_cm / shape.width.cm)
                    shape.width = Cm(width_cm)
                valid = shape.width.cm <= width_cm + .001 and valid
        return valid

    def _bookmark(self, first, last, name):
        self.bookmark_id += 1
        start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
        start.set(qn("w:id"), str(self.bookmark_id))
        start.set(qn("w:name"), name)
        end.set(qn("w:id"), str(self.bookmark_id))
        first.insert(1 if first.find(qn("w:pPr")) is not None else 0, start)
        last.append(end)

    def _prepare_bookmarks(self):
        # A measured retry may replace paragraph nodes; regenerate only our own
        # bookmarks and leave any pre-existing document bookmarks untouched.
        owned = {node.get(qn("w:id")) for node in self.document.element.iter(qn("w:bookmarkStart"))
                 if node.get(qn("w:name"), "").startswith(("YL_Block_", "YL_Para_"))}
        for tag in ("w:bookmarkStart", "w:bookmarkEnd"):
            for node in list(self.document.element.iter(qn(tag))):
                if node.get(qn("w:id")) in owned:
                    node.getparent().remove(node)
        # Source ranges can be added after layout construction, including a
        # continuation between measured saves. Allocate after all live IDs.
        self.bookmark_id = max((int(node.get(qn("w:id")))
                                for node in self.document.element.iter(qn("w:bookmarkStart"))), default=0)
        self.paragraph_names = {}
        for index, block in enumerate(self.blocks):
            paragraphs = self._paragraphs(block["elements"])
            for i, paragraph in enumerate(paragraphs):
                paragraph.paragraph_format.keep_together = True
                paragraph.paragraph_format.keep_with_next = i < len(paragraphs) - 1
            block["bookmark"] = f"YL_Block_{index:05d}"
            self._bookmark(paragraphs[0]._p, paragraphs[-1]._p, block["bookmark"])
        for group in self.group_chains:
            if sum(self._block_height(block) for block in group) <= self.height_cm:
                paragraphs = [p for block in group for p in self._paragraphs(block["elements"])]
                for index, paragraph in enumerate(paragraphs):
                    paragraph.paragraph_format.keep_with_next = index < len(paragraphs) - 1
        for index, paragraph in enumerate(iter_paragraphs(self.document)):
            name = f"YL_Para_{index:06d}"
            self.paragraph_names[paragraph._p] = name
            self._bookmark(paragraph._p, paragraph._p, name)
        return [[self.paragraph_names[node] for node in row] for row in self.label_rows]

    def save(self, out_path):
        path = Path(out_path).resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        expected_text = [p.text for p in iter_paragraphs(self.document) if p.text]
        original_images = len(self.document.inline_shapes)
        for block in self.blocks:
            self._horizontal(block)
            for node in block["elements"]:
                self._image_widths(node, self.width_cm, cap=True)
            if self._block_height(block) > self.height_cm:
                self._split(block)
            if self._block_height(block) > self.height_cm:
                self._shrink(block)
        label_names = self._prepare_bookmarks()
        self.document.save(path)
        word = measure_word(path, self.blocks, label_names, list(self.paragraph_names.values()))
        attempts = 1
        if word.get("available") and word.get("status") == "OK":
            measurements = word.get("paragraphs", {})
            overflow = [block for block in self.blocks if self._block_height(block, measurements) > self.height_cm
                        or block["key"] in word.get("spans_pages", [])]
            if overflow:
                for block in overflow:
                    previous = len(self.transformations)
                    self._split(block)
                    if self._block_height(block, measurements) > self.height_cm or len(self.transformations) == previous:
                        self._shrink(block, measurements)
                label_names = self._prepare_bookmarks()
                self.document.save(path)
                word = measure_word(path, self.blocks, label_names, list(self.paragraph_names.values()))
                attempts += 1
        reloaded = Document(path)
        actual_text = [p.text for p in iter_paragraphs(reloaded) if p.text]
        word.setdefault("status", "OK" if word.get("available") else "UNAVAILABLE")
        available = word.get("available", False) and word["status"] == "OK"
        measured_paragraphs = word.get("paragraphs", {}) if available else {}
        calculated = [{"key": b["key"], "height_cm": round(self._block_height(b, measured_paragraphs), 4),
                       "height_basis": "Word line count" if available else "estimated line count"}
                      for b in self.blocks]
        expected_images = original_images + sum(t.get("panel_count", 0) - t.get("original_inline_shapes", 0)
                                                for t in self.transformations if t["kind"] == "split")
        after = _snapshot(self.source_dirs)
        before_files, after_files = self.before["files"], after["files"]
        # The caller may write beside the source images. Only these two declared
        # outputs are outside the input-integrity check; all other files count.
        outputs = {str(path), str(path.with_suffix(".layout-qc.json"))}
        before_files = {name: value for name, value in before_files.items() if name not in outputs}
        after_files = {name: value for name, value in after_files.items() if name not in outputs}
        integrity = {"before_md5": before_files, "after_md5": after_files,
                     "added": sorted(after_files.keys() - before_files.keys()),
                     "removed": sorted(before_files.keys() - after_files.keys()),
                     "changed": sorted(name for name in before_files.keys() & after_files.keys()
                                       if before_files[name] != after_files[name]),
                     "missing_directories": sorted(set(self.before["missing_directories"] + after["missing_directories"]))}
        bookmark_names = {node.get(qn("w:name")) for node in reloaded.element.iter(qn("w:bookmarkStart"))}
        widths = [shape.width.cm for shape in reloaded.inline_shapes]
        checks = {
            "block_height": all(b["height_cm"] <= self.height_cm for b in calculated),
            "actual_pagination": available and not word.get("spans_pages"),
            "image_count": len(reloaded.inline_shapes) == expected_images and (not available or word["inline_shapes"] == expected_images),
            "minimum_image_width": all(width >= MIN_IMAGE_CM - .001 for width in widths)
                                   and (not available or all(width >= MIN_IMAGE_CM - .001 for width in word["image_widths_cm"])),
            "maximum_image_width": all(self._image_widths(node, self.width_cm) for b in self.blocks for node in b["elements"]),
            "label_alignment": available and word.get("label_alignment", False),
            "text_integrity": actual_text == expected_text,
            "question_count": len(self.blocks) > 0 and all(b["bookmark"] in bookmark_names for b in self.blocks)
                              and (not available or len(word["blocks"]) == len(self.blocks)),
            "source_integrity": not any(integrity[key] for key in ("added", "removed", "changed", "missing_directories")),
        }
        if not available:
            self.issues.append(f"Word COM 실제 페이지/라벨 검증 불가: {word.get('error', word['status'])}")
        for name, passed in checks.items():
            if not passed:
                self.issues.append(f"QC 검사 실패: {name}")
        report = {"format_spec": FORMAT_SPEC, "verdict": "OK" if all(checks.values()) and not self.issues else "CHECK",
                  "document": str(path), "document_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                  "checks": checks, "issues": self.issues, "blocks": calculated, "word": word,
                  "measurement_attempts": attempts,
                  "printable_width_cm": self.width_cm, "printable_height_cm": self.height_cm,
                  "transformations": self.transformations, "source_integrity": integrity,
                  "expected_inline_shapes": expected_images, "inline_shapes": len(reloaded.inline_shapes),
                  "image_widths_cm": widths, "expected_question_count": len(self.blocks)}
        path.with_suffix(".layout-qc.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        return report
