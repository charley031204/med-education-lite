#!/usr/bin/env python3
"""Prepare native RAW Markdown drafts, render evidence, and check document structure."""

import argparse
from datetime import date
from io import BytesIO
import json
import math
from pathlib import Path, PurePosixPath
import re
import sys
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from content import box_text, literal


CHUNK_LIMIT = 15
ASSET_FOLDER = "동봉 이미지"
MARKER = re.compile(r"^<!-- page ([1-9][0-9]*) -->$", re.MULTILINE)
PAGE_LIKE = re.compile(r"^[ \t]*<!--[ \t]*(?:page[ \t]*\d+|p\.[ \t]*\d+).*?-->[ \t]*$",
                       re.MULTILINE | re.IGNORECASE)
FIGURE = re.compile(r"^> \*\*\[Figure ([1-9][0-9]*)[a-z]*(?: / [^\n]+)?\]\*\*[ \t]*$",
                    re.MULTILINE)
EMBED = re.compile(r"!\[\[([^\]\n]+)\]\]")
ASSET_NAME = re.compile(r"(?P<slug>[\w-]+)_p(?P<page>[0-9]{3,})(?:[a-z]+|_full)\.png")
TABLE_DIVIDER = re.compile(r"^\|(?:\s*:?-{3,}:?\s*\|)+$", re.MULTILINE)


def source_type(source):
    kind = source.suffix.lower()
    if kind not in (".pdf", ".pptx"):
        raise ValueError("Expected a .pdf or .pptx input")
    return kind


def page_count(source):
    if source_type(source) == ".pdf":
        import pymupdf

        with pymupdf.open(source) as document:
            return len(document)
    with ZipFile(source) as package:
        root = ElementTree.fromstring(package.read("_rels/.rels"))
        target = next((relation.get("Target") for relation in root
                       if relation.get("Type", "").endswith("/officeDocument")
                       and relation.get("TargetMode") != "External"), None)
        if not target:
            raise ValueError("Presentation relationship is missing")
        name = str(PurePosixPath(target.lstrip("/")))
        presentation = ElementTree.fromstring(package.read(name))
        slides = presentation.find("{*}sldIdLst")
        return len(slides) if slides is not None else 0


def select_pages(total, spec, limit=None):
    if total < 1:
        raise ValueError("Input has no pages")
    numbers = set()
    if spec is None:
        numbers.update(range(1, total + 1))
    else:
        for piece in spec.split(","):
            if not re.fullmatch(r"[0-9]+(?:\s*-\s*[0-9]+)?", piece.strip()):
                raise ValueError("Expected page numbers or inclusive ranges, e.g. 1-5,7")
            parts = piece.split("-")
            start, end = int(parts[0]), int(parts[-1])
            if not 1 <= start <= end <= total:
                raise ValueError(f"Page range {piece.strip()} is outside 1..{total} or reversed")
            numbers.update(range(start, end + 1))
    if limit is not None and len(numbers) > limit:
        raise ValueError(f"Select at most {limit} pages")
    return sorted(numbers)


def new_paths(*paths):
    resolved = [path.resolve() for path in paths]
    outputs = set(resolved)
    if len(outputs) != len(resolved) or any(outputs.intersection(path.parents) for path in resolved):
        raise ValueError("Outputs must use distinct paths, with no output inside another")
    for path in paths:
        if path.exists() or path.is_symlink():
            raise FileExistsError(f"Output already exists: {path}")
        if any(parent.exists() and not parent.is_dir() for parent in path.parents):
            raise ValueError(f"An output parent is not a directory: {path}")


def write_text(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8", newline="\n") as handle:
        handle.write(data)


def write_bytes(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as handle:
        handle.write(data)


def print_json(data):
    print(json.dumps(data, ensure_ascii=False))


def plan(args):
    if not 1 <= args.size <= CHUNK_LIMIT:
        raise ValueError(f"Chunk size must be 1..{CHUNK_LIMIT}")
    total = page_count(args.source)
    numbers = select_pages(total, args.pages)
    chunks = []
    for index, start in enumerate(range(0, len(numbers), args.size), 1):
        pages = numbers[start:start + args.size]
        chunks.append({"chunk": f"{index:02}", "pages": pages, "page_spec": ",".join(map(str, pages))})
    print_json({"source": str(args.source.resolve()), "total_pages": total, "chunks": chunks})


def native_pages(source, numbers, *, details, notes=False):
    if source_type(source) == ".pdf":
        import pymupdf
        from pdf_native import read_pdf_page

        with pymupdf.open(source) as document:
            if document.needs_pass:
                raise ValueError("The PDF requires a password")
            for number in numbers:
                yield read_pdf_page(document[number - 1], details)
    else:
        from pptx import Presentation
        from pptx_native import read_slide

        deck = Presentation(source)
        for number in numbers:
            yield read_slide(deck, number, details, notes)


def packet_text(pages):
    body = "\n\n".join(page.packet() for page in pages)
    return "NATIVE EXTRACTION — UNREVIEWED\n\n" + body + "\n"


def draft(args):
    outputs = [args.out] + ([args.packet] if args.packet else [])
    new_paths(*outputs)
    if args.notes and not args.packet:
        raise ValueError("Use --notes with --packet")
    numbers = select_pages(page_count(args.source), args.pages, CHUNK_LIMIT)
    pages = list(native_pages(args.source, numbers, details=bool(args.packet), notes=args.notes))
    write_text(args.out, "\n\n".join(page.draft() for page in pages) + "\n")
    if args.packet:
        write_text(args.packet, packet_text(pages))
    print_json({"status": "UNREVIEWED", "draft": str(args.out.resolve()), "pages": numbers,
                "packet": str(args.packet.resolve()) if args.packet else None,
                "tables": sum(len(page.tables) for page in pages),
                "hidden_pages": [page.number for page in pages if page.hidden],
                "review": {page.number: sorted(page.warnings) for page in pages if page.warnings}})


def packet(args):
    numbers = select_pages(page_count(args.source), args.pages, CHUNK_LIMIT)
    print(packet_text(native_pages(args.source, numbers, details=True, notes=args.notes)), end="")


def coordinates(value, *, integers=False):
    try:
        coords = [int(piece) if integers else float(piece) for piece in value.split(",")]
    except ValueError as error:
        raise ValueError("Expected four numeric coordinates: x0,y0,x1,y1") from error
    if len(coords) != 4 or not all(math.isfinite(coord) for coord in coords):
        raise ValueError("Expected four finite coordinates: x0,y0,x1,y1")
    x0, y0, x1, y1 = coords
    if x0 < 0 or y0 < 0 or x1 <= x0 or y1 <= y0:
        raise ValueError("Crop bounds must have positive area within the page")
    return tuple(coords)


def pdf_numbers(args):
    if source_type(args.source) != ".pdf":
        raise ValueError("This command requires PDF input")
    if not 72 <= args.dpi <= 600:
        raise ValueError("DPI must be 72..600")
    return select_pages(page_count(args.source), args.pages, CHUNK_LIMIT)


def render(args):
    import pymupdf

    numbers = pdf_numbers(args)
    if args.box and len(numbers) != 1:
        raise ValueError("Use --box with one page")
    paths = [args.out_dir / f"page_{number:03}.png" for number in numbers]
    new_paths(*paths)
    with pymupdf.open(args.source) as document:
        for number, path in zip(numbers, paths, strict=True):
            page = document[number - 1]
            box = pymupdf.Rect(coordinates(args.box)) if args.box else page.rect
            if not page.rect.contains(box):
                raise ValueError("Crop bounds extend beyond the displayed page")
            write_bytes(path, page.get_pixmap(dpi=args.dpi, clip=box).tobytes("png"))
    print_json({"pages": numbers, "renders": [str(path.resolve()) for path in paths]})


def figure_label(index):
    label = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        label = chr(ord("a") + remainder) + label
    return label


def pdf_figures(args):
    import pymupdf
    from pdf_native import figure_candidates

    numbers = pdf_numbers(args)
    with pymupdf.open(args.source) as document:
        candidates = []
        for number in numbers:
            for index, box in enumerate(figure_candidates(document[number - 1]), 1):
                path = args.out_dir / f"page_{number:03}_{figure_label(index)}.png"
                candidates.append((number, path, box))
        new_paths(*(path for _, path, _ in candidates))
        for number, path, box in candidates:
            page = document[number - 1]
            write_bytes(path, page.get_pixmap(dpi=args.dpi, clip=box).tobytes("png"))
    return {"status": "UNREVIEWED", "pages": numbers,
            "candidates": [{"page": number, "file": str(path.resolve()), "display_pt": box_text(box)}
                           for number, path, box in candidates]}


def pptx_figures(args):
    from PIL import Image
    from pptx import Presentation
    from pptx_native import tag, walk_shapes

    deck = Presentation(args.source)
    numbers = select_pages(len(deck.slides), args.pages, CHUNK_LIMIT)
    candidates = []
    for number in numbers:
        pictures = [shape for shape, _ in walk_shapes(deck.slides[number - 1].shapes)
                    if tag(shape._element) == "pic"]
        for index, shape in enumerate(pictures, 1):
            path = args.out_dir / f"page_{number:03}_{figure_label(index)}.png"
            candidates.append((number, path, shape))
    new_paths(*(path for _, path, _ in candidates))
    assets = []
    unresolved = set()
    for number, path, shape in candidates:
        try:
            with Image.open(BytesIO(shape.image.blob)) as image:
                image = image.convert("RGBA")
        except (AttributeError, KeyError, OSError, ValueError):
            unresolved.add(number)
            continue
        output = BytesIO()
        image.save(output, format="PNG")
        write_bytes(path, output.getvalue())
        assets.append({"page": number, "file": str(path.resolve())})
    return {"status": "UNREVIEWED", "pages": numbers, "candidates": assets,
            "unresolved_picture_pages": sorted(unresolved),
            "review": "Original picture pixels. Compare with the displayed crop and composition."}


def figures(args):
    if source_type(args.source) == ".pdf":
        print_json(pdf_figures(args))
    else:
        print_json(pptx_figures(args))


def crop(args):
    from PIL import Image

    new_paths(args.out)
    box = coordinates(args.box, integers=True)
    with Image.open(args.image) as image:
        if image.format != "PNG":
            raise ValueError("Expected a PNG render")
        if box[2] > image.width or box[3] > image.height:
            raise ValueError("Crop bounds extend beyond the PNG")
        cropped = image.crop(box)
        output = BytesIO()
        cropped.save(output, format="PNG")
    write_bytes(args.out, output.getvalue())
    print_json({"image": str(args.out.resolve()), "width": cropped.width, "height": cropped.height})


def find_asset(link, assets, page, slug):
    from PIL import Image

    target, divider, width = link.partition("|")
    prefix = ASSET_FOLDER + "/"
    if not target.startswith(prefix) or divider != "|" or width != "600":
        return None, "invalid_embed_format"
    filename = target[len(prefix):]
    match = ASSET_NAME.fullmatch(filename)
    if not match:
        return None, "invalid_asset_filename"
    if match["page"] != f"{page:03}" or (slug is not None and match["slug"] != slug):
        return None, "wrong_asset_page_or_slug"
    path = assets / filename
    if path.resolve().parent != assets.resolve() or not path.is_file():
        return None, "missing_or_outside_asset"
    try:
        with Image.open(path) as image:
            if image.format != "PNG":
                return None, "invalid_png"
            image.verify()
    except (OSError, SyntaxError, ValueError):
        return None, "invalid_png"
    return filename, None


def mechanical_check(raw, expected, assets, slug=None):
    markers = list(MARKER.finditer(raw))
    numbers = [int(marker[1]) for marker in markers]
    errors = []
    warnings = []
    if numbers != expected:
        errors.append({"type": "page_order_or_coverage", "expected": expected, "actual": numbers})
    malformed = sum(not MARKER.fullmatch(match[0].strip("\n\r")) for match in PAGE_LIKE.finditer(raw))
    if malformed:
        errors.append({"type": "malformed_page_marker", "count": malformed})
    if markers and EMBED.search(raw[:markers[0].start()]):
        errors.append({"type": "embed_before_first_page"})
    table_count = figure_count = embed_count = 0
    used = set()
    for index, marker in enumerate(markers):
        number = int(marker[1])
        stop = markers[index + 1].start() if index + 1 < len(markers) else len(raw)
        body = raw[marker.end():stop]
        if not body.strip():
            errors.append({"type": "empty_page_block", "page": number})
        figures = list(FIGURE.finditer(body))
        embeds = list(EMBED.finditer(body))
        table_count += len(TABLE_DIVIDER.findall(body))
        figure_count += len(figures)
        embed_count += len(embeds)
        missing_figures = 0
        for figure_index, figure in enumerate(figures):
            if int(figure[1]) != number:
                errors.append({"type": "wrong_figure_page", "page": number})
            figure_end = figures[figure_index + 1].start() if figure_index + 1 < len(figures) else len(body)
            matches = [embed for embed in embeds if figure.end() <= embed.start() < figure_end]
            if not matches:
                missing_figures += 1
            if len(matches) > 1:
                errors.append({"type": "multiple_embeds_for_figure", "page": number})
        if missing_figures:
            warnings.append({"type": "figure_without_embed", "page": number, "count": missing_figures})
        for embed in embeds:
            if not figures or embed.start() < figures[0].end():
                errors.append({"type": "embed_without_figure", "page": number})
            line_start = body.rfind("\n", 0, embed.start()) + 1
            line_end = body.find("\n", embed.end())
            line_end = len(body) if line_end < 0 else line_end
            if body[line_start:line_end] != embed[0]:
                errors.append({"type": "embed_must_be_on_its_own_line", "page": number})
            filename, error = find_asset(embed[1], assets, number, slug)
            if error:
                errors.append({"type": error, "page": number})
            elif filename in used:
                errors.append({"type": "duplicate_embed", "page": number})
            else:
                used.add(filename)
    if slug and assets.is_dir():
        unused = []
        expected_pages = set(expected)
        for path in assets.iterdir():
            match = ASSET_NAME.fullmatch(path.name)
            if match and match["slug"] == slug and int(match["page"]) in expected_pages and path.name not in used:
                unused.append(path.name)
        if unused:
            warnings.append({"type": "unused_assets", "files": sorted(unused)})
    return {"mechanical_only": True, "ok": not errors, "pages": numbers, "tables": table_count,
            "figures": figure_count, "embeds": embed_count, "errors": errors, "warnings": warnings}


def check_slug(slug):
    if slug is not None and not re.fullmatch(r"[\w-]+", slug):
        raise ValueError("Use only letters, digits, underscores and hyphens in --slug")


def validate(args):
    check_slug(args.slug)
    numbers = select_pages(page_count(args.source), args.pages)
    assets = args.assets or args.raw.parent / ASSET_FOLDER
    summary = mechanical_check(args.raw.read_text(encoding="utf-8"), numbers, assets, args.slug)
    print_json(summary)
    return 0 if summary["ok"] else 1


def merge(args):
    check_slug(args.slug)
    new_paths(args.out)
    numbers = select_pages(page_count(args.source), args.pages)
    chunks = []
    for path in args.chunk:
        body = path.read_text(encoding="utf-8")
        first = MARKER.match(body)
        if first is None or len(MARKER.findall(body)) > CHUNK_LIMIT:
            raise ValueError(f"Chunk must begin at a page marker and have at most {CHUNK_LIMIT} pages: {path}")
        chunks.append((int(first[1]), body.rstrip("\n")))
    chunks.sort(key=lambda chunk: chunk[0])
    title = args.title or args.source.stem
    if any(character in title + args.source.name for character in "\r\n"):
        raise ValueError("Title and source filename must fit on one line")
    conversion_date = date.fromisoformat(args.date) if args.date else date.today()
    header = (f"# {literal(title)} — RAW MD\n"
              f"> 변환 일시: {conversion_date.isoformat()}\n"
              f"> 원본 파일: {literal(args.source.name)}\n"
              f"> 총 페이지: {len(numbers)}\n"
              f"> 변환 방식: Subagent 분할 — {len(chunks)} Chunk\n\n---\n\n")
    raw = header + "\n\n".join(body for _, body in chunks) + "\n"
    summary = mechanical_check(raw, numbers, args.out.parent / ASSET_FOLDER, args.slug)
    if summary["ok"]:
        write_text(args.out, raw)
        summary["output"] = str(args.out.resolve())
    print_json(summary)
    return 0 if summary["ok"] else 1


def build_parser():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)

    def source_command(function, *, require_pages=False):
        command = commands.add_parser(function.__name__, description=function.__doc__)
        command.add_argument("source", type=Path)
        command.add_argument("--pages", required=require_pages, help="Original page numbers, e.g. 1-5,7")
        command.set_defaults(handler=function)
        return command

    source_command(plan).add_argument("--size", type=int, default=CHUNK_LIMIT)
    draft_args = source_command(draft, require_pages=True)
    draft_args.add_argument("--out", required=True, type=Path)
    draft_args.add_argument("--packet", type=Path)
    draft_args.add_argument("--notes", action="store_true", help="Include PPTX speaker notes in the packet")
    source_command(packet, require_pages=True).add_argument(
        "--notes", action="store_true", help="Include PPTX speaker notes")
    for function in (render, figures):
        render_args = source_command(function, require_pages=True)
        render_args.add_argument("--out-dir", required=True, type=Path)
        render_args.add_argument("--dpi", type=int, default=160, help="PDF render resolution")
        if function is render:
            render_args.add_argument("--box", help="Single page crop in displayed-page points")
    crop_args = commands.add_parser("crop")
    crop_args.add_argument("image", type=Path)
    crop_args.add_argument("--box", required=True, help="Crop in PNG pixels: x0,y0,x1,y1")
    crop_args.add_argument("--out", required=True, type=Path)
    crop_args.set_defaults(handler=crop)
    merge_args = source_command(merge)
    merge_args.add_argument("--chunk", action="append", required=True, type=Path)
    merge_args.add_argument("--out", required=True, type=Path)
    merge_args.add_argument("--title")
    merge_args.add_argument("--date")
    merge_args.add_argument("--slug")
    validate_args = source_command(validate)
    validate_args.add_argument("raw", type=Path)
    validate_args.add_argument("--assets", type=Path)
    validate_args.add_argument("--slug")
    return parser


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    args = build_parser().parse_args()
    try:
        return args.handler(args) or 0
    except (OSError, ValueError, RuntimeError, BadZipFile, ElementTree.ParseError, KeyError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
