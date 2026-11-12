"""Return focused local lecture and Script excerpts for one question."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import sys

SHARED = Path(__file__).resolve().parents[2] / "yama-word-obsidian" / "scripts"
sys.path.insert(0, str(SHARED))
from question_index import EMBED, normalize, read_json, stamp, verify_stamp, write_index

PART = re.compile(r"^\s*#{1,6}\s+(?:PART|PT\.?)\s*([12])\b", re.I)
RAW_PAGE = re.compile(r"^\s*<!--\s*page\s+(\d+)\s*-->\s*$", re.I)
SCRIPT_PAGE = re.compile(r"^\s*(?:[-#>*]+\s*)?\[\s*Page\s+(\d+)(\s*[?？])?\s*\]", re.I)
TIME = re.compile(r"^\s*(?:\[)?((?:\d{1,2}:)?\d{1,2}:\d{2}(?:[.,]\d{1,3})?)(?:\]|\s|$)")
TEXT_LIMIT = 800


def row(number: int, page: str | None, text: str, timestamp: str | None = None) -> dict:
    return {"line": number, "page": page, "text": text, **({"timestamp": timestamp} if timestamp else {})}


def text_rows(lines: list[str], role: str) -> list[dict]:
    part = 1
    has_part_one = role == "script" and any((match := PART.match(line)) and match[1] == "1"
                                          for line in lines)
    if has_part_one:
        part = 0
    page = timestamp = None
    rows = []
    for number, line in enumerate(lines, 1):
        part_tag = PART.match(line) if role == "script" else None
        if part_tag:
            part = int(part_tag[1])
            page = timestamp = None
            continue
        if role == "script" and part != 1:
            continue
        page_tag = SCRIPT_PAGE.match(line) if role == "script" else RAW_PAGE.match(line)
        if page_tag:
            page = page_tag[1] + ("?" if role == "script" and page_tag[2] else "")
            if role != "script" or not line[page_tag.end():].strip():
                continue
        if role == "script" and (time_tag := TIME.match(line)):
            timestamp = time_tag[1]
        if line.strip():
            rows.append(row(number, page, line, timestamp))
    return rows


def indexed_document(path: Path, role: str) -> dict:
    source = stamp(path)
    suffix = path.suffix.lower()
    lines = []
    if suffix in (".md", ".txt"):
        lines = path.read_text(encoding="utf-8-sig").splitlines()
        rows = text_rows(lines, role)
    elif suffix == ".pdf" and role == "lecture":
        import pymupdf

        with pymupdf.open(path) as doc:
            rows = [row(i, str(page.number + 1), line)
                    for page in doc for i, line in enumerate(page.get_text().splitlines(), 1) if line.strip()]
    elif suffix == ".docx":
        from docx import Document
        from docx.oxml.ns import qn

        doc = Document(path)
        lines = []
        for paragraph in doc.element.body.iter(qn("w:p")):
            lines.append("".join(t.text or "" for t in paragraph.iter(qn("w:t"))))
        rows = text_rows(lines, role)
    else:
        raise ValueError(f"{role} 자료의 형식을 지원하지 않습니다. {path}")
    verify_stamp(source)
    parts = {tag[1] for line in lines if (tag := PART.match(line))} if role == "script" else set()
    return {"role": role, "source": source, "rows": rows,
            **({"summary_only": True} if parts == {"2"} else {})}


def build_index(lecture: list[Path], script: list[Path]) -> dict:
    if not lecture and not script:
        raise ValueError("강의록이나 Script 파일을 지정하세요.")
    return {"documents": [indexed_document(path, role)
                          for role, paths in (("lecture", lecture), ("script", script)) for path in paths]}


def location(source: Path, item: dict) -> str:
    if source.suffix.lower() == ".pdf":
        return f"PDF p.{item['page']} 텍스트 줄 {item['line']}"
    where = f"Word 본문 문단 {item['line']}" if source.suffix.lower() == ".docx" else f"줄 {item['line']}"
    if item.get("timestamp"):
        where = f"{item['timestamp']} / {where}"
    if item.get("page"):
        where = f"p.{item['page']} / {where}"
    return where


def excerpt(item: dict, terms: list[str]) -> tuple[str, bool]:
    text = item["text"]
    if len(text) <= TEXT_LIMIT:
        return text, False
    starts = [m.start() for term in terms if (m := re.search(re.escape(term), text, re.I))]
    start = max(0, min(starts, default=0) - 300)
    return text[start:start + TEXT_LIMIT], True


def packet(doc: dict, source: Path, focus: int, context: int, terms: list[str]) -> dict:
    rows = doc["rows"]
    item = rows[focus]
    first, last = max(0, focus - context), min(len(rows), focus + context + 1)
    selected = []
    for i in range(first, last):
        evidence = rows[i]
        text, clipped = excerpt(evidence, terms if i == focus else [])
        selected.append({"location": location(source, evidence), "text": text,
                         **({"clipped": True, "notice": "발췌 일부만 표시했습니다."} if clipped else {})})
    figures = []
    if item.get("page") is not None:
        for candidate in rows:
            if candidate["page"] == item["page"]:
                for match in EMBED.finditer(candidate["text"]):
                    image = source.parent / match[1]
                    if not image.is_file() and len(Path(match[1]).parts) == 1:
                        image = source.parent / "동봉 이미지" / match[1]
                    figures.append({"reference": match[1], "location": location(source, candidate),
                                    "file": str(image.resolve()) if image.is_file() else None})
                    if len(figures) == 6:
                        break
            if len(figures) == 6:
                break
    return {"location": location(source, item), "line": item["line"], "page": item.get("page"),
            "excerpts": selected, **({"figures": figures} if figures else {})}


def find_evidence(index: dict, *, terms: list[str], pages: list[str], at: int | None,
                  scope: str | None, context: int, limit: int) -> dict:
    if not (terms or pages or at is not None):
        raise ValueError("검색어, 페이지 또는 줄 번호를 지정하세요.")
    if at is not None and not scope:
        raise ValueError("줄 위치를 열 때 --file로 자료 하나를 선택하세요.")
    needles = [normalize(term) for term in terms]
    scoped = [doc for doc in index["documents"]
              if scope is None or normalize(scope) in normalize(doc["source"]["path"])]
    if scope and len(scoped) != 1:
        raise ValueError("--file에 인덱스의 고유 파일명 또는 전체 경로를 지정하세요.")
    if at is not None and not pages and Path(scoped[0]["source"]["path"]).suffix.lower() == ".pdf":
        raise ValueError("PDF 텍스트 줄을 열 때 --page로 PDF 페이지를 지정하세요.")
    output = {}
    for role in ("lecture", "script"):
        sources, ranked = [], []
        for doc in scoped:
            if doc["role"] != role:
                continue
            path = verify_stamp(doc["source"])
            sources.append({"file": str(path), "text_rows": len(doc["rows"]),
                            **({"summary_only": True} if doc.get("summary_only") else {})})
            for i, item in enumerate(doc["rows"]):
                page_matches = item["page"] and item["page"].rstrip("?") in pages
                if pages and not page_matches:
                    continue
                if at is not None and item["line"] != at:
                    continue
                score = sum(needle in normalize(item["text"]) for needle in needles)
                if needles and not score:
                    continue
                ranked.append((-score, len(sources), i, doc, path))
        ranked.sort(key=lambda hit: hit[:3])
        selected = []
        for _, doc_position, i, doc, path in ranked:
            if any(p == doc_position and abs(i - other) <= context for p, other, _ in selected):
                continue
            selected.append((doc_position, i, {"file": str(path), **packet(doc, path, i, context, terms)}))
            if len(selected) == limit:
                break
        if selected:
            status = "found"
        elif sources:
            status = "no_match"
        elif any(doc["role"] == role for doc in index["documents"]):
            status = "out_of_scope"
        else:
            status = "not_provided"
        output[role] = {"status": status, "sources": sources, "matched_rows": len(ranked),
                        "matches_limited": len(ranked) > len(selected),
                        "results": [result for _, _, result in selected]}
    return output


def main(argv=None) -> int:
    cli = argparse.ArgumentParser(description=__doc__)
    commands = cli.add_subparsers(dest="command", required=True)
    build = commands.add_parser("build")
    for option in ("lecture", "script"):
        build.add_argument(f"--{option}", action="append", type=Path, default=[])
    build.add_argument("--out", type=Path, required=True)
    search = commands.add_parser("search")
    search.add_argument("--index", type=Path, required=True)
    search.add_argument("--term", action="append", default=[])
    search.add_argument("--page", action="append", default=[])
    search.add_argument("--at", type=int)
    search.add_argument("--file")
    search.add_argument("--context", type=int, choices=range(0, 5), default=1)
    search.add_argument("--limit", type=int, choices=range(1, 5), default=2)
    args = cli.parse_args(argv)
    try:
        if args.command == "build":
            if any(args.out.resolve() == path.resolve() for path in args.lecture + args.script):
                raise ValueError("출력 인덱스에 원본 경로를 사용할 수 없습니다.")
            index = build_index(args.lecture, args.script)
            write_index(args.out, index, "evidence", [doc["source"] for doc in index["documents"]])
            result = {"index": str(args.out.resolve()),
                      "sources": [{"role": doc["role"], "file": doc["source"]["path"], "rows": len(doc["rows"])}
                                  for doc in index["documents"]]}
        else:
            if any(not normalize(term) for term in args.term):
                raise ValueError("빈 검색어는 사용할 수 없습니다.")
            result = find_evidence(read_json(args.index), terms=args.term, pages=args.page,
                                   at=args.at, scope=args.file, context=args.context, limit=args.limit)
    except (ImportError, OSError, ValueError) as exc:
        print(f"CHECK — {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
