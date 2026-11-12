"""Index local question text and image references without extracting every figure."""

from __future__ import annotations

import argparse
from collections import defaultdict, deque
import hashlib
import json
from pathlib import Path
import re
import sys
import unicodedata
from zipfile import BadZipFile, ZipFile

YEAR_ATOM = r"(?:(?:건양대)?\d{1,2}|연도미상)"
YEAR_PREFIX = re.compile(rf"^\[\s*({YEAR_ATOM}(?:\s*,\s*(?:{YEAR_ATOM}))*)\s*\]\s*(.*)$")
HISTORY = re.compile(r"^※\s*출제 이력:\s*(.*)$")
HIST_ITEM = re.compile(
    r"^\s*(?P<year>(?:건양대)?\d{2,4}|연도미상)(?:년)?\s+"
    r"(?:(?P<round>.+?)\s+)?(?P<number>주관식\s*\d+|\d+(?:\s*[~\-–]\s*\d+)?(?:외)?|무번호)(?:번)?\s*$"
)
HISTORY_YEAR = re.compile(r"(?:건양대)?\d{2,4}(?=년)|연도미상")
EMBED = re.compile(r"!\[\[([^|\]]+)(?:\|[^\]]*)?\]\]")
CHOICE = re.compile(
    r"^\s*(?:보\s*기\s*[:：]\s*)?([①-⑳]|[가나다라마바사아자차카타파하ㄱㄴㄷㄹㅁㅂㅅ][.)]"
    r"|\(\d{1,2}\)|\d{1,2}[.)])"
)
EDITOR = re.compile(r"^\[\s*(정리자|복기자|편집자)\s*(comment|코멘트|註|주|note)\s*\]$", re.I)
QUESTION_HEADER = re.compile(r"^\[!question\][+-]?\s+(?:\[(객관식|단답형|보기형|서술형|매칭형)\]\s*)?(.*)$")
PAGE_SUFFIX = re.compile(r"\(p\.(\d+(?:\s*(?:,|~)\s*\d+)*)\)")
QID_TAG = re.compile(r"^<!-- yama-qid: ([\w:@.-]+) -->$")
PACKAGE_PART = "customXml/yama-shared-photos.json"
TYPES = ("객관식", "단답형", "보기형", "서술형", "매칭형")
WORD_SCRIPTS = Path(__file__).resolve().parents[2] / "yama-chapter-word" / "scripts"


def stamp(path: Path) -> dict:
    stat = path.stat()
    return {"path": str(path.resolve()), "size": stat.st_size, "mtime_ns": stat.st_mtime_ns}


def verify_stamp(record: dict) -> Path:
    path = Path(record["path"])
    if record != stamp(path):
        raise ValueError(f"원본이 변경되었습니다. 인덱스를 다시 만드세요. {path}")
    return path


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_index(path: Path, content: dict, kind: str, sources: list[dict]) -> None:
    if path.suffix.lower() != ".json":
        raise ValueError("인덱스는 별도의 .json 파일로 저장하세요.")
    owners = [source["path"] for source in sources]
    if path.exists():
        existing = read_json(path)
        if existing.get("index_kind") != kind or existing.get("index_sources") != owners:
            raise ValueError(f"기존 파일에 인덱스를 쓸 수 없습니다. {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    value = {**content, "index_kind": kind, "index_sources": owners}
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def normalize(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", value).casefold().split())


def year_label(value: int | str) -> str:
    if isinstance(value, int):
        if value == 0:
            return "연도미상"
        return f"건양대{value + 1000:02d}" if value < 0 else f"{value:02d}"
    return value.strip()[-2:] if re.fullmatch(r"\d{2,4}", value.strip()) else value.strip()


def history_items(text: str) -> list[dict]:
    result = []
    for item in re.split(r"\s*[·,]\s*", text):
        match = HIST_ITEM.fullmatch(item)
        if match:
            result.append({**match.groupdict(), "year": year_label(match["year"])})
    return result


def text_item(text: str) -> dict:
    return {"kind": "text", "text": text}


def markdown_items(line: str) -> list[dict]:
    result = []
    position = 0
    for match in EMBED.finditer(line):
        if match.start() > position:
            result.append(text_item(line[position:match.start()]))
        result.append({"kind": "image", "path": match[1]})
        position = match.end()
    if position < len(line):
        result.append(text_item(line[position:]))
    return result


def question_text(question: dict) -> str:
    stem = question["stem"]
    if stem:
        return stem
    return " ".join(i["text"] for i in question["body"] if i["kind"] == "text")


def choice_markers(question: dict) -> list[str]:
    markers = []
    in_merge = False
    matching = re.search(r"연결하시오|짝지으시오|대응시키시오|매칭", question["stem"])
    for item in question["body"]:
        # Korean statements in a 보기 box can support separate answer choices.
        # Numbered options inside a box remain options; statement-only boxes stay out.
        boxed = item["kind"] == "box"
        children = item["items"] if boxed else [item]
        for child in children:
            if child["kind"] != "text":
                continue
            line = child["text"]
            if re.match(r"^※.*선지\s*통합", line):
                in_merge = True
            tagged = YEAR_PREFIX.match(line)
            if tagged and in_merge:
                continue
            match = CHOICE.match(tagged[2] if tagged else line)
            if match and (not boxed or matching or not re.match(r"[가-힣ㄱ-ㅎ]", match[1])):
                markers.append(match[1])
    return markers


def finalize_question(question: dict) -> None:
    question["choice_markers"] = choice_markers(question)
    question["history_items"] = history_items(question.get("history", ""))
    pages = []
    for found in PAGE_SUFFIX.finditer(question["stem"]):
        for item in found[1].split(","):
            ends = [int(v) for v in item.split("~")]
            if len(ends) == 1:
                pages.extend(ends)
            elif 0 <= ends[-1] - ends[0] <= 1000:
                pages.extend(range(ends[0], ends[-1] + 1))
    question["pages"] = list(dict.fromkeys(pages + question.get("pages", [])))
    if question.get("type") is None:
        stem = question["stem"]
        if re.search(r"연결하시오|짝지으시오|대응시키시오|매칭", stem) and question["choice_markers"]:
            question["type"] = "매칭형"
        elif question["choice_markers"]:
            question["type"] = "객관식"
        elif re.search(r"서술|설명하시오|비교하시오|기술하시오", stem):
            question["type"] = "서술형"
        else:
            question["type"] = "단답형"
            question["type_uncertain"] = True


def new_question(qid: str, section: str, section_order: int, stem: str, years: list[str],
                 location: str, **extra) -> dict:
    return {"qid": qid, "section": section, "section_order": section_order,
            "stem": stem, "years": years, "location": location, "body": [],
            "history": "", "type": None, "type_uncertain": False, **extra}


def paragraph_items(paragraph, doc) -> list[dict]:
    from docx.oxml.ns import qn

    body = []
    buffer = ""

    def flush():
        nonlocal buffer
        if buffer:
            body.append(text_item(buffer))
            buffer = ""

    for node in paragraph._p.iter():
        if node.tag == qn("w:t"):
            buffer += node.text or ""
        elif node.tag == qn("w:tab"):
            buffer += "\t"
        elif node.tag in (qn("w:br"), qn("w:cr")):
            flush()
        elif node.tag == qn("a:blip"):
            flush()
            rid = node.get(qn("r:embed"))
            if not rid:
                raise ValueError("Word 그림이 외부 파일을 참조합니다. 그림이 내장된 Word가 필요합니다.")
            part = doc.part.related_parts.get(rid)
            if part is None:
                raise ValueError(f"Word 그림 관계를 찾을 수 없습니다. {rid}")
            body.append({"kind": "image", "member": str(part.partname).lstrip("/")})
    flush()
    return body


def word_layout(paragraph) -> tuple[bool, bool, float]:
    sizes = [run.font.size.pt for run in paragraph.runs if run.font.size is not None]
    indent = paragraph.paragraph_format.left_indent
    return any(run.bold for run in paragraph.runs), indent is None or not indent.pt, max(sizes, default=0)


def verified_photo_word(path: Path) -> dict | None:
    with ZipFile(path) as archive:
        metadata = {name: archive.read(name) for name in ("_rels/.rels", "[Content_Types].xml")}
        has_photo = (PACKAGE_PART in archive.namelist() or
                     any(PACKAGE_PART.encode() in part for part in metadata.values()))
    if not has_photo:
        return None
    sys.path.insert(0, str(WORD_SCRIPTS))
    try:
        from shared_photos import grouping_plan, read_docx_manifest
        from yama_photo_word import audit_word_payload
        envelope = read_docx_manifest(path)
        if envelope is None:
            raise ValueError("Word에 기록된 공통사진 자료가 없습니다.")
        payload = envelope["payload"]
        word_qc = payload.get("word_qc", {})
        if word_qc.get("verdict") != "OK" or word_qc.get("word", {}).get("status") != "OK":
            raise ValueError("공통사진 Word의 레이아웃 검수 결과를 확인할 수 없습니다.")
        errors = audit_word_payload(path, payload, envelope["manifest"])
        if errors:
            raise ValueError(f"공통사진 Word와 출처가 일치하지 않습니다. {errors[:3]}")
        envelope["groups"] = grouping_plan(envelope["manifest"])["parents"]
    finally:
        sys.path.pop(0)
    return envelope


def apply_photo_records(envelope: dict, questions: list[dict]) -> list[dict]:
    records = {item["bookmark"]: item for item in envelope["payload"]["word_questions"]}
    aliases = {item["path"]: item for item in envelope["aliases"]}
    for q in questions:
        record = records.pop(q.pop("bookmark", None), None)
        if record is None:
            raise ValueError("공통사진 Word의 문항 ID와 실제 문항이 일치하지 않습니다.")
        q["qid"] = record["id"]
        q["pages"] = record.get("pages", [])
        q["source_order"] = record.get("source_order")
        if record.get("type") in TYPES:
            q["type"] = record["type"]
            q["type_uncertain"] = bool(record.get("type_uncertain"))
        # Keep the actual Word text and substitute the reviewed original aliases for shared references.
        group_images = defaultdict(deque)
        for membership in record.get("image_memberships", []):
            if membership["occurrence_id"] in record.get("shared_occurrences", []):
                group_images[membership["photo_id"]].append(aliases[membership["original_path"]])
        body = []
        for item in q["body"]:
            if item["kind"] != "text":
                body.append(item)
                continue
            pattern = re.compile(r"\[공통 사진 ([^\]]+?) 참조\]")
            position = 0
            for match in pattern.finditer(item["text"]):
                if not group_images[match[1]]:
                    raise ValueError(f"공통사진 연결을 확인할 수 없습니다. {q['qid']} {match[1]}")
                if position < match.start():
                    body.append(text_item(item["text"][position:match.start()]))
                alias = group_images[match[1]].popleft()
                body.append({"kind": "image", "member": alias["part"], "label": alias["path"],
                             "shared_photo": match[1]})
                position = match.end()
            if position < len(item["text"]):
                body.append(text_item(item["text"][position:]))
        q["body"] = body
        if any(group_images.values()):
            raise ValueError(f"Word에서 일부 공통사진 참조를 찾지 못했습니다. {q['qid']}")
    if records:
        raise ValueError("공통사진 Word의 일부 문항을 읽을 수 없습니다.")
    return [
        {"id": group["id"], "question_ids": group["question_ids"],
         "images": [{"photo_id": image["photo_id"],
                     "kind": "image", "member": aliases[image["path"]]["part"],
                     "label": image["path"], "question_ids": image["question_ids"],
                     "display_notes": image.get("display_notes", [])}
                    for image in group["images"]]}
        for group in envelope["groups"]
    ]


def word_questions(path: Path) -> dict:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph

    envelope = verified_photo_word(path)
    doc = Document(path)
    questions, editor_notes, warnings = [], [], []
    current = pending = None
    section, section_order, title, subtitle = "", 0, "", ""
    in_editor = False
    parent_bookmarks = {row["bookmark"] for row in envelope["payload"]["word_parents"]} if envelope else set()
    parent_ends = set()
    pending_bookmark = None
    for position, element in enumerate(doc.element.body, 1):
        starts = {node.get(qn("w:name")): node.get(qn("w:id"))
                  for node in element.iter(qn("w:bookmarkStart"))}
        ends = {node.get(qn("w:id")) for node in element.iter(qn("w:bookmarkEnd"))}
        parent_ends.update(v for k, v in starts.items() if k in parent_bookmarks)
        if parent_ends:
            parent_ends.difference_update(ends)
            continue
        if bookmark := next((b for b in starts if b and b.startswith("YSP_Q_")), None):
            pending_bookmark = bookmark

        if element.tag == qn("w:tbl"):
            table_items = [item for child in element.iter(qn("w:p"))
                           for item in paragraph_items(Paragraph(child, doc), doc)]
            if not table_items:
                continue
            first = table_items[0]
            if first.get("text", "").startswith("공통지문"):
                if pending is not None:
                    raise ValueError(f"공통지문 표에 대응하는 문항이 없습니다. 본문 요소 {position}")
                pending = table_items
            elif current is None:
                raise ValueError(f"표 위에 대응하는 문항이 없습니다. 본문 요소 {position}")
            elif any(node.get(qn("w:val")) == "YamaStatementBox"
                     for node in element.xpath("./w:tblPr/w:tblCaption")):
                current["body"].append({"kind": "box", "label": "보기", "items": table_items})
            else:
                current["body"].extend(table_items)
            continue
        if element.tag != qn("w:p"):
            continue
        paragraph = Paragraph(element, doc)
        body = paragraph_items(paragraph, doc)
        if not body:
            continue
        line = body[0].get("text", "")
        bold, unindented, size = word_layout(paragraph)
        centered = paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
        tagged = YEAR_PREFIX.fullmatch(line)
        is_question = tagged and bold and unindented
        if pending is not None and not is_question:
            raise ValueError(f"공통지문 다음에 문항이 없습니다. 본문 요소 {position}")
        if centered and size >= 15 and not title:
            title = line
        elif centered and not subtitle and line.endswith("교수님)"):
            subtitle = line
        if EDITOR.fullmatch(line.strip()):
            in_editor = True
            continue
        if in_editor and not unindented and not tagged:
            editor_notes.extend(i["text"] for i in body if i["kind"] == "text")
            continue
        in_editor = False
        if tagged and unindented and not is_question:
            raise ValueError(f"연도 태그가 있는 문항의 Word 형식을 확인하세요. 본문 요소 {position}")
        if is_question:
            qid = f"q{len(questions)+1:03d}"
            current = new_question(qid, section or subtitle or title or path.stem, section_order,
                                   tagged[2], [y.strip() for y in tagged[1].split(",")],
                                   f"Word 본문 요소 {position}",
                                   **({"bookmark": pending_bookmark} if envelope else {}))
            pending_bookmark = None
            current["body"].extend(body[1:])
            if pending:
                current["context"] = pending
                pending = None
            questions.append(current)
        elif bold and unindented and size >= 12 and not centered and line:
            if line.startswith("[") and not tagged:
                raise ValueError(f"문항 머리의 연도 태그를 확인하세요. 본문 요소 {position}")
            section, section_order, current = line, section_order + 1, None
        elif current is not None:
            history = HISTORY.fullmatch(line.strip())
            if history:
                current["history"] = history[1]
                current["body"].extend(body[1:])
            elif line.lstrip().startswith("▸"):
                current.setdefault("editor_notes", []).extend(i["text"] for i in body if i["kind"] == "text")
            else:
                current["body"].extend(body)
    if pending:
        raise ValueError("파일 끝 공통지문에 대응하는 문항이 없습니다.")
    if not questions:
        raise ValueError("연도 태그가 붙은 문항을 찾을 수 없습니다.")
    groups = apply_photo_records(envelope, questions) if envelope else []
    for q in questions:
        if not q["stem"].strip() and not any(i["kind"] == "image" for i in q["body"]):
            raise ValueError(f"빈 문두에 연결된 그림이 없습니다. {q['qid']}")
    if any(re.match(r"^\s*(?:상기|위)\s*(?:증례|환자|소견|세균|사례|지문)", q["stem"])
           and not q.get("context") for q in questions):
        warnings.append("공통지문이 필요한 문항의 지문 연결을 원본에서 확인하세요.")
    return {"title": title, "editor_notes": editor_notes, "warnings": warnings,
            "questions": questions, "groups": groups, "original_aliases": envelope["aliases"] if envelope else []}


def quote_line(line: str) -> tuple[int, str]:
    depth = 0
    while line.startswith(">"):
        depth += 1
        line = line[1:].removeprefix(" ")
    return depth, line


def raw_questions(text: str) -> list[dict]:
    sys.path.insert(0, str(WORD_SCRIPTS))
    try:
        import build_yama_docx as builder
        from yama_bank import parse_bank
    finally:
        sys.path.pop(0)
    questions = []
    for order, chapter in enumerate(parse_bank(text, builder)):
        for record in chapter["questions"]:
            year = year_label(record["year"])
            first, last = record["_source_span"]
            location = f"줄 {first}" + (f"–{last}" if last != first else "")
            q = new_question(record["_source_id"], chapter["topic"], order, record["stem"], [year],
                             location, source_tag=record["_original_tag"],
                             source_qid=record.get("_source_qid"))
            if record["no"]:
                year_text = year + ("년" if year != "연도미상" else "")
                round_text = record["round"] if re.fullmatch(r"\d+차|차수미상", record["round"]) else ""
                q["history"] = " ".join(filter(None, (year_text, round_text, f"{record['no']}번")))
            for line in record["body"]:
                if line.strip():
                    q["body"].extend(markdown_items(line))
            questions.append(q)
    return questions


def markdown_questions(path: Path) -> dict:
    text = path.read_text(encoding="utf-8-sig")
    if re.search(r"^\s*(?:<!-- question qid=|\*\*Q\.)", text, re.M):
        return {"questions": raw_questions(text), "title": path.stem,
                "groups": [], "editor_notes": [], "warnings": []}
    questions = []
    current = None
    section, section_order, depth, box = path.stem, 0, 0, None
    shared_photos = {}
    photo_id = None
    group_depth = None
    for number, line in enumerate(text.splitlines(), 1):
        if re.match(r"^#{2,6}\s", line):
            section = line.split(" ", 1)[1]
            current, section_order = None, section_order + 1
            continue
        d, inner = quote_line(line)
        callout = QUESTION_HEADER.fullmatch(inner.strip()) if d else None
        if callout and not callout[1] and callout[2] == "[공통사진 문항]":
            current = None
            group_depth = d
            shared_photos = {}
            photo_id = None
            continue
        if group_depth is not None and d == group_depth:
            if header := re.fullmatch(r"공통 사진\s+(.+)", inner):
                photo_id = header[1]
            elif photo_id and EMBED.search(inner):
                shared_photos[photo_id] = next(i for i in markdown_items(inner) if i["kind"] == "image")
                photo_id = None
        elif group_depth is not None and d < group_depth:
            group_depth = None
            shared_photos = {}
        if callout:
            qid = f"q{len(questions)+1:03d}"
            box = None
            depth = d
            header = callout[2]
            tags = YEAR_PREFIX.fullmatch(header)
            years = [y.strip() for y in tags[1].split(",")] if tags else [
                year_label(m[0].removesuffix("년")) for m in re.finditer(r"(?:건양대)?\d{2,4}년|연도미상", header)
            ]
            current = new_question(qid, section, section_order, "", years, f"줄 {number}",
                                   type=callout[1], history=tags[2] if tags else header)
            questions.append(current)
            continue
        if current is None:
            continue
        if d < depth or inner.startswith("[!answer]"):
            current = None
            continue
        qid_tag = QID_TAG.fullmatch(inner.strip())
        if qid_tag:
            current["qid"] = qid_tag[1]
            continue
        if inner == "<!-- yama-stem-in-body -->":
            current["stem_in_body"] = True
            continue
        if inner.startswith("[!quote] "):
            label = inner.removeprefix("[!quote] ")
            if label.startswith("공통지문"):
                box = [text_item(label)]
                current["context"] = box
            else:
                record = {"kind": "box", "label": label, "items": []}
                box = record["items"]
                current["body"].append(record)
            continue
        if d == depth:
            box = None
        if not inner.strip():
            continue
        if box is not None:
            box.extend(markdown_items(inner))
        elif not current["stem"] and not current.get("stem_in_body"):
            current["stem"] = inner
        else:
            current["body"].extend(markdown_items(inner))
            if reference := re.fullmatch(r"\[?공통 사진\s+(.+?)\s+참조\]?", inner):
                photo = shared_photos.get(reference[1])
                if photo:
                    current["body"].append({**photo, "shared_photo": reference[1]})
                else:
                    current.setdefault("warnings", []).append(f"공통 사진 {reference[1]}을 찾을 수 없습니다.")
    if not questions or any(not q["stem"] and not q.get("stem_in_body") for q in questions):
        raise ValueError("문항을 찾지 못했거나 문두가 비어 있습니다. 문항 번호와 원문을 확인하세요.")
    return {"questions": questions, "title": path.stem, "groups": [], "editor_notes": [], "warnings": []}


def make_index(path: Path) -> dict:
    if path.suffix.lower() not in (".docx", ".md"):
        raise ValueError("질문 인덱스는 DOCX 또는 Markdown을 지원합니다.")
    source = stamp(path)
    result = word_questions(path) if path.suffix.lower() == ".docx" else markdown_questions(path)
    verify_stamp(source)
    result["source"] = source
    for question in result["questions"]:
        finalize_question(question)
    if len({q["qid"] for q in result["questions"]}) != len(result["questions"]):
        raise ValueError("중복된 문항 ID가 있습니다.")
    return result


def match_questions(index: dict, *, qid=None, year=None, exam_round=None, number=None,
                    section=None, text=None) -> list[dict]:
    if not (qid or number or text):
        raise ValueError("문항 ID, 원본 문항번호 또는 문두 일부를 지정하세요.")
    result = []
    for question in index["questions"]:
        if qid and qid not in (question["qid"], question.get("source_qid")):
            continue
        if section and normalize(section) not in normalize(question["section"]):
            continue
        if text and normalize(text) not in normalize(question_text(question)):
            continue
        if number or exam_round:
            matches = [
                item for item in question["history_items"]
                if (not year or year_label(year) == item["year"])
                and (not number or normalize(str(number)) == normalize(item["number"]))
                and (not exam_round or normalize(exam_round) == normalize(item["round"] or ""))
            ]
            if not matches:
                continue
        elif year and year_label(year) not in question["years"]:
            continue
        result.append(question)
    return result


def materialize(item: dict, source: Path, output: Path, *, verify_only=False) -> Path:
    if "member" in item:
        member = item["member"]
        with ZipFile(source) as archive:
            data = archive.read(member)
        suffix = Path(member).suffix.lower()
    else:
        reference = Path(item["path"])
        target = source.parent / reference
        if not target.is_file() and len(reference.parts) == 1:
            target = source.parent / "동봉 이미지" / reference
        if not target.is_file():
            raise ValueError(f"실제 그림 파일을 찾을 수 없습니다. {reference}")
        data, suffix = target.read_bytes(), target.suffix.lower()
    if not re.fullmatch(r"\.[a-z0-9]{1,8}", suffix):
        raise ValueError("그림 파일 형식을 확인할 수 없습니다.")
    digest = hashlib.sha256(data).hexdigest()
    target = output / f"yama_{digest[:20]}{suffix}"
    if target.exists():
        if target.read_bytes() != data:
            raise ValueError(f"다른 그림 파일이 같은 경로에 있습니다. {target}")
    elif verify_only:
        raise ValueError(f"Markdown에 필요한 그림 파일이 없습니다. {target}")
    else:
        output.mkdir(parents=True, exist_ok=True)
        with target.open("xb") as stream:
            stream.write(data)
    return target.resolve()


def question_figures(items: list[dict], source: Path, output: Path | None) -> list[dict]:
    result = []
    for item in items:
        if item["kind"] == "box":
            item = {**item, "items": question_figures(item["items"], source, output)}
        elif item["kind"] == "image":
            item = {**item, **({"local_file": str(materialize(item, source, output))} if output else {})}
        result.append(item)
    return result


def lookup(args) -> dict:
    index = read_json(args.index)
    source = verify_stamp(index["source"])
    matches = match_questions(index, qid=args.qid, year=args.year, exam_round=args.round,
                              number=args.number, section=args.section, text=args.text)
    if len(matches) != 1:
        return {"status": "ambiguous" if matches else "not_found", "count": len(matches),
                "candidates": [{**{k: q[k] for k in ("qid", "section", "history", "location")},
                                "stem": question_text(q)[:140]} for q in matches[:args.limit]]}
    q = matches[0].copy()
    for field in ("context", "body"):
        if field in q:
            q[field] = question_figures(q[field], source, args.figures)
    return {"status": "ok", "source": str(source), "question": q}


def main(argv=None) -> int:
    cli = argparse.ArgumentParser(description=__doc__)
    commands = cli.add_subparsers(dest="command", required=True)
    create = commands.add_parser("build")
    create.add_argument("--source", type=Path, required=True)
    create.add_argument("--out", type=Path, required=True)
    find = commands.add_parser("lookup")
    find.add_argument("--index", type=Path, required=True)
    for option in ("qid", "year", "round", "number", "section", "text"):
        find.add_argument(f"--{option}")
    find.add_argument("--figures", type=Path)
    find.add_argument("--limit", type=int, choices=range(1, 11), default=5)
    args = cli.parse_args(argv)
    try:
        if args.command == "build":
            if args.out.resolve() == args.source.resolve():
                raise ValueError("출력 인덱스에 원본 경로를 사용할 수 없습니다.")
            index = make_index(args.source)
            write_index(args.out, index, "questions", [index["source"]])
            result = {"status": "ok", "index": str(args.out.resolve()), "count": len(index["questions"]),
                      "warnings": index["warnings"]}
        else:
            result = lookup(args)
    except (BadZipFile, ImportError, KeyError, OSError, ValueError) as exc:
        print(f"CHECK — {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["status"] == "ok" else 1


if __name__ == "__main__":
    raise SystemExit(main())
