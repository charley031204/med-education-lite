"""Reviewed photo groups for both Word builders; source leaves remain intact."""
import copy
import hashlib
import json
import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.text.paragraph import Paragraph

import shared_photos as core
from yama_bank import professor_in_current


REVISION = "yama-print/1.0.7+shared-photos-1.3.4"
SCHEMA = "yama-shared-photos/1"
IMG = re.compile(r"!\[\[([^\]|]+)(?:\|(\d+))?\]\]")


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def history_entry(raw, year_tag):
    year = year_tag(raw["year"])
    year += "" if year == "연도미상" else "년"
    number = raw.get("no")
    if not number:
        number = "번호 미상"
    elif number not in {"무번호", "번호미상", "번호 미상"} and not number.endswith("번"):
        number += "번"
    return " ".join(part for part in (year, raw.get("round", ""), number) if part)


def source_history_from_raw(raw, builder):
    """Preserve the source tag's year, examination round, and question number."""
    return "※ 출제 이력: " + history_entry(raw, builder.ytag)


def source_questions(chapters, rawroot, resolve_image, *, year_prof=None, matching_table=None):
    """Take the entire selected photo inventory before semantic merging."""
    result = []
    source_order = 0
    year_prof, matching_table = year_prof or {}, matching_table or {}
    for ci, chapter in enumerate(chapters):
        info = matching_table.get(chapter["topic"], {})
        current = chapter.get("current") or info.get("current") or "미상"
        for qi, question in enumerate(chapter["questions"]):
            source_order += 1
            qid = question.setdefault("_source_id", f"raw-c{ci + 1:04d}-q{qi + 1:05d}")
            year = question["year"]
            year_key = f"건양대{year + 1000:02d}" if year < 0 else (f"{year:02d}" if year else "연도미상")
            author = question.get("_explicit_prof") or next(
                (m[year_key] for key, m in year_prof.items()
                 if key != "default" and key in chapter["topic"] and year_key in m),
                year_prof.get("default", {}).get(year_key))
            other = bool(author and not professor_in_current(author, current))
            question["_source_section"] = ("타 교수 야마 · " if other else "") + chapter["topic"]
            question["_source_order"] = source_order
            images = []
            for bi, line in enumerate(question["body"]):
                for ii, match in enumerate(IMG.finditer(line)):
                    path = resolve_image(rawroot, match.group(1)).resolve()
                    images.append({"path": str(path), "original_reference": match.group(1),
                                   "occurrence_id": f"{qid}:body:{bi}:image:{ii}",
                                   "body_index": bi, "image_index": ii})
            section = f"{ci}:{chapter['topic']}" + (f":other:{author}" if other else "")
            result.append({"id": qid, "section": section,
                           "images": images,
                           "payload": {"source_question": copy.deepcopy(question)}})
    return result


def prepare_review(args, chapters, rawroot, resolve_image, matching_table=None):
    stem_images = sum(len(IMG.findall(q["stem"])) for ch in chapters for q in ch["questions"])
    if stem_images:
        return None, {"schema": SCHEMA, "verdict": "CHECK", "review_complete": False,
                      "source_image_count": stem_images + sum(len(IMG.findall(line)) for ch in chapters
                          for q in ch["questions"] for line in q["body"]),
                      "issues": ["문두 내부 이미지가 있다. 원문 위치를 보존하는 별도 입력 검토가 필요하다."]}
    year_prof = json.loads(Path(args.year_prof).read_text(encoding="utf-8")) if args.year_prof else None
    questions = source_questions(chapters, rawroot, resolve_image, year_prof=year_prof, matching_table=matching_table)
    if not any(q["images"] for q in questions):
        return None, {"schema": SCHEMA, "verdict": "NOT_APPLICABLE", "source_image_count": 0}
    for chapter in chapters:
        for question in chapter["questions"]:
            question["_preserve_leaf"] = True
    source_digest = sha(args.rawmd)
    review_path = getattr(args, "photo_review", None)
    if review_path:
        try:
            manifest = json.loads(Path(review_path).read_text(encoding="utf-8"))
            errors = core.validate_manifest(manifest, questions, source_digest, rawroot)
        except (OSError, ValueError, TypeError, KeyError) as exc:
            manifest, errors = None, [f"사진 검토 manifest를 읽을 수 없음: {exc}"]
        if not errors:
            return PhotoWordSession(manifest, questions, rawroot), None
        return None, {"schema": SCHEMA, "verdict": "CHECK", "review_complete": False,
                      "source_image_count": sum(len(q["images"]) for q in questions),
                      "review_file": str(review_path), "issues": errors}
    work = getattr(args, "photo_work", None)
    output = Path(work) if work else Path(args.out) / "_photo_review"
    output.mkdir(parents=True, exist_ok=True)
    try:
        manifest = core.prepare_manifest(questions, source_digest, rawroot)
    except (OSError, ValueError, KeyError) as exc:
        return None, {"schema": SCHEMA, "verdict": "CHECK", "review_complete": False,
                      "source_image_count": sum(len(q["images"]) for q in questions),
                      "issues": [f"전체 사진 inventory를 만들 수 없음: {exc}"]}
    suffix = manifest["inventory_digest"][:12]
    path = output / f"{Path(args.rawmd).stem}.{suffix}.photo-review.json"
    if not path.exists():
        core.export_manifest(manifest, path)
    sheets = output / f"{Path(args.rawmd).stem}.{suffix}.contact-sheets"
    core.contact_sheets(manifest, rawroot, sheets)
    return None, {"schema": SCHEMA, "verdict": "CHECK", "review_complete": False,
                  "source_image_count": len(manifest["inventory"]["appearances"]),
                  "review_file": str(path), "contact_sheets": str(sheets),
                  "issues": ["전체 사진 및 모든 고유 대표 쌍의 독립 검토를 완료한 뒤 --photo-review로 다시 실행한다."]}


def _bookmark(doc, elements, name):
    paragraphs = [p for element in elements for p in element.iter(qn("w:p"))]
    if not paragraphs:
        raise ValueError(f"Empty bookmark range: {name}")
    bid = max((int(n.get(qn("w:id"))) for n in doc.element.iter(qn("w:bookmarkStart"))), default=0) + 1
    start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end.set(qn("w:id"), str(bid))
    paragraphs[0].insert(1 if paragraphs[0].find(qn("w:pPr")) is not None else 0, start)
    paragraphs[-1].append(end)


def _range_nodes(doc, name):
    active, found, bid = False, [], None
    for node in doc.element.body.iter():
        if node.tag == qn("w:bookmarkStart") and node.get(qn("w:name")) == name:
            active, bid = True, node.get(qn("w:id"))
            continue
        if active and node.tag == qn("w:bookmarkEnd") and node.get(qn("w:id")) == bid:
            return found
        if active:
            found.append(node)
    raise ValueError(f"Missing or unclosed bookmark: {name}")


def _document_order(doc):
    """Bind visible paragraphs/drawings, including material outside source ranges."""
    order = []
    for p in doc.element.body.iter(qn("w:p")):
        text = Paragraph(p, doc).text
        pictures = [hashlib.sha256(doc.part.related_parts[n.get(qn("r:embed"))].blob).hexdigest()
                    for n in p.iter(qn("a:blip"))]
        if text or pictures:
            order.append({"text": text, "images": pictures})
    return order


def audit_word_payload(path, payload, manifest=None):
    """Compare actual XML text and actual relationship media with the display plan."""
    import build_yama_docx as builder
    if manifest is None:
        try:
            envelope = core.read_docx_manifest(path)
            manifest = envelope["manifest"] if envelope else None
        except (ValueError, KeyError, OSError) as exc:
            return [f"Document binding is invalid: {exc}"]
    if manifest is None:
        return ["A verified source manifest is required for Word audit"]
    errors = core.validate_manifest(manifest)
    if errors:
        return errors
    canonical_plan = core.grouping_plan(manifest)
    if payload.get("plan") != canonical_plan:
        errors.append("Word grouping plan differs from the reviewed source plan")
    if payload.get("questions") != manifest["inventory"]["questions"]:
        errors.append("Word source question inventory differs from the reviewed inventory")
    if payload.get("source_digest") != manifest["source_digest"]:
        errors.append("Word source digest differs from the reviewed source")
    source = {q["id"]: q for q in manifest["inventory"]["questions"]}
    parents = {p["id"]: p for p in canonical_plan["parents"]}
    owners = {q["id"]: p for p in canonical_plan["parents"] for q in p["questions"]}
    leaves = payload.get("word_questions", [])
    if sorted(q.get("id", "") for q in leaves) != sorted(source):
        errors.append("Word leaves omit, duplicate, or introduce source question ids")
    files = {f["path"]: f["sha256"] for f in manifest["inventory"]["files"]}
    def image_sequence(digests, block_key):
        for transform in payload.get("image_transformations", []):
            if transform.get("block") != block_key:
                continue
            digests = [d for value in digests for d in
                       (transform["panel_sha256"] if value == transform["source_sha256"] else [value])]
        return digests
    for leaf in leaves:
        original = source.get(leaf.get("id"))
        if original is None:
            continue
        raw = original.get("payload", {}).get("source_question", {})
        if not isinstance(raw, dict):
            errors.append(f"{leaf['id']}: missing structured original question")
            continue
        for key, expected in (("stem", raw["stem"]), ("source_body", raw["body"]),
                              ("years", [raw["year"]]),
                              ("nos", [f"{builder.ytag(raw['year'])}-{raw['no'] or '?'}"])):
            if leaf.get(key) != expected:
                errors.append(f"{leaf['id']}: {key} differs from the reviewed source")
        if leaf.get("source_section") != raw.get("_source_section") or leaf.get("section") != original["section"]:
            errors.append(f"{leaf['id']}: source section differs from the reviewed source")
        parent = owners[leaf["id"]]
        memberships = next(q["image_memberships"] for q in parent["questions"] if q["id"] == leaf["id"])
        if leaf.get("image_memberships") != memberships:
            errors.append(f"{leaf['id']}: image ownership differs from the source plan")
        shared = set(leaf.get("shared_occurrences", []))
        allowed = {m["occurrence_id"] for m in memberships
                   if next(i for i in parent["images"] if i["photo_id"] == m["photo_id"])["shared"]}
        if len(shared) != len(leaf.get("shared_occurrences", [])) or shared != allowed:
            errors.append(f"{leaf['id']}: invalid shared image occurrences")
        display = {"stem": raw["stem"], "body": copy.deepcopy(raw["body"]), "years": [raw["year"]],
                   "nos": [f"{builder.ytag(raw['year'])}-{raw['no'] or '?'}"], "extra_choices": [],
                   "_source_history": source_history_from_raw(raw, builder),
                   "_pages": leaf.get("pages", []), "authors": leaf.get("authors", []),
                   "_context_box": leaf.get("context_box"), "_answer": leaf.get("answer")}
        for bi, line in enumerate(display["body"]):
            ii = [0]
            def replace(match):
                image = next((m for m in memberships if m.get("body_index") == bi and m.get("image_index", 0) == ii[0]), None)
                ii[0] += 1
                return f"[공통 사진 {image['photo_id']} 참조]" if image and image["occurrence_id"] in shared else match.group(0)
            display["body"][bi] = IMG.sub(replace, line)
        expected = builder.expected_question_lines(display, leaf.get("author_label", False), leaf.get("page_mode", False))
        if leaf.get("source_history") != builder.source_history(display):
            errors.append(f"{leaf['id']}: source history differs from original year/round/number")
        if leaf.get("expected_lines") != expected:
            errors.append(f"{leaf['id']}: expected text is not derived from the original question")
        expected_images = image_sequence([files[m["original_path"]] for m in memberships if m["occurrence_id"] not in shared], leaf.get("layout_block"))
        if leaf.get("image_sha256") != expected_images:
            errors.append(f"{leaf['id']}: expected photo hashes differ from original ownership")
    for header in payload.get("word_parents", []):
        parent = parents.get(header.get("id"))
        if parent is None:
            errors.append(f"Unknown Word photo parent: {header.get('id')}")
            continue
        canonical_images = {i["photo_id"]: i for i in parent["images"] if i["shared"]}
        images = header.get("shared_images", [])
        if not images or images != list(canonical_images.values()):
            errors.append(f"{header['id']}: parent photo ownership differs from source plan")
        lines = []
        for item in images:
            lines.append(f"공통 사진 {item['photo_id']}" + (" 참조 (계속)" if header.get("continuation") else ""))
            if not header.get("continuation"):
                lines.extend(item.get("display_notes", []))
        if header.get("expected_lines") != lines:
            errors.append(f"{header['id']}: parent captions differ from the source plan")
        hashes = [] if header.get("continuation") else [files[i["path"]] for i in images]
        if header.get("image_sha256") != image_sequence(hashes, header.get("layout_block")):
            errors.append(f"{header['id']}: parent image hashes differ from source plan")
    headers = payload.get("word_parents", [])
    for parent_id, parent in parents.items():
        parent_headers = [h for h in headers if h.get("id") == parent_id]
        expected_initial = int(any(i["shared"] for i in parent["images"]))
        if sum(not h.get("continuation") for h in parent_headers) != expected_initial:
            errors.append(f"{parent_id}: common photos require exactly one initial header")
        blocks = [h.get("layout_block") for h in parent_headers]
        if len(blocks) != len(set(blocks)):
            errors.append(f"{parent_id}: duplicate shared-photo headers for one source leaf")
    for leaf in leaves:
        if leaf.get("id") not in owners:
            continue
        parent = owners[leaf["id"]]
        shared = set(leaf.get("shared_occurrences", []))
        required_photos = {m["photo_id"] for m in leaf.get("image_memberships", []) if m["occurrence_id"] in shared}
        shown_photos = {i["photo_id"] for h in headers if h.get("id") == parent["id"]
                        for i in h.get("shared_images", [])}
        if not required_photos.issubset(shown_photos):
            errors.append(f"{leaf['id']}: referenced common photo has no audited Word parent")
    if sum(bool(h.get("continuation")) for h in headers) != sum(bool(x.get("continuation")) for x in payload.get("continuation_ledger", [])):
        errors.append("Word continuation headers differ from the measured continuation ledger")
    measured = {b["key"]: b for b in payload.get("word_qc", {}).get("word", {}).get("blocks", [])}
    leaf_blocks = {q.get("layout_block"): q for q in leaves}
    expected_ledger = []
    for header in headers:
        leaf = leaf_blocks.get(header.get("layout_block"))
        block = measured.get(header.get("layout_block"))
        if not leaf or not block or leaf["id"] not in parents.get(header["id"], {}).get("question_ids", []):
            errors.append("A shared photo header is not measured with its own first source leaf")
            continue
        expected_ledger.append({"parent_id": header["id"], "first_question_id": leaf["id"],
                                "page": block.get("page_first"), "continuation": bool(header.get("continuation"))})
    order_key = lambda entry: (entry["parent_id"], entry["first_question_id"])
    if sorted(expected_ledger, key=order_key) != sorted(payload.get("continuation_ledger", []), key=order_key):
        errors.append("Shared-photo page chunks differ from measured parent/leaf bounds")
    chunks = [(entry["parent_id"], entry["page"]) for entry in expected_ledger]
    if len(chunks) != len(set(chunks)):
        errors.append("A photo parent has duplicate headers on the same measured page")
    doc = Document(path)
    if payload.get("document_order") != _document_order(doc):
        errors.append("Word contains omitted, reordered, or unaccounted paragraphs/images")
    records = payload.get("word_questions", []) + payload.get("word_parents", [])
    node_owners = {}
    for record in records:
        try:
            nodes = _range_nodes(doc, record["bookmark"])
            for node in nodes:
                if node.tag in (qn("w:t"), qn("a:blip")):
                    node_owners.setdefault(node, []).append(record)
            # Include each paragraph that contains a bookmarked text node once.
            paragraphs = []
            for node in nodes:
                if node.tag in (qn("w:t"), qn("w:br"), qn("w:tab")):
                    parent = node
                    while parent is not None and parent.tag != qn("w:p"):
                        parent = parent.getparent()
                    if parent is not None and parent not in paragraphs:
                        paragraphs.append(parent)
            lines = [Paragraph(p, doc).text for p in paragraphs if Paragraph(p, doc).text]
            digests = [hashlib.sha256(doc.part.related_parts[n.get(qn("r:embed"))].blob).hexdigest()
                       for n in nodes if n.tag == qn("a:blip")]
            if lines != record["expected_lines"]:
                errors.append(f"{record['id']}: Word text differs from source display plan")
            if digests != record["image_sha256"]:
                errors.append(f"{record['id']}: Word image assignment differs from source display plan")
        except (ValueError, KeyError) as exc:
            errors.append(f"{record.get('id')}: {exc}")
    for node in doc.element.body.iter(qn("a:blip")):
        if len(node_owners.get(node, [])) != 1:
            errors.append("Every actual Word picture must belong to exactly one reviewed source or parent range")
    labels = {q["payload"]["source_question"].get("_source_section", "") for q in source.values()}
    bases = {label.removeprefix("타 교수 야마 · ") for label in labels}
    seen_content, current_section, current_heading, checked_leaves = False, None, None, set()
    question_text = re.compile(r"^(?:\[(?:(?:건양대)?\d{2,4}|연도미상)(?:[,\]])|※\s*출제 이력:|[①②③④⑤⑥⑦⑧⑨])")
    for paragraph in doc.element.body.iter(qn("w:p")):
        text = Paragraph(paragraph, doc).text
        paragraph_object = Paragraph(paragraph, doc)
        if text and any(run.bold and run.font.size is not None and run.font.size.pt >= 12 for run in paragraph_object.runs):
            current_heading = text
        text_nodes = list(paragraph.iter(qn("w:t")))
        picture_nodes = list(paragraph.iter(qn("a:blip")))
        relevant = text_nodes + picture_nodes
        owners_here = {record["bookmark"]: record for node in relevant for record in node_owners.get(node, [])}
        if owners_here:
            seen_content = True
            if any(len(node_owners.get(node, [])) != 1 for node in relevant) or len(owners_here) != 1:
                errors.append("A source paragraph has overlapping or partly unaccounted bookmark content")
            record = next(iter(owners_here.values()))
            if record.get("id") in source and record["id"] not in checked_leaves:
                expected_section = source[record["id"]]["payload"]["source_question"].get("_source_section")
                if current_section is None and len(bases) == 1:
                    current_section = next(iter(bases))
                if current_section != expected_section:
                    errors.append(f"{record['id']}: actual Word section differs from the source section")
                if record.get("source_section_heading") != current_heading:
                    errors.append(f"{record['id']}: actual Word heading differs from the source range heading")
                checked_leaves.add(record["id"])
            continue
        if not text:
            continue
        matched = None
        for base in bases:
            if text in (base, "타 교수 야마 · " + base):
                matched = text
                break
            if text.startswith(base + " — "):
                matched = ("타 교수 야마 · " if " — 출제: " in text else "") + base
                break
        if matched:
            current_section = matched
            continue
        if text == "타 교수 야마":
            current_section = None
            continue
        if question_text.match(text):
            errors.append("An actual question/choice/history paragraph is outside every reviewed source range")
        elif seen_content and text != "이번 연도 담당 교수와 기출 출제 교수가 다릅니다. 출제 스타일이 달라질 수 있으므로 참고용으로 활용하십시오.":
            errors.append("An unaccounted paragraph appears after source content starts")
    return errors


audit_shared_photo_docx = audit_word_payload


class PhotoWordSession:
    """Shared renderer API: emit(doc, layout, qs, emit_one, builder, author_label)."""
    def __init__(self, manifest, questions, image_root):
        self.manifest, self.questions = manifest, questions
        self.image_root = image_root if isinstance(image_root, dict) else Path(image_root)
        self.plan = core.grouping_plan(manifest)
        self.parents = {q: parent for parent in self.plan["parents"] for q in parent["question_ids"]}
        self.source = {q["id"]: q for q in questions}
        self.word_questions, self.word_parents, self.groups = [], [], []
        self.shared_image_count = 0

    def _path(self, path):
        return Path(self.image_root[path]) if isinstance(self.image_root, dict) else self.image_root / path

    def _header(self, doc, parent_id, shared, builder, continuation=False):
        before = len(doc.element.body)
        labels = []
        for index, image in enumerate(shared):
            label = f"공통 사진 {image['photo_id']}" + (" 참조 (계속)" if continuation else "")
            labels.append(label)
            p = doc.add_paragraph()
            builder.set_run_font(p.add_run(label), 10, bold=True)
            if not continuation:
                builder.add_image(doc, self._path(image["path"]), None, [])
            for note in ([] if continuation else image.get("display_notes", [])):
                labels.append(note)
                p = doc.add_paragraph()
                builder.set_run_font(p.add_run(note), 9)
        elements = list(doc.element.body)[before - 1:-1]
        bookmark = f"YSP_P_{len(self.word_parents) + 1:05d}"
        _bookmark(doc, elements, bookmark)
        record = {"id": parent_id, "bookmark": bookmark, "expected_lines": labels,
                  "image_sha256": [] if continuation else [sha(self._path(i["path"])) for i in shared],
                  "shared_images": copy.deepcopy(shared), "continuation": continuation}
        self.word_parents.append(record)
        if not continuation:
            self.shared_image_count += len(shared)
        return elements, record

    def emit(self, doc, layout, questions, emit_one, builder, author_label=False):
        headings = [p.text for p in doc.paragraphs if p.text and any(
            run.bold and run.font.size is not None and run.font.size.pt >= 12 for run in p.runs)]
        section_heading = headings[-1] if headings else None
        by_id = {q.get("_source_id"): q for q in questions if q.get("_source_id")}
        done = set()
        for question in questions:
            qid = question.get("_source_id")
            if qid in done:
                continue
            parent = self.parents.get(qid)
            members = ([by_id[i] for i in parent["question_ids"] if i in by_id] if parent else [question])
            # Preserve the selected order inside a shared group.
            members.sort(key=lambda q: questions.index(q))
            ids = {q.get("_source_id") for q in members}
            shared = [i for i in parent["images"] if len(ids.intersection(i["question_ids"])) > 1] if parent else []
            parent_id = parent["id"] if parent else f"single:{len(self.word_questions) + 1}"
            group_record = {"id": parent_id, "shared": shared, "leaves": [], "headers": {}}
            group_start = len(doc.element.body)
            if shared:
                _, header = self._header(doc, parent_id, shared, builder)
                group_record["headers"][members[0].get("_source_id")] = header
            for index, original in enumerate(members):
                q = copy.deepcopy(original)
                qid = q.get("_source_id") or "merged:" + hashlib.sha256(json.dumps(q["nos"]).encode()).hexdigest()[:16]
                q["_source_history"] = source_history_from_raw(self.source[qid]["payload"]["source_question"], builder)
                membership = next((x["image_memberships"] for x in parent["questions"] if x["id"] == qid), []) if parent else []
                shared_ids = {i["photo_id"] for i in shared}
                substitutions = {(m["body_index"], m.get("image_index", 0)): m
                                 for m in membership if m["photo_id"] in shared_ids}
                for bi, line in enumerate(q["body"]):
                    ii = [0]
                    def replace(match):
                        key = (bi, ii[0]); ii[0] += 1
                        item = substitutions.get(key)
                        return f"[공통 사진 {item['photo_id']} 참조]" if item else match.group(0)
                    q["body"][bi] = IMG.sub(replace, line)
                start = len(doc.element.body)
                emit_one(q, author_label, group_start if shared and index == 0 else None)
                if shared and index == 0:
                    group_record["headers"][original.get("_source_id")]["layout_block"] = layout.blocks[-1]["key"]
                elements = list(doc.element.body)[start - 1:-1]
                bookmark = f"YSP_Q_{len(self.word_questions) + 1:05d}"
                _bookmark(doc, elements, bookmark)
                expected_images = [sha(builder.resolve_image(self.image_root, m.group(1)))
                                   for line in q["body"] for m in IMG.finditer(line)]
                expected_images += [sha(builder.resolve_image(self.image_root, m.group(1)))
                                    for _, line in q.get("extra_choices", []) for m in IMG.finditer(line)]
                source_raw = next(iter(original.get("_source_questions", [])), {})
                record = {"id": qid, "section": parent["section"] if parent else "unpictured",
                          "source_section": getattr(self, "section_label", None) or source_raw.get("_source_section", "기출"),
                          "source_section_heading": section_heading,
                          "source_order": source_raw.get("_source_order", len(self.word_questions) + 1),
                          "years": q["years"], "nos": q["nos"], "stem": original["stem"],
                          "source_body": copy.deepcopy(original["body"]),
                          "extra_choices": copy.deepcopy(original.get("extra_choices", [])),
                          "note": original.get("note"),
                          "context_box": copy.deepcopy(original.get("_context_box")),
                          "answer": copy.deepcopy(original.get("_answer")),
                          "source_history": builder.source_history(q), "pages": q.get("_pages", []),
                          "authors": q.get("authors", []), "bookmark": bookmark,
                          "author_label": author_label, "page_mode": self.page_mode,
                          "layout_block": layout.blocks[-1]["key"],
                          "expected_lines": builder.expected_question_lines(q, author_label, self.page_mode),
                          "image_sha256": expected_images, "image_memberships": membership,
                          "shared_occurrences": [m["occurrence_id"] for m in substitutions.values()],
                          "source_questions": copy.deepcopy(original.get("_source_questions", []))}
                self.word_questions.append(record)
                group_record["leaves"].append({"id": qid, "block": layout.blocks[-1]})
                done.add(qid)
            if shared:
                self.groups.append(group_record)
                layout.group_chains.append([leaf["block"] for leaf in group_record["leaves"]])

    def add_continuations(self, doc, layout, out, report, builder):
        """Add a text reference before the first leaf on a new measured page."""
        self.continuation_ledger = []
        for _ in range(len(self.word_questions) + 1):
            word = report.get("word", {})
            if word.get("status") != "OK":
                return report
            pages = {b["key"]: b for b in word.get("blocks", [])}
            changed = False
            for group in self.groups:
                current_page = None
                for leaf in group["leaves"]:
                    block, qid = leaf["block"], leaf["id"]
                    location = pages.get(block["key"], {})
                    page = location.get("page_first")
                    if qid in group["headers"]:
                        current_page = page
                        continue
                    if page is not None and page != current_page:
                        elements, header = self._header(doc, group["id"], group["shared"], builder, True)
                        before = block["elements"][0]
                        for element in elements:
                            element.getparent().remove(element)
                            before.addprevious(element)
                        block["elements"] = elements + block["elements"]
                        group["headers"][qid] = header
                        header["layout_block"] = block["key"]
                        changed = True
                        current_page = page
            if not changed:
                for group in self.groups:
                    for leaf in group["leaves"]:
                        if leaf["id"] in group["headers"]:
                            self.continuation_ledger.append({"parent_id": group["id"], "first_question_id": leaf["id"],
                                "page": pages.get(leaf["block"]["key"], {}).get("page_first"),
                                "continuation": group["headers"][leaf["id"]]["continuation"]})
                return report
            report = layout.save(out)
        report["verdict"] = "CHECK"
        report.setdefault("issues", []).append("Shared-photo continuation pagination did not converge")
        return report

    def finalize(self, out, layout_report):
        transformations = [copy.deepcopy(t) for t in layout_report.get("transformations", []) if t["kind"] == "split"]
        for record in self.word_questions + self.word_parents:
            record["source_image_sha256"] = list(record["image_sha256"])
            for transform in transformations:
                if transform.get("block") != record.get("layout_block"):
                    continue
                record["image_sha256"] = [digest for value in record["image_sha256"]
                    for digest in (transform["panel_sha256"] if value == transform["source_sha256"] else [value])]
        payload = {"schema": "yama-shared-photo-word/1", "source_digest": self.manifest["source_digest"],
                   "questions": self.questions, "plan": self.plan,
                   "word_questions": self.word_questions, "word_parents": self.word_parents,
                   "continuation_ledger": getattr(self, "continuation_ledger", []),
                   "image_transformations": transformations,
                   "word_qc": {key: copy.deepcopy(layout_report.get(key)) for key in ("verdict", "checks", "word")}}
        payload["document_order"] = _document_order(Document(out))
        errors = core.validate_manifest(self.manifest, self.questions, self.manifest["source_digest"], self.image_root)
        errors += audit_word_payload(out, payload, self.manifest)
        actual = [q["id"] for q in self.word_questions if q["id"] in self.source]
        if sorted(actual) != sorted(self.source) or len(actual) != len(set(actual)):
            errors.append("Source photo question coverage differs from Word leaves")
        for q in self.word_questions:
            original = self.source.get(q["id"], {}).get("payload", {}).get("source_question")
            if original and (q["stem"] != original["stem"] or q["source_body"] != original["body"]):
                errors.append(f"{q['id']}: source reconstruction differs")
        binding_ok = False
        if not errors:
            core.embed_docx_manifest(out, self.manifest, payload, image_root=self.image_root)
            binding_ok = core.read_docx_manifest(out) is not None
        layout_report["document_sha256"] = sha(out)
        layout_report["format_revision"] = REVISION
        summary = {"schema": SCHEMA, "verdict": "OK" if not errors and binding_ok else "CHECK",
                   "manifest_digest": core.manifest_digest(self.manifest),
                   "inventory_digest": self.manifest["inventory_digest"],
                   "source_question_count": len(self.questions), "parent_count": len(self.plan["parents"]),
                   "photo_count": len(self.plan["photos"]),
                   "source_image_count": len(self.manifest["inventory"]["appearances"]),
                   "review_complete": not core.validate_manifest(self.manifest),
                   "source_reconstruction_ok": not errors, "document_binding_ok": binding_ok,
                   "continuation_ledger": payload["continuation_ledger"], "issues": errors}
        layout_report["shared_photos"] = summary
        Path(out).with_suffix(".layout-qc.json").write_text(json.dumps(layout_report, ensure_ascii=False, indent=2), encoding="utf-8")
        return summary


def build_shared_photo_docx(payload, manifest, image_paths, out=None, *, out_path=None,
                            title="기출 문항집", question_ids=None):
    """Rebuild a validated source payload using the same Word renderer and QC."""
    from types import SimpleNamespace
    from docx.shared import Mm
    import build_yama_docx as builder
    from yama_layout import YamaLayout

    errors = core.validate_manifest(manifest, image_root=image_paths)
    if errors:
        raise ValueError("; ".join(errors))
    out = out_path or out
    if not out:
        raise ValueError("An output DOCX path is required")
    if question_ids is not None:
        all_ids = [q["id"] for q in payload["word_questions"]]
        if set(question_ids) != set(all_ids) or len(question_ids) != len(all_ids):
            raise ValueError("Filtered shared-photo Word export needs a new source-bound projection review")
        payload = copy.deepcopy(payload)
        payload["word_questions"].sort(key=lambda q: question_ids.index(q["id"]))
    session = PhotoWordSession(manifest, payload["questions"], image_paths)
    session.page_mode = any(q.get("pages") for q in payload["word_questions"])
    references = dict(image_paths)
    for source in payload["questions"]:
        for entry in source["images"]:
            reference = entry.get("original_reference", entry["path"])
            path = image_paths[entry["path"]]
            if reference in references and Path(references[reference]) != Path(path):
                raise ValueError(f"Ambiguous original image reference: {reference}")
            references[reference] = path
    def resolve(_root, reference):
        if reference not in references:
            raise ValueError(f"Unmapped source photo: {reference}")
        return Path(references[reference])
    proxy = SimpleNamespace(**{name: getattr(builder, name) for name in
        ("set_run_font", "add_image", "expected_question_lines", "source_history", "ytag")}, resolve_image=resolve)
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.2)
    builder.set_run_font(doc.add_paragraph().add_run(title), 16, bold=True)
    builder.add_page_number(section)
    layout = YamaLayout(doc, source_dirs={Path(p).parent for p in image_paths.values()})
    questions = []
    for record in payload["word_questions"]:
        questions.append({"_source_id": record["id"], "years": record["years"], "nos": record["nos"],
            "stem": record["stem"], "body": copy.deepcopy(record["source_body"]),
            "extra_choices": copy.deepcopy(record.get("extra_choices", [])), "note": record.get("note"),
            "_source_history": record.get("source_history"),
            "_context_box": copy.deepcopy(record.get("context_box")),
            "_answer": copy.deepcopy(record.get("answer")),
            "_pages": record.get("pages", []), "authors": record.get("authors", []),
            "_source_questions": record.get("source_questions", []), "dup_count": 1})
    checks, missing = [], []
    def emit_one(q, author_label=False, block_start=None):
        start = len(doc.element.body)
        display = copy.deepcopy(q)
        for index, line in enumerate(display["body"]):
            display["body"][index] = IMG.sub(lambda m: f"![[{resolve(None, m.group(1))}]]", line)
        # The standard renderer resolves absolute image paths directly.
        builder.render_question(doc, display, Path("."), missing, author_label, session.page_mode)
        elements = list(doc.element.body)[start - 1:-1]
        checks.append(builder.document_lines(doc, elements) == builder.expected_question_lines(q, author_label, session.page_mode))
        layout.add_block(block_start if block_start is not None else start, f"{len(checks)}:{q['_source_id']}")
    sections = []
    for record in payload["word_questions"]:
        key = (record["section"], record.get("source_section_heading"))
        if key not in sections:
            sections.append(key)
    for key in sections:
        source_record = next(record for record in payload["word_questions"]
                             if (record["section"], record.get("source_section_heading")) == key)
        session.section_label = source_record["source_section"]
        builder.render_section_heading(doc, key[1] or source_record["source_section"])
        subset = [q for q, record in zip(questions, payload["word_questions"])
                  if (record["section"], record.get("source_section_heading")) == key]
        session.emit(doc, layout, subset, emit_one, proxy, source_record.get("author_label", False))
    measured = layout.save(out)
    measured = session.add_continuations(doc, layout, out, measured, proxy)
    shared = session.finalize(out, measured)
    source_ok = all(checks) and len(checks) == len(questions) and shared["source_reconstruction_ok"]
    return {"output": str(out), "verdict": "OK" if source_ok and not missing and measured["verdict"] == shared["verdict"] == "OK" else "CHECK",
            "source_content_ok": source_ok, "missing_images": missing, "layout": measured,
            "shared_photos": shared, "expected_blocks": len(questions),
            "chapters": [{"count_ok": True, "merged_groups": [], "merge_candidates": []}]}
