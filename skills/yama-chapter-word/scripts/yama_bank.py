"""Index a local exam bank and retrieve bounded packets for lecture review."""
import argparse
import csv
import hashlib
import json
import re
from collections import Counter, defaultdict
from copy import deepcopy
from pathlib import Path
import unicodedata


SCHEMA = "yama-bank/1"
MARKED_START = re.compile(r"<!-- question qid=(\S+) topic=(\S+) year=(\d{4}) -->")
MARKED_HEAD = re.compile(r"\*\*Q\.\s*(?P<no>.+?)\.\*\*\s+\[(?P<tag>[^]]+)](?P<flags>.*)")
QUESTION_HEAD = re.compile(r"^\s*\*\*Q\.")
PROF = re.compile(r"\*과거 교수:\s*(.*?)\*")
UNKNOWN_PROF = {"", "-", "—", "미상", "미표기", "원문 미표기"}
TYPE_CUES = re.compile(r"서술하시오|설명하시오|기술하시오|정의하시오|쓰시오|적으시오")
MARKUP = re.compile(r"(\*\*|__)(?=\S)(.+?)(?<=\S)\1|<u>(.*?)</u>")


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def check_output_paths(outputs, inputs):
    resolved_inputs = {Path(path).resolve() for path in inputs if path}
    seen = set()
    for path in outputs:
        if not path:
            continue
        resolved = Path(path).resolve()
        if resolved in resolved_inputs or resolved in seen:
            raise ValueError(f"Output path conflicts with an input or another output: {path}")
        seen.add(resolved)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def folded(text):
    return "".join(unicodedata.normalize("NFC", text or "").casefold().split())


def professor_in_current(author, current):
    def names(text):
        return {folded(re.sub(r"\s*교수(?:님)?$", "", part.split("(", 1)[0].strip()))
                for part in re.split(r"[/·]", text or "") if part.strip()}

    return bool(names(author).intersection(names(current)))


def display_markup(text):
    # Obsidian quote prefixes and bold/underline delimiters are presentation markup.
    text = re.sub(r"^>\s?", "", text)
    return MARKUP.sub(lambda m: m.group(2) if m.group(1) else m.group(3), text)


def marked_chapters(lines, builder):
    """Read the organizer's delimited question/body format, with line references."""
    chapters, current, index = [], None, 0
    while index < len(lines):
        line = lines[index].strip()
        if chapter := builder.CHAPTER_RE.fullmatch(line):
            topic, past, prof = builder.parse_chapter_title(chapter[1])
            current = {"topic": topic, "past": past, "current": prof, "questions": []}
            chapters.append(current)
        elif re.match(r"^#{1,3}\s", line):
            current = None
        elif line.startswith("<!-- question qid="):
            marker = MARKED_START.fullmatch(line)
            if marker is None or current is None:
                raise ValueError(f"Invalid question marker at line {index + 1}")
            qid, topic_id, year = marker.groups()
            header = MARKED_HEAD.fullmatch(lines[index + 1].strip()) if index + 1 < len(lines) else None
            if header is None:
                raise ValueError(f"Invalid question header at line {index + 2}")
            body_marker = f"<!-- body {qid} start -->"
            end_marker = f"<!-- body {qid} end -->"
            question_end = f"<!-- question {qid} end -->"
            pos = index + 2
            explicit_prof = None
            while pos < len(lines) and lines[pos].strip() != body_marker:
                metadata = lines[pos].strip()
                match = PROF.fullmatch(metadata)
                if metadata and not match:
                    raise ValueError(f"Unexpected question metadata at line {pos + 1}")
                if match and match[1] not in UNKNOWN_PROF:
                    explicit_prof = match[1]
                pos += 1
            start = pos + 1
            pos = start
            while pos < len(lines) and lines[pos].strip() != end_marker:
                if lines[pos].lstrip().startswith("<!-- question"):
                    raise ValueError(f"Missing body end for {qid}")
                pos += 1
            if pos + 1 >= len(lines) or lines[pos + 1].strip() != question_end:
                raise ValueError(f"Missing closing markers for {qid}")
            original = "\n".join(lines[start:pos]).strip("\n")
            body = [display_markup(part) for part in original.splitlines()]
            while body and not body[0].strip():
                body.pop(0)
            if not body:
                raise ValueError(f"Empty body for {qid}")
            # Leading figures and their captions remain together in the body.
            first = body[0].lstrip()
            if builder.IMG_RE.search(first) or first.startswith("[Figure"):
                stem = ""
            else:
                stem = body.pop(0)
            source_tag = header["tag"]
            tag = re.search(r"\b(?:\d{4}|\d{2})(?:본1)?[_ ]+(.+)", source_tag)
            exam_round = tag[1] if tag and re.fullmatch(r"\d+차|차수미상", tag[1]) else ""
            namespace = builder.KY_OFFSET if source_tag.startswith("건양대") else 0
            current["questions"].append({
                "_source_id": qid,
                "_source_qid": qid,
                "_source_topic_id": topic_id,
                "_source_span": [index + 1, pos + 2],
                "_source_year": int(year),
                "_original_tag": source_tag,
                "_original_body_sha256": hashlib.sha256(original.encode("utf-8")).hexdigest(),
                "_explicit_prof": explicit_prof,
                "no": header["no"], "stem": stem,
                "year": int(year) % 100 + namespace, "round": exam_round,
                "inferred": "임의분류" in header["flags"].replace(" ", ""),
                "xref": None, "body": body,
            })
            index = pos + 1
        elif QUESTION_HEAD.match(line):
            raise ValueError(f"Question without organizer delimiters at line {index + 1}")
        index += 1
    return chapters


def parse_bank(text, builder):
    lines = text.splitlines()
    if any(line.lstrip().startswith("<!-- question qid=") for line in lines):
        chapters = marked_chapters(lines, builder)
    else:
        chapters = builder.parse_chapters(text)
        heads = [i for i, line in enumerate(lines, 1) if QUESTION_HEAD.match(line)]
        matched = [i for i in heads if builder.Q_RE.fullmatch(lines[i - 1].strip())]
        questions = [q for ch in chapters for q in ch["questions"]]
        if len(heads) != len(matched) or len(matched) != len(questions):
            unsupported = sorted(set(heads) - set(matched))
            raise ValueError(f"Question headers did not all parse ({len(questions)}/{len(heads)}). "
                             f"Check source lines {unsupported[:12]}.")
        for question, line in zip(questions, matched, strict=True):
            question["_source_span"] = [line, line]
            heading = lines[line - 1].strip()
            match = builder.Q_RE.fullmatch(heading)
            question["_original_tag"] = re.match(r"\*\*\s*\[([^]]+)]", heading[match.end("stem"):])[1]
    questions = [q for ch in chapters for q in ch["questions"]]
    if not questions:
        raise ValueError("No exam questions parsed. Check the organizer RAW Markdown format.")
    counts = Counter(q["_source_id"] for q in questions)
    for question in questions:
        if counts[question["_source_id"]] > 1:
            question["_source_id"] += "@" + str(question["_source_span"][0])
    return chapters


def csv_rows(path):
    if path is None:
        return []
    with Path(path).open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or not {"topic", "topic_id"}.intersection(reader.fieldnames):
            raise ValueError("Question CSV has no topic or topic_id column.")
        return list(reader)


def topic_ids(value):
    if not value:
        return []
    if value.lstrip().startswith("["):
        values = json.loads(value)
        if not isinstance(values, list) or not all(isinstance(v, str) for v in values):
            raise ValueError("topic_candidates must contain topic identifiers.")
        return values
    return [item.strip() for item in re.split(r"[|;]", value) if item.strip()]


def source_year(value):
    match = re.search(r"\d{2,4}", value or "")
    return int(match[0]) % 100 if match else 0


def find_csv(raw):
    options = [raw.with_name(raw.stem + "_questions.csv")]
    for suffix in ("_전연도_rawmd", "_연도순_rawmd", "_rawmd"):
        if raw.stem.endswith(suffix):
            options.append(raw.with_name(raw.stem.removesuffix(suffix) + "_questions.csv"))
    existing = list(dict.fromkeys(path for path in options if path.is_file()))
    return existing[0] if len(existing) == 1 else None


def index_bank(raw, csv_path=None):
    import build_yama_docx as builder

    raw = Path(raw).resolve()
    csv_path = Path(csv_path).resolve() if csv_path else find_csv(raw)
    raw_bytes = raw.read_bytes()
    raw_text = raw_bytes.decode("utf-8-sig")
    chapters = parse_bank(raw_text, builder)
    rows = csv_rows(csv_path)
    by_qid, by_no, by_topic = defaultdict(list), defaultdict(list), defaultdict(set)
    for row in rows:
        if row.get("qid"):
            by_qid[row["qid"], row.get("topic_id")].append(row)
        year = source_year(row.get("exam_year") or row.get("year") or row.get("tag"))
        by_no[year, (row.get("src_no") or row.get("no") or "").strip()].append(row)
        if row.get("topic") and row.get("topic_id"):
            by_topic[folded(row["topic"])].add(row["topic_id"])
    counts = Counter()
    for chapter in chapters:
        for q in chapter["questions"]:
            if q.get("_source_qid"):
                matches = by_qid[q["_source_qid"], q["_source_topic_id"]]
                by_topic[folded(chapter["topic"])].add(q["_source_topic_id"])
            else:
                matches = [row for row in by_no[q["year"], q["no"]]
                           if folded(display_markup(row.get("stem") or "")) == folded(display_markup(q["stem"]))]
            if len(matches) == 1:
                row = matches[0]
                q["_csv"] = {
                    "qid": row.get("qid"), "topic": row.get("topic"), "topic_id": row.get("topic_id"),
                    "topic_candidates": topic_ids(row.get("topic_candidates")),
                    "prof_past": row.get("prof_past"), "prof_evidence": row.get("prof_evidence"),
                    "inferred": (row.get("inferred") or "").lower() in {"true", "1", "yes"},
                }
                if row.get("exam_round") and not q["round"]:
                    q["round"] = row["exam_round"]
                counts["linked"] += 1
            elif matches:
                counts["ambiguous"] += 1
            else:
                counts["unlinked"] += 1
    return {
        "schema": SCHEMA, "raw_name": raw.name, "raw_sha256": hashlib.sha256(raw_bytes).hexdigest(),
        "csv_name": csv_path.name if csv_path else None, "csv_sha256": sha(csv_path) if csv_path else None,
        "csv_links": dict(counts), "csv_has_candidates": bool(rows and "topic_candidates" in rows[0]),
        "topics": {key: sorted(ids) for key, ids in by_topic.items()},
        "table": builder.parse_matching_table(raw_text), "chapters": chapters,
    }


def index_csv_path(bank, raw, explicit=None):
    if explicit:
        return Path(explicit)
    name = bank["csv_name"]
    return Path(raw).resolve().parent / name if name else None


def source_index(path, raw=None, csv_path=None):
    bank = read_json(path)
    if bank.get("schema") != SCHEMA:
        raise ValueError("Unsupported question index. Rebuild the local index from the source bank.")
    if raw is not None:
        if sha(raw) != bank["raw_sha256"]:
            raise ValueError("Question bank changed. Rebuild the local index and review the new selection.")
        csv_file = index_csv_path(bank, raw, csv_path)
        if csv_file and (not csv_file.is_file() or sha(csv_file) != bank["csv_sha256"]):
            raise ValueError("Question CSV changed or is unavailable. Rebuild the index or supply its original --csv.")
    return bank


def number_range(value):
    match = re.fullmatch(r"(\d{2}|\d{4}):(\d+)-(\d+)", value)
    if not match or int(match[2]) > int(match[3]):
        raise argparse.ArgumentTypeError("Use YYYY:START-END, e.g. 2025:11-18.")
    return int(match[1]) % 100, int(match[2]), int(match[3])


def candidate_packet(bank, topic, aliases, terms, ranges, include_types, offset, limit):
    topic_keys = {folded(s) for s in [topic, *aliases]}
    ids = {tid for key in topic_keys for tid in bank["topics"].get(key, [])}
    ids.update(tid for values in bank["topics"].values() for tid in values if folded(tid) in topic_keys)
    alias_values = [s for s in [topic, *aliases] if s]
    reason_counts, candidates = Counter(), []
    for chapter in bank["chapters"]:
        assigned = folded(chapter["topic"]) in topic_keys
        for question in chapter["questions"]:
            hint = question.get("_csv", {})
            question_topic = question.get("_source_topic_id") or hint.get("topic_id")
            in_scope = assigned or question_topic in ids
            reasons = ["assigned_topic"] if in_scope else []
            hits = hint.get("topic_candidates", [])
            if any(hit in ids or folded(hit) in topic_keys for hit in hits):
                reasons.append("csv_candidate")
            if in_scope and hits and not any(hit in ids or folded(hit) in topic_keys for hit in hits):
                reasons.append("csv_assignment_conflict")
            content = "\n".join([question["stem"], *question["body"]])
            compact = folded(content)
            matches = [term for term in terms if folded(term) and folded(term) in compact]
            if matches:
                reasons.append("lecture_term")
            if include_types and TYPE_CUES.search(content):
                reasons.append("written_question")
            if question["no"].isdigit() and any(question["year"] == year and lo <= int(question["no"]) <= hi
                                                   for year, lo, hi in ranges):
                reasons.append("exam_number_range")
            if not reasons:
                continue
            if question["inferred"] or hint.get("inferred"):
                reasons.append("inferred")
            reason_counts.update(reasons)
            candidates.append({
                "id": question["_source_id"], "source_qid": question.get("_source_qid"),
                "chapter": chapter["topic"], "source_lines": question["_source_span"],
                "source_tag": question["_original_tag"], "number": question["no"],
                "year": question["year"], "round": question["round"],
                "reasons": reasons, "lecture_terms": matches,
                "source_prof": question.get("_explicit_prof"), "csv": hint,
                "stem": question["stem"], "body": question["body"],
            })
    return {
        "raw_sha256": bank["raw_sha256"], "csv_sha256": bank["csv_sha256"], "topic": topic, "topic_aliases": alias_values,
        "csv_name": bank["csv_name"], "csv_links": bank["csv_links"],
        "csv_has_candidates": bank["csv_has_candidates"],
        "n_bank": sum(len(ch["questions"]) for ch in bank["chapters"]),
        "n_candidates": len(candidates), "reason_counts": dict(reason_counts),
        "offset": offset, "next_offset": offset + limit if offset + limit < len(candidates) else None,
        "candidates": candidates[offset:offset + limit],
    }


def select_chapters(bank, selection_path):
    selection = read_json(selection_path)
    if selection.get("raw_sha256") != bank["raw_sha256"] or selection.get("csv_sha256") != bank["csv_sha256"]:
        raise ValueError("Selection refers to a different RAW or CSV revision.")
    plan = selection.get("chapters")
    if not isinstance(plan, list) or not plan:
        raise ValueError("Selection must specify at least one chapter.")
    questions = {q["_source_id"]: (ch, q) for ch in bank["chapters"] for q in ch["questions"]}
    taken, seen_sources, selected = set(), set(), []
    for entry in plan:
        topic = entry["topic"].strip()
        current = entry["current"].strip()
        source_ids = entry["source_ids"]
        if not topic or not current or not isinstance(source_ids, list) or not source_ids:
            raise ValueError("Each selected chapter needs a topic, current lecturer and source IDs.")
        chapter = {"topic": topic, "past": None, "current": current, "questions": [], "_selected_by_id": True}
        for source_id in source_ids:
            if source_id in taken or source_id not in questions:
                raise ValueError(f"Unknown or repeated selected question ID: {source_id}")
            taken.add(source_id)
            source_chapter, question = questions[source_id]
            source_key = (question.get("_source_qid"), question.get("_original_body_sha256"))
            if source_key[0] and source_key in seen_sources:
                raise ValueError(f"Same organizer question selected under multiple topic placements: {source_id}")
            seen_sources.add(source_key)
            copied = deepcopy(question)
            if source_chapter["topic"] != topic:
                copied["_pulled_from"] = source_chapter["topic"]
            chapter["questions"].append(copied)
        selected.append(chapter)
    return selected


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    index = commands.add_parser("index")
    index.add_argument("--rawmd", required=True)
    index.add_argument("--csv")
    index.add_argument("--out", required=True)
    packet = commands.add_parser("packet")
    packet.add_argument("--index", required=True)
    packet.add_argument("--topic", required=True, help="Exact topic name or topic ID")
    packet.add_argument("--alias", action="append", default=[])
    packet.add_argument("--term", action="append", default=[])
    packet.add_argument("--number-range", type=number_range, action="append", default=[])
    packet.add_argument("--include-written", action="store_true")
    packet.add_argument("--offset", type=int, default=0)
    packet.add_argument("--limit", type=int, default=20)
    packet.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == "index":
            raw = Path(args.rawmd).resolve()
            csv_path = Path(args.csv).resolve() if args.csv else find_csv(raw)
            check_output_paths([args.out], [raw, csv_path])
            bank = index_bank(raw, csv_path)
            write_json(args.out, bank)
            result = {"out": args.out, "raw_sha256": bank["raw_sha256"], "csv_sha256": bank["csv_sha256"],
                      "csv_links": bank["csv_links"], "csv_has_candidates": bank["csv_has_candidates"],
                      "chapters": [{"topic": ch["topic"], "current": ch["current"], "n": len(ch["questions"])}
                                   for ch in bank["chapters"] if ch["questions"]]}
        else:
            if args.offset < 0 or args.limit < 1:
                raise ValueError("Offset must be nonnegative and limit must be positive.")
            check_output_paths([args.out], [args.index])
            bank = source_index(args.index)
            data = candidate_packet(bank, args.topic, args.alias, args.term, args.number_range,
                                    args.include_written, args.offset, args.limit)
            write_json(args.out, data)
            result = {"out": args.out, **{key: value for key, value in data.items() if key != "candidates"}}
    except (OSError, ValueError, TypeError, KeyError) as exc:
        parser.exit(2, f"{exc}\n")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
