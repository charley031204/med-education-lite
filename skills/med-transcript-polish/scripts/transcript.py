"""Split a Whisper transcript, retrieve lecture passages, and assemble a reviewed MD."""

import argparse
import json
import re
from bisect import bisect_left
from collections import Counter, defaultdict
from pathlib import Path


TIMECODE = r"\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d{1,3})?"
TIME_MARKER = re.compile(
    rf"(?m)^[ \t]*(?:\[{TIMECODE}[ \t]*(?:-->|[–—-])[ \t]*{TIMECODE}\]"
    rf"|{TIMECODE}[ \t]+-->[ \t]+{TIMECODE}(?=[ \t]*\r?$))[ \t]*(?:\r?\n)?"
)
PAGE_MARKER = re.compile(r"(?m)^[ \t]*<!--[ \t]*page[ \t]+([1-9]\d*)[ \t]*-->[ \t]*$", re.I)
PAGE_TAG = re.compile(r"\[Page [1-9]\d*\??\]")
JOIN_MARKER = re.compile(r"<!-- join: (start|space|paragraph|tight|empty) -->")
SKIP_NOTE = "(이 페이지는 교수님이 넘어가심)"
PART_1 = "## PART 1 — 교정 스크립트 (순수 정제)"
PART_2 = "## PART 2 — 강의 정리본 (녹음본 요약)"


def read_text(path, encoding="utf-8-sig"):
    with path.open(encoding=encoding, newline="") as handle:
        return handle.read()


def write_new(path, body):
    with path.open("x", encoding="utf-8", newline="") as handle:
        handle.write(body)


def split_text(body, max_chars):
    """Keep every source character and prefer boundaries after sentences."""
    markers = list(TIME_MARKER.finditer(body))
    marker_starts = [marker.start() for marker in markers]

    def containing_timestamp(position):
        prior = bisect_left(marker_starts, position) - 1
        if prior >= 0 and position <= markers[prior].end():
            return markers[prior]
        return None

    start = 0
    while start < len(body):
        stop = min(start + max_chars, len(body))
        if stop == len(body):
            yield body[start:]
            break

        # Do not cut inside a timestamp or between CR and LF.
        marker = containing_timestamp(stop)
        if marker:
            stop = marker.start()
        if body[stop - 1:stop + 1] == "\r\n":
            stop -= 1

        lower = min(start + max_chars // 2, stop)
        tail = body[lower:stop]
        sentence_ends = list(re.finditer(r"[.!?。？！…][ \t\r\n]+", tail))
        spaces = list(re.finditer(r"\s+", tail))
        for boundaries in (sentence_ends, spaces):
            for boundary in reversed(boundaries):
                candidate = lower + boundary.end()
                # A boundary after a timing header alone would obscure its context.
                if containing_timestamp(candidate):
                    continue
                if candidate < len(body) and body[candidate - 1:candidate + 1] == "\r\n":
                    continue
                stop = candidate
                break
            else:
                continue
            break
        if stop <= start:
            raise ValueError("Chunk size cannot accommodate a timestamp.")
        yield body[start:stop]
        start = stop


def split_command(args):
    source = read_text(args.source, args.encoding)
    if not TIME_MARKER.sub("", source).strip():
        raise ValueError("The transcript contains no utterance text.")
    args.work.mkdir(parents=True, exist_ok=False)
    for folder in ("raw", "polished", "topics", "pages"):
        (args.work / folder).mkdir()
    count = 0
    for count, body in enumerate(split_text(source, args.max_chars), 1):
        write_new(args.work / "raw" / f"{count:04d}.txt", body)
    return {"chunks": count, "source_chars": len(source), "work": str(args.work)}


def lecture_pages(path):
    body = read_text(path).replace("\r\n", "\n").replace("\r", "\n")
    markers = list(PAGE_MARKER.finditer(body))
    if not markers:
        raise ValueError("No <!-- page N --> markers in the lecture MD.")
    pages = {}
    for index, marker in enumerate(markers):
        page = int(marker.group(1))
        if page in pages:
            raise ValueError(f"Duplicate lecture page marker: {page}")
        end = markers[index + 1].start() if index + 1 < len(markers) else len(body)
        pages[page] = body[marker.end():end].strip()
    return pages


def match_expression(term):
    # Latin identifiers must match whole terms. Hangul can carry attached particles.
    before = r"(?<![A-Za-z0-9])" if re.match(r"[A-Za-z0-9]", term) else ""
    after = r"(?![A-Za-z0-9])" if re.search(r"[A-Za-z0-9]$", term) else ""
    words = r"\s+".join(re.escape(word) for word in term.split())
    return re.compile(before + words + after, re.I)


def candidates(pages, terms, limit, max_chars):
    unique = {}
    for term in terms:
        term = " ".join(term.split())
        if term:
            unique.setdefault(term.casefold(), term)
    if not unique:
        raise ValueError("Provide at least one nonempty search term.")
    patterns = {term: match_expression(term) for term in unique.values()}
    hits = {}
    counts = Counter()
    for page, body in pages.items():
        visible = re.sub(r"<!--.*?-->|!\[\[[^\n]*?\]\]", "", body, flags=re.S)
        matches = {term: pattern.search(visible) for term, pattern in patterns.items()}
        matches = {term: match for term, match in matches.items() if match}
        if matches:
            hits[page] = (visible, matches)
            counts.update(matches.keys())
    scores = {
        page: (len(matches), sum(len(pages) + 1 - counts[term] for term in matches))
        for page, (_, matches) in hits.items()
    }
    order = sorted(hits, key=scores.__getitem__, reverse=True)
    ties = [page for page in order if scores[page] == scores[order[0]]] if order else []
    results = []
    for page in order[:limit]:
        visible, matches = hits[page]
        position = min(match.start() for match in matches.values())
        start = max(0, min(position - max_chars // 4, len(visible) - max_chars))
        results.append({
            "page": page,
            "matched_terms": list(matches),
            "excerpt": visible[start:start + max_chars],
        })
    return {
        "total_matches": len(hits),
        "top_score_page_count": len(ties),
        "candidates": results,
    }


def lecture_command(args):
    pages = lecture_pages(args.lecture)
    if args.page is None:
        return candidates(pages, args.terms, args.limit, args.max_chars)
    if args.page not in pages:
        raise ValueError(f"Lecture page {args.page} does not exist.")
    body = pages[args.page]
    stop = min(len(body), args.offset + args.max_chars)
    if args.offset > len(body):
        raise ValueError("The page offset exceeds the page length.")
    return {
        "page": args.page,
        "body": body[args.offset:stop],
        "next_offset": stop if stop < len(body) else None,
    }


def reviewed_piece(path, original, first):
    marker, _, body = read_text(path).partition("\n")
    join = JOIN_MARKER.fullmatch(marker.rstrip("\r"))
    if not join:
        raise ValueError(f"Missing or invalid join marker in {path.name}.")
    mode = join.group(1)
    body = body.replace("\r\n", "\n").replace("\r", "\n").strip()
    raw_is_empty = not TIME_MARKER.sub("", original).strip()
    if mode == "empty" and raw_is_empty and not body:
        return mode, ""
    if not body or raw_is_empty or mode == "empty" or (mode == "start") != first:
        raise ValueError(f"Unexpected {mode} join or empty utterance body in {path.name}.")
    if TIME_MARKER.search(body) or any(
        PAGE_TAG.fullmatch(line.strip()) or line.strip() in {PART_1, PART_2, SKIP_NOTE}
        for line in body.split("\n")
    ):
        raise ValueError(f"Unexpected timing, page annotation or PART header in {path.name}.")
    return mode, body


def annotate(path, body, mode, lecture, seen):
    lines = body.split("\n")
    insertions = defaultdict(list)
    spoken_lines = set()
    counts = Counter()
    header, _, rows = read_text(path).partition("\n")
    header = header.rstrip("\r")
    unresolved = header.startswith("# unresolved\t") and bool(header.split("\t", 1)[1].strip())
    if header != "# reviewed" and not unresolved:
        raise ValueError(f"{path.name} needs a page review status on its first line.")
    for number, row in enumerate(rows.splitlines(), 2):
        if not row.strip():
            continue
        fields = row.split("\t", 3)
        if len(fields) != 4 or not fields[0].isdigit() or not fields[1].isdigit():
            raise ValueError(f"{path.name}:{number} needs line, page, status and evidence.")
        line, page = map(int, fields[:2])
        status, evidence = fields[2:]
        if status not in {"confirmed", "uncertain", "skipped"} or not evidence.strip():
            raise ValueError(f"{path.name}:{number} needs a valid status and evidence.")
        if not 1 <= line <= len(lines) or not lines[line - 1].strip():
            raise ValueError(f"{path.name}:{number} needs an existing nonempty body line.")
        if line == 1 and mode in {"space", "tight", "empty"}:
            raise ValueError(f"{path.name}:{number} targets a continuing or empty body.")
        if page not in lecture or page in seen:
            raise ValueError(f"{path.name}:{number} references an unknown or repeated page: {page}.")
        if line in spoken_lines:
            raise ValueError(f"{path.name}:{number} follows a spoken page annotation on the same line.")
        if status != "skipped":
            spoken_lines.add(line)
        seen.add(page)
        counts[status] += 1
        suffix = "?" if status == "uncertain" else ""
        insertions[line].append(f"[Page {page}{suffix}]")
        if status == "skipped":
            insertions[line].append(SKIP_NOTE)

    tagged = []
    for index, line in enumerate(lines, 1):
        tagged.extend(insertions.get(index, ()))
        tagged.append(line)
    return "\n".join(tagged), counts, unresolved


def assemble_command(args):
    originals = list((args.work / "raw").glob("*.txt"))
    if not originals:
        raise ValueError("No input chunks in the work folder.")
    expected = [f"{index:04d}" for index in range(1, len(originals) + 1)]
    if {path.stem for path in originals} != set(expected):
        raise ValueError("Input chunk numbers are missing or unexpected.")
    polished = {path.stem for path in (args.work / "polished").glob("*.md")}
    if polished != set(expected):
        raise ValueError(f"Missing or unexpected polished chunks: {sorted(polished ^ set(expected))}")
    chunks = [read_text(args.work / "raw" / f"{key}.txt", "utf-8") for key in expected]
    if "".join(chunks) != read_text(args.source, args.encoding):
        raise ValueError("The raw chunks do not reconstruct the original transcript.")

    page_files = {path.stem for path in (args.work / "pages").glob("*.tsv")}
    use_tags = args.lecture is not None and not args.no_page_tags
    if use_tags and page_files.difference(expected):
        raise ValueError("Page review files contain unexpected chunk numbers.")
    if page_files and not args.lecture and not args.no_page_tags:
        raise ValueError("Provide the lecture MD to use reviewed page annotations.")
    lecture = lecture_pages(args.lecture) if use_tags else {}
    seen = set()
    counts = Counter()
    unresolved = []
    plain = []
    tagged = []
    separators = {"start": "", "space": " ", "paragraph": "\n\n", "tight": ""}
    for key, original in zip(expected, chunks):
        mode, body = reviewed_piece(args.work / "polished" / f"{key}.md", original, not plain)
        if mode == "empty":
            if use_tags and key in page_files:
                raise ValueError(f"Empty chunk {key} cannot carry page annotations.")
            continue
        plain.append(separators[mode] + body)
        if use_tags:
            if key not in page_files:
                raise ValueError(f"Missing page review for chunk {key}.")
            body, added, open_match = annotate(
                args.work / "pages" / f"{key}.tsv", body, mode, lecture, seen
            )
            counts.update(added)
            if open_match or added["uncertain"]:
                unresolved.append(key)
        tagged.append(separators[mode] + body)

    joined = "".join(tagged)
    untagged = "\n".join(
        line for line in joined.split("\n") if not PAGE_TAG.fullmatch(line) and line != SKIP_NOTE
    )
    if not joined or untagged != "".join(plain):
        raise ValueError("Page annotation altered the utterance text.")

    part2 = read_text(args.part2).replace("\r\n", "\n").replace("\r", "\n").strip()
    if not part2 or PAGE_TAG.search(part2) or any(line in {PART_1, PART_2} for line in part2.split("\n")):
        raise ValueError("PART 2 needs a body without page tags or duplicate PART headers.")
    title = args.title.strip()
    if not title or "\n" in title or "\r" in title:
        raise ValueError("The title must be one nonempty line.")
    write_new(args.output, f"# {title}\n\n---\n\n{PART_1}\n\n{joined}\n\n---\n\n{PART_2}\n\n{part2}\n")
    page_review = "off"
    if use_tags:
        if counts["confirmed"] + counts["uncertain"] == 0:
            page_review = "no-matches"
        elif unresolved:
            page_review = "partial"
        else:
            page_review = "reviewed"
    return {
        "chunks": len(chunks),
        "page_tagging": use_tags,
        "page_review": page_review,
        "unresolved_chunk_count": len(unresolved),
        "first_unresolved_chunks": unresolved[:10],
        "page_annotations": dict(counts),
        "output": str(args.output),
    }


def integer_at_least(minimum):
    def convert(value):
        number = int(value)
        if number < minimum:
            raise argparse.ArgumentTypeError(f"Must be at least {minimum}.")
        return number

    return convert


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    subcommands = parser.add_subparsers(dest="command", required=True)
    split = subcommands.add_parser("split", help="Write exact, bounded transcript chunks.")
    split.add_argument("source", type=Path)
    split.add_argument("work", type=Path)
    split.add_argument("--max-chars", type=integer_at_least(128), default=6500)
    split.add_argument("--encoding", default="utf-8-sig")
    split.set_defaults(run=split_command)

    lecture = subcommands.add_parser("lecture", help="Find candidates or read one known page.")
    lecture.add_argument("lecture", type=Path)
    selection = lecture.add_mutually_exclusive_group(required=True)
    selection.add_argument("--terms", nargs="+")
    selection.add_argument("--page", type=integer_at_least(1))
    lecture.add_argument("--offset", type=integer_at_least(0), default=0)
    lecture.add_argument("--limit", type=integer_at_least(1), default=5)
    lecture.add_argument("--max-chars", type=integer_at_least(32), default=1000)
    lecture.set_defaults(run=lecture_command)

    assemble = subcommands.add_parser("assemble", help="Validate and assemble PART 1 and PART 2.")
    assemble.add_argument("source", type=Path)
    assemble.add_argument("work", type=Path)
    assemble.add_argument("part2", type=Path)
    assemble.add_argument("output", type=Path)
    assemble.add_argument("--title", required=True)
    assemble.add_argument("--lecture", type=Path)
    assemble.add_argument("--no-page-tags", action="store_true")
    assemble.add_argument("--encoding", default="utf-8-sig")
    assemble.set_defaults(run=assemble_command)

    args = parser.parse_args()
    try:
        result = args.run(args)
    except (OSError, ValueError, LookupError) as error:
        parser.exit(1, f"{error}\n")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
