"""Conservative, display-only line breaks; never reconstruct missing wording."""
import re
from itertools import groupby


BR_RE = re.compile(r"<br\s*/?\s*>", re.I)
SERIES = ("가나다라마바사", "ㄱㄴㄷㄹㅁㅂㅅ")
STATEMENT_RE = re.compile(r"(?<!\S)(?:\(([가나다라마바사ㄱㄴㄷㄹㅁㅂㅅ])\)|([가나다라마바사ㄱㄴㄷㄹㅁㅂㅅ])[).])(?=\s)")


def _statement_lines(line):
    markers = list(STATEMENT_RE.finditer(line))
    if len(markers) < 2 or "![[" in line:
        return [line]
    labels = "".join(m.group(1) or m.group(2) for m in markers)
    if not any(series.startswith(labels) for series in SERIES):
        return [line]
    prefix = line[:markers[0].start()].strip()
    if prefix and not prefix.endswith(("?", ":", "？", "：")):
        return [line]
    # Connective phrases may be references, not statements. A short statement
    # such as '다) 수혈' is still valid in a clear consecutive list.
    ends = [m.start() for m in markers[1:]] + [len(line)]
    contents = [line[m.end():end].strip() for m, end in zip(markers, ends)]
    if any(not text for text in contents):
        return [line]
    if any(text.endswith(("과", "와", "및", "또는", "혹은", "중", ",", ";"))
           for text in contents[:-1]):
        return [line]
    return ([prefix] if prefix else []) + [
        line[m.start():end].strip() for m, end in zip(markers, ends)]


def display_lines(text):
    """Restore explicit HTML breaks and split only clear, consecutive statements."""
    return [part for line in BR_RE.sub("\n", text).split("\n")
            for part in _statement_lines(line)]


def statement_blocks(lines):
    """Group complete consecutive statement lines; leave all other lines in place."""
    lines = list(lines)
    index = 0
    while index < len(lines):
        end = index
        labels, contents = "", []
        while end < len(lines):
            line = lines[end].strip()
            markers = list(STATEMENT_RE.finditer(line))
            if len(markers) != 1 or markers[0].start() != 0 or "![[" in line:
                break
            marker = markers[0]
            labels += marker.group(1) or marker.group(2)
            contents.append(line[marker.end():].strip())
            end += 1
        valid = (len(labels) >= 2 and any(s.startswith(labels) for s in SERIES)
                 and all(contents) and not any(
                     text.endswith(("과", "와", "및", "또는", "혹은", "중", ",", ";"))
                     for text in contents[:-1]))
        if valid:
            yield True, lines[index:end]
            index = end
        else:
            yield False, [lines[index]]
            index += 1


def stem_parts(lines):
    """Keep the question lead above its history and move following boxes below it."""
    blocks = list(statement_blocks(lines))
    first_box = next((i for i, (boxed, _) in enumerate(blocks) if boxed), len(blocks))
    lead = [line for _, block in blocks[:first_box] for line in block]
    if not any(line.strip() for line in lead):
        return [line for _, block in blocks for line in block], []
    return lead, blocks[first_box:]


def year_statement_blocks(entries):
    """Only adjacent entries from the same year can share a statement box."""
    for year, group in groupby(entries, key=lambda entry: entry[0]):
        for boxed, lines in statement_blocks(line for _, line in group):
            yield year, boxed, lines


def display_text_preserved(source, lines):
    """Independent invariant: only HTML break markup and whitespace may differ."""
    original = re.sub(r"\s", "", BR_RE.sub("", source))
    displayed = re.sub(r"\s", "", "".join(lines))
    return original == displayed
