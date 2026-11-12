#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_yama_docx.py — exam-yama-organizer rawmd에서 챕터를 추출해
숙방(해환) 양식의 야마 Word(.docx)를 생성한다.

사용:
  python build_yama_docx.py --rawmd RAWMD.md --topic "현미경검사" \
      [--topic ...] [--out DIR] [--subject 과목] [--exam "1차"] \
      [--comments comments.txt] [--extras extras.json] [--ver 1] [--outname NAME]

출력: docx 파일 + stdout으로 검증 리포트 JSON
"""
import argparse
import hashlib
import json
import re
import sys
import unicodedata
from copy import deepcopy
from types import SimpleNamespace
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor
from docx.text.paragraph import Paragraph

from yama_layout import YamaLayout
from yama_photo_word import history_entry, prepare_review
from yama_bank import check_output_paths, find_csv, index_csv_path, parse_bank, professor_in_current, select_chapters, source_index, write_json
from yama_readability import display_lines, display_text_preserved, statement_blocks, stem_parts, year_statement_blocks

# 태그 변형 수용: [25본1_1차] · [22본1_차수미상] · [18본1_신장학-약리학] · [연도미상_교수명]
# 차수 표기가 없거나 연도가 미상인 원본이 실제로 존재하므로 좁게 잡으면 문항이 조용히 누락된다.
Q_RE = re.compile(
    r"^\*\*Q\.\s*(?:(?P<no>(?:주관식\s*\d+|[0-9무번호외~\-–]+))\.\s*)?(?P<stem>.*?)\*\*\s*"
    r"\[(?:(?P<src>건양대)?(?P<year>\d{2})(?:본1)?_(?P<round>[^\]]+)|연도미상_(?P<who>[^\]]+))\]\s*(?P<marker>.*)$"
)
# 건양대 기출은 충남대와 별개 학교이므로 연도 네임스페이스를 분리한다.
# -1000 오프셋: 내림차순 정렬 시 충남대 연도(25..14) 뒤에 건양대(23..16) 순으로 붙는다.
KY_OFFSET = -1000


def ytag(y):
    """연도 표시. 0 = 원본에 연도 표기가 없는 문항(연도미상). 음수 = 건양대 기출."""
    if y == 0:
        return "연도미상"
    if y < 0:
        return f"건양대{y - KY_OFFSET:02d}"
    return f"{y:02d}"


def yparse(v):
    """JSON 스펙의 year 값(25 · "25" · "건양대23") → 내부 연도 정수."""
    if isinstance(v, str):
        v = v.strip()
        if v.startswith("건양대"):
            return int(v[3:]) + KY_OFFSET
        if v == "연도미상":
            return 0
        return int(v)
    return int(v)
IMG_RE = re.compile(r"!\[\[([^\]|]+?)(?:\|(\d+))?\]\]")
NOTE_RE = re.compile(r"^\*\[(.+)\]\*$")
CHAPTER_RE = re.compile(r"^####\s+(.*)$")
# rawmd 임의 분류 검증 이력 주석 — 문서 렌더링·병합 비교에서 모두 제외
RAWMETA_RE = re.compile(r"^※\s*\[(?:검증완료|재배치)\]")
# rawmd 내 명시적 교차참조: "→ [25본1_1차]14번과 완전히 동일한 문항"
XREF_RE = re.compile(r"→\s*\[(\d{2})본1_[^\]]+\]\s*([0-9무번호외~\-–]+)번과.*동일")

GRAY = RGBColor(0x88, 0x88, 0x88)
RED = RGBColor(0xC0, 0x00, 0x00)
BLUE = RGBColor(0x0B, 0x50, 0xC0)  # 정답 표시용

CIRCLED = {c: str(i + 1) for i, c in enumerate("①②③④⑤⑥⑦⑧⑨")}
MARKER_RE = re.compile(r"^\s*(?:\(?([가나다라마바사ㄱㄴㄷㄹㅁㅂㅅ])[).]|([①②③④⑤⑥⑦⑧⑨])|\(?(\d{1,2})\))")


def choice_marker(line):
    """선지 줄의 마커를 정규형으로: '가'/'나'/... 또는 '1'/'2'/... 없으면 None."""
    m = MARKER_RE.match(line)
    if not m:
        return None
    if m.group(1):
        return m.group(1)
    if m.group(2):
        return CIRCLED[m.group(2)]
    return str(int(m.group(3)))


def norm_correct(tok):
    tok = str(tok).strip().strip("()）.")
    return CIRCLED.get(tok, tok)


# ---------------------------------------------------------------- parsing
def parse_matching_table(text):
    """## 주제–교수 매칭 표 → {주제: {past, current, n, verdict}}"""
    rows, in_tbl = {}, False
    for line in text.splitlines():
        if line.startswith("## 주제"):
            in_tbl = True
            continue
        if in_tbl:
            if line.startswith("## "):
                break
            if not line.strip().startswith("|"):
                continue
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            if len(cells) < 6 or cells[1] in ("주제", ""):
                continue
            if re.fullmatch(r"[-: ]+", cells[0] or "-"):
                continue
            rows[cells[1]] = {
                "past": cells[2],
                "current": cells[3],
                "n": cells[4],
                "verdict": cells[5],
            }
    return rows


def parse_chapter_title(title):
    """#### 헤더 → (topic, past_prof, current_prof)"""
    t = title.strip()
    m = re.match(
        r"^⚠️?\s*담당\s*변경\s*—\s*(.+?)\s*—\s*기출\s*연도\s*[::]\s*(.+?)\s*/\s*현재\s*[::]\s*(.+)$",
        t,
    )
    if m:
        return m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
    if "—" in t:
        topic, prof = t.rsplit("—", 1)
        return topic.strip(), None, prof.strip()
    return t, None, None


def parse_chapters(text):
    """rawmd 전체 → [{topic, past, current, questions:[...]}] (파일 순서 유지)"""
    chapters = []
    cur, curq = None, None
    for raw in text.splitlines():
        if re.match(r"^#{1,3}\s", raw):
            cur, curq = None, None
            continue
        m = CHAPTER_RE.match(raw)
        if m:
            topic, past, current = parse_chapter_title(m.group(1))
            cur = {"topic": topic, "past": past, "current": current, "questions": []}
            chapters.append(cur)
            curq = None
            continue
        if cur is None:
            continue
        qm = Q_RE.match(raw.strip())
        if qm:
            marker = qm.group("marker") or ""
            xm = XREF_RE.search(marker)
            curq = {
                "_source_id": f"raw-c{len(chapters):04d}-q{len(cur['questions']) + 1:05d}",
                "no": (qm.group("no") or "").strip(),
                "stem": qm.group("stem").strip(),
                "year": (int(qm.group("year")) + (KY_OFFSET if qm.group("src") else 0))
                        if qm.group("year") else 0,
                "round": (qm.group("round") or qm.group("who") or "미상").strip(),
                "inferred": "임의 분류" in marker,
                "xref": (int(xm.group(1)), xm.group(2)) if xm else None,
                "body": [],
            }
            cur["questions"].append(curq)
            continue
        if curq is not None:
            # rawmd 관리용 메타 주석(SKILL.md 5-4)은 시험지에 싣지 않는다
            if RAWMETA_RE.match(raw.strip()):
                continue
            curq["body"].append(raw.rstrip())
    for ch in chapters:
        for q in ch["questions"]:
            # 챕터 끝의 수평선(---)이 마지막 문항의 선지로 딸려 들어가지 않게 제거
            while q["body"] and (not q["body"][-1].strip()
                                 or set(q["body"][-1].strip()) <= {"-", "*", "_"}
                                 and len(q["body"][-1].strip()) >= 3):
                q["body"].pop()
            while q["body"] and not q["body"][0].strip():
                q["body"].pop(0)
    return chapters


# ---------------------------------------------------------------- merging
ANNOT_RE = re.compile(
    r"\(\s*(?:[^()]*야마|(?:\d+\s*점\s*/\s*)?[가-힣]{2,4}(?:\s*(?:교수님?|study question))?|배점\s*\?\s*/\s*[가-힣]{2,4})\s*\)"
)


def seq_no(no):
    """정렬용 시험지 문항 순번. '19'→19, 'p002-q19-2'→19, '주관식16'→16."""
    m = re.search(r"q(\d+)", no) or re.search(r"(\d+)", no)
    return int(m.group(1)) if m else 10 ** 6


def year_no_key(q):
    """연도 내림차순 → 그 연도 안에서 문항 번호 오름차순."""
    y = max(q["years"])
    pref = ytag(y) + "-"
    nos = [n[len(pref):] for n in q["nos"] if n.startswith(pref)]
    return (-y, min((seq_no(n), n) for n in nos) if nos else (10 ** 6, ""))


def norm_stem(s):
    s = unicodedata.normalize("NFC", s)
    s = ANNOT_RE.sub("", s)
    s = re.sub(r"\s+", "", s)
    return s.rstrip(".?!。 ")


def norm_body(lines):
    keep = [l for l in lines if l.strip() and not IMG_RE.search(l)]
    return re.sub(r"\s+", "", unicodedata.normalize("NFC", "\n".join(keep)))


CHOICE_LINE_RE = re.compile(r"^\s*(?:[가나다라마바사]\)|[①②③④⑤⑥⑦⑧⑨]|\(?\d{1,2}\))\s*")


def norm_choice(line):
    return re.sub(r"\s+", "", unicodedata.normalize("NFC", CHOICE_LINE_RE.sub("", line)))


def _bigrams(s):
    return {s[i:i + 2] for i in range(len(s) - 1)}


GENERIC_WORDS_RE = re.compile(
    r"다음|중에서|중|설명|대한|대해|관한|옳은|올바른|올바르지|틀린|않은|아닌|것은|것을|"
    r"모두|고르시오|고르면|고르는|선지|보기|쓰시오|적으시오|서술하시오|정의하시오|"
    r"문항|문제|미복기|무엇|의미|간략히|간단히"
)


def _is_specific_stem(norm):
    """문두에 일반 지시어를 제외한 내용어가 충분히 남는지 — 남으면 같은 문두=같은 문제."""
    residual = GENERIC_WORDS_RE.sub("", norm)
    residual = re.sub(r"[을를이가은는의와과로에서및등\?\.\,\(\)0-9]+", "", residual)
    return len(residual) >= 4


def resolve_image(rawmd_dir, reference):
    """명시 상대경로 우선, Obsidian의 형제 이미지 폴더 파일명 참조도 지원."""
    root = Path(rawmd_dir)
    candidates = [root / reference, root / "동봉 이미지" / reference,
                  root / "동봉 이미지" / Path(reference).name]
    return next((path for path in candidates if path.is_file()), candidates[0])


def merge_questions(questions, merge_map=None, rawmd_dir=None):
    """Group exact repeats and reviewed merge-map members, preserving each source history."""
    n = len(questions)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        if any(questions[k].get("_preserve_leaf") for k in (i, j)):
            return
        if any(IMG_RE.search(line) for k in (i, j) for line in questions[k]["body"]):
            return
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[max(ri, rj)] = min(ri, rj)

    # Full text, including numbering, annotations and qualifiers, must agree
    # before a question is merged without a human decision.
    seen = {}
    for i, q in enumerate(questions):
        key = (q["stem"].strip(), tuple(line.strip() for line in q["body"] if line.strip()),
               q.get("_explicit_prof"))
        if key in seen:
            union(seen[key], i)
        else:
            seen[key] = i

    # A reviewed merge map must identify one source for each member.
    refs = {}
    for i, q in enumerate(questions):
        refs.setdefault((q["year"], q["no"]), []).append(i)
    ref_idx = {key: ids[0] for key, ids in refs.items() if len(ids) == 1}
    # ③ merge_map: [{"members": [{"year": 23, "no": "14"},
    #                             {"year": 11, "stem": "western blot"}], "note": "..."}]
    #    무번호·중복번호 문항은 "stem"(문두 부분 문자열)으로 지정한다.
    map_notes = {}
    for grp in (merge_map or []):
        idxs, unresolved_members = [], []
        for m in grp.get("members", []):
            y = yparse(m["year"])
            if "no" in m and (y, str(m["no"])) in ref_idx:
                idxs.append(ref_idx[(y, str(m["no"]))])
            elif "stem" in m:
                matches = [i for i, q in enumerate(questions)
                           if q["year"] == y and m["stem"] in q["stem"]]
                if len(matches) == 1:
                    idxs.extend(matches)
                else:
                    unresolved_members.append(m)
            else:
                unresolved_members.append(m)
        if idxs and unresolved_members:
            raise ValueError(f"Merge map members are missing or ambiguous: {unresolved_members}")
        for a, b in zip(idxs, idxs[1:]):
            union(a, b)
        if idxs and grp.get("note"):
            map_notes[find(idxs[0])] = grp["note"]

    groups = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(questions[i])

    merged = []
    for root in sorted(groups):
        g = groups[root]
        rep = max(g, key=lambda q: len(q["stem"]) + sum(len(b) for b in q["body"]))
        rep_choices = {line.strip() for line in rep["body"] if line.strip()}
        seen_images, duplicates = {}, {}

        def unique_images(line, source):
            def keep_image(match):
                if len(g) == 1:
                    return match.group(0)
                reference = match.group(1)
                path = resolve_image(rawmd_dir, reference) if rawmd_dir is not None else None
                digest = hashlib.md5(path.read_bytes()).hexdigest() if path and path.is_file() else None
                key = ("md5", digest) if digest else ("unverified", reference)
                provenance = {"path": reference, "year": ytag(source["year"]),
                              "no": source["no"] or "?"}
                if key in seen_images:
                    if digest:
                        duplicates.setdefault(digest, {"md5": digest,
                            "representative": seen_images[key], "duplicates": []})["duplicates"].append(provenance)
                    return ""
                seen_images[key] = provenance
                return match.group(0)
            return IMG_RE.sub(keep_image, line)

        body = []
        for line in rep["body"]:
            cleaned = unique_images(line, rep)
            if cleaned.strip() or not line.strip():
                body.append(cleaned)
        extra = []  # 대표에 없는 타 연도 복기 선지·이미지
        for q in g:
            if q is rep:
                continue
            if q["stem"] != rep["stem"]:
                extra.append((q["year"], q["stem"]))
            for line in q["body"]:
                if not line.strip():
                    continue
                im = IMG_RE.search(line)
                if im:
                    # 비대표 문항의 이미지도 소실 없이 함께 수록
                    cleaned = unique_images(line, q).strip()
                    if cleaned:
                        extra.append((q["year"], cleaned))
                    continue
                if line.strip() not in rep_choices:
                    extra.append((q["year"], line.strip()))
                    rep_choices.add(line.strip())
        note = map_notes.get(root)
        if len(g) > 1 and not note:
            note = "여러 출제 이력의 동일 문항을 한 문항으로 수록함."
        if duplicates:
            renamed = any(item["path"] != group["representative"]["path"]
                          for group in duplicates.values() for item in group["duplicates"])
            identity = "파일명만 다를 뿐 같은 파일이며" if renamed else "동일 파일의 중복 참조이며"
            disclosure = (f"동일 MD5 이미지의 대표 1장만 수록했습니다. {identity}, "
                          "실제 시험지에서 같은 사진을 사용했는지 상류 추출에서 복제됐는지는 미확인입니다.")
            note = f"{note} {disclosure}" if note else disclosure
        merged.append(
            {
                "years": sorted({q["year"] for q in g}, reverse=True),
                "nos": [f"{ytag(q['year'])}-{q['no'] or '?'}" for q in g],
                "stem": rep["stem"],
                "body": body,
                "extra_choices": extra,
                "note": note,
                "image_dup": bool(duplicates),
                "image_duplicates": list(duplicates.values()),
                "inferred": any(q["inferred"] for q in g),
                "dup_count": len(g),
                "_source_id": rep.get("_source_id") if len(g) == 1 else None,
                "_source_questions": deepcopy(g),
            }
        )
    return merged


def scan_inferred_candidates(selected, chapters, limit=40):
    """선택하지 않은 챕터에 있는 '임의 분류' 문항 중, 선택 챕터 내용과 겹치는 것을 후보로 뽑는다.

    임의 분류는 원본 시험지에 주제 표기가 없어 자동 배정된 문항이라 오배치가 잦다.
    오배치된 문항은 그 교수 파트에서 통째로 누락되므로, 사람이 훑을 수 있는 후보 목록을
    항상 만들어 두는 편이 안전하다. 판정은 스크립트가 하지 않고 Claude/사용자가 한다.
    """
    profile = set()
    for ch in selected:
        for q in ch["questions"]:
            profile |= _bigrams(norm_stem(q["stem"]))
    if not profile:
        return []
    sel_ids = {id(ch) for ch in selected}
    cands = []
    for ch in chapters:
        if id(ch) in sel_ids:
            continue
        for q in ch["questions"]:
            if not q["inferred"]:
                continue
            bg = _bigrams(norm_stem(q["stem"]))
            if len(bg) < 4:
                continue
            score = len(bg & profile) / len(bg)
            if score >= 0.55:
                cands.append(
                    {
                        "year": q["year"],
                        "no": q["no"] or "?",
                        "stem": q["stem"][:80],
                        "current_chapter": ch["topic"],
                        "overlap": round(score, 2),
                    }
                )
    cands.sort(key=lambda c: -c["overlap"])
    return cands[:limit]


def find_merge_candidates(merged):
    """병합되지 않은 블록 중 문두가 의미상 겹칠 가능성이 있는 쌍을 리포트용으로 추출.

    짧은 문두의 2-gram이 긴 문두에 절반 이상 포함되면 후보. 부정형(않는/없는/아닌)
    방향이 서로 다르면 다른 문항일 가능성이 크므로 표시만 하고 판단은 사람이/Claude가 한다.
    """
    NEG = re.compile(r"않는|않은|없는|아닌|틀린|관련\s*없")
    cands = []
    stems = [(i, norm_stem(m["stem"])) for i, m in enumerate(merged)]
    for a in range(len(stems)):
        for b in range(a + 1, len(stems)):
            i, si = stems[a]
            j, sj = stems[b]
            if len(si) < 4 or len(sj) < 4:
                continue
            if merged[i]["years"] == merged[j]["years"] and len(merged[i]["years"]) > 1:
                continue
            short, long_ = (si, sj) if len(si) <= len(sj) else (sj, si)
            bg = _bigrams(short)
            if not bg:
                continue
            overlap = len(bg & _bigrams(long_)) / len(bg)
            if overlap >= 0.5:
                cands.append(
                    {
                        "a": {"tag": merged[i]["nos"], "stem": merged[i]["stem"][:60]},
                        "b": {"tag": merged[j]["nos"], "stem": merged[j]["stem"][:60]},
                        "overlap": round(overlap, 2),
                        "negation_mismatch": bool(NEG.search(merged[i]["stem"])) != bool(NEG.search(merged[j]["stem"])),
                    }
                )
    cands.sort(key=lambda c: -c["overlap"])
    return cands[:40]


# ---------------------------------------------------------------- docx helpers
def set_run_font(run, size=10.5, bold=False, italic=False, color=None):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), "바탕")


def keep_block(paragraphs):
    """문항 블록이 페이지 경계에서 갈라지지 않도록 keep 속성 부여."""
    for p in paragraphs:
        p.paragraph_format.keep_together = True
    for p in paragraphs[:-1]:
        p.paragraph_format.keep_with_next = True


def render_context_box(doc, box):
    """세트형 문항의 공통지문 박스.

    페이지순 배열은 한 증례를 공유하는 세트를 정의상 흩어놓는다. 그러면 '상기 질환은?'
    같은 조응형 문두가 지문 없이 남거나(고아), 더 나쁘게는 다른 연도의 증례 뒤에 붙어
    조용히 잘못된 환자에 연결된다. 문항마다 지문을 복제해 자립시킨다.
    """
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(12)
    sp.paragraph_format.space_after = Pt(0)
    sp.paragraph_format.keep_with_next = True
    set_run_font(sp.add_run(""), 2)

    tbl = doc.add_table(rows=1, cols=1)
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "8")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)
    tbl.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))

    cp = tbl.cell(0, 0).paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    cp.paragraph_format.keep_together = True
    cp.paragraph_format.keep_with_next = True
    if box.get("label"):
        set_run_font(cp.add_run(f"{box['label']}\n"), 9, italic=True, color=GRAY)
    set_run_font(cp.add_run(box["text"]), 10)


def render_statement_box(doc, lines, correct, year=None):
    """문항 내부의 가/나/다 보기. 공통지문과 구별되는 한 셀 표."""
    style_name = "Yama Statement Box"
    if style_name not in doc.styles:
        doc.styles.add_style(style_name, WD_STYLE_TYPE.TABLE)
    table = doc.add_table(rows=1, cols=1)
    table.style = style_name
    table.autofit = False
    section = doc.sections[-1]
    width = section.page_width - section.left_margin - section.right_margin
    table.columns[0].width = width
    table.cell(0, 0).width = width
    caption = OxmlElement("w:tblCaption")
    caption.set(qn("w:val"), "YamaStatementBox")
    table._tbl.tblPr.append(caption)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "6")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    table._tbl.tblPr.append(borders)
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    paragraphs = []
    for index, line in enumerate(lines):
        p = table.cell(0, 0).paragraphs[0] if index == 0 else table.cell(0, 0).add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        if year is not None:
            set_run_font(p.add_run(f"[{ytag(year)}] "), 9, color=GRAY)
        is_correct = choice_marker(line) in correct
        set_run_font(p.add_run(line.strip()), 10.5, bold=is_correct,
                     color=BLUE if is_correct else None)
        paragraphs.append(p)
    return paragraphs


def add_page_number(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    p._p.append(fld)


def px_to_cm(px):
    return px / 96.0 * 2.54


def add_image(doc, path, width_px, report_missing):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if path.exists():
        width_cm = max(8.0, min(px_to_cm(width_px) if width_px else 15.0, 15.0))
        p.add_run().add_picture(str(path), width=Cm(width_cm))
    else:
        run = p.add_run(f"[이미지 누락: {path.name}]")
        set_run_font(run, 9, italic=True, color=RED)
        report_missing.append(str(path))
    return p


# ---------------------------------------------------------------- rendering
def source_history(q):
    if q.get("_source_history"):
        return q["_source_history"]
    parts = []
    for source in sorted(q.get("_source_questions", []), key=lambda item: item["year"], reverse=True):
        parts.append(history_entry(source, ytag))
    if not parts:
        for member in q["nos"]:
            year, no = member.split("-", 1)
            parts.append(f"{year}{'' if year == '연도미상' else '년'} {no}번")
    return "※ 출제 이력: " + " · ".join(parts)


def fmt_pages(pages):
    """[3,4,5,9] → 'p.3~5, 9' — 연속 구간은 묶어서 읽기 쉽게."""
    if not pages:
        return "p.미상"
    ps = sorted(set(pages))
    parts, start, prev = [], ps[0], ps[0]
    for x in ps[1:] + [None]:
        if x == prev + 1:
            prev = x
            continue
        parts.append(str(start) if start == prev else f"{start}~{prev}")
        if x is not None:
            start = prev = x
    return "p." + ", ".join(parts)


def render_question(doc, q, rawmd_dir, missing_images, author_label=False, page_mode=False):
    paras = []
    box = q.get("_context_box")
    if box:
        render_context_box(doc, box)
    tag = "[" + ",".join(ytag(y) for y in q["years"]) + "]"
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4 if box else 14)
    p.paragraph_format.space_after = Pt(4)
    ans = q.get("_answer") or {}
    correct = {norm_correct(c) for c in ans.get("correct", [])}
    lead, stem_blocks = stem_parts(display_lines(q["stem"]))
    for index, line in enumerate(lead or [""]):
        prefix = f"{tag} " if index == 0 else "\n"
        color = BLUE if choice_marker(line) in correct else None
        set_run_font(p.add_run(prefix + line), 10.5, bold=True, color=color)
    if page_mode:
        set_run_font(p.add_run(f"  ({fmt_pages(q.get('_pages'))})"), 9, color=GRAY)
    if author_label and q.get("authors"):
        set_run_font(p.add_run(f"  (출제: {', '.join(q['authors'])})"), 9, italic=True, color=GRAY)
    paras.append(p)

    history = doc.add_paragraph()
    history.paragraph_format.left_indent = Cm(0.4)
    history.paragraph_format.space_after = Pt(1)
    set_run_font(history.add_run(source_history(q)), 9, color=GRAY)
    paras.append(history)

    body_blocks = statement_blocks(part for source in q["body"] for part in display_lines(source))
    prev_blank = False
    for boxed, lines in [*stem_blocks, *body_blocks]:
        if boxed:
            paras.extend(render_statement_box(doc, lines, correct))
            prev_blank = False
            continue
        line = lines[0]
        if not line.strip():
            prev_blank = True
            continue
        im = IMG_RE.search(line)
        if im:
            remainder = IMG_RE.sub("", line).strip()
            if remainder:
                bp = doc.add_paragraph()
                set_run_font(bp.add_run(remainder), 10.5)
                paras.append(bp)
            for match in IMG_RE.finditer(line):
                rel, width = match.group(1), match.group(2)
                paras.append(add_image(doc, resolve_image(rawmd_dir, rel), int(width) if width else None, missing_images))
            prev_blank = False
            continue
        bp = doc.add_paragraph()
        bp.paragraph_format.left_indent = Cm(0.4)
        bp.paragraph_format.space_before = Pt(6 if prev_blank else 0)
        bp.paragraph_format.space_after = Pt(1)
        nm = NOTE_RE.match(line.strip())
        if nm:
            set_run_font(bp.add_run(f"[{nm.group(1)}]"), 9, italic=True, color=GRAY)
        elif correct and choice_marker(line.strip()) in correct:
            set_run_font(bp.add_run(line.strip()), 10.5, bold=True, color=BLUE)
        else:
            set_run_font(bp.add_run(line.strip()), 10.5)
        paras.append(bp)
        prev_blank = False

    # 선지 통합: 대표 복기에 없는 타 연도 선지 덧붙임
    if q.get("extra_choices"):
        np_ = doc.add_paragraph()
        np_.paragraph_format.left_indent = Cm(0.4)
        np_.paragraph_format.space_before = Pt(4)
        np_.paragraph_format.space_after = Pt(1)
        set_run_font(
            np_.add_run("※ 선지 통합 — 아래는 다른 연도 복기에서만 등장한 선지"),
            9, italic=True, color=GRAY,
        )
        paras.append(np_)
        for yr, boxed, lines in year_statement_blocks(
                (yr, part) for yr, source in q["extra_choices"] for part in display_lines(source)):
            if boxed:
                paras.extend(render_statement_box(doc, lines, correct, year=yr))
                continue
            line = lines[0]
            im = IMG_RE.search(line)
            if im:
                cap = doc.add_paragraph()
                cap.paragraph_format.left_indent = Cm(0.8)
                set_run_font(cap.add_run(f"[{ytag(yr)}년 복기 이미지]"), 9, italic=True, color=GRAY)
                paras.append(cap)
                paras.append(
                    add_image(doc, resolve_image(rawmd_dir, im.group(1)),
                              int(im.group(2)) if im.group(2) else None, missing_images)
                )
                continue
            bp = doc.add_paragraph()
            bp.paragraph_format.left_indent = Cm(0.8)
            bp.paragraph_format.space_after = Pt(1)
            set_run_font(bp.add_run(f"[{ytag(yr)}] "), 9, color=GRAY)
            if correct and choice_marker(line) in correct:
                set_run_font(bp.add_run(line), 10, bold=True, color=BLUE)
            else:
                set_run_font(bp.add_run(line), 10)
            paras.append(bp)
    if q.get("note"):
        np_ = doc.add_paragraph()
        np_.paragraph_format.left_indent = Cm(0.4)
        np_.paragraph_format.space_after = Pt(1)
        set_run_font(np_.add_run(f"※ 병합 공시: {q['note']}"), 9, italic=True, color=GRAY)
        paras.append(np_)

    # ---- 정답 포함 버전: 해설(객관식) / 키워드 답안(서술형)
    ans = q.get("_answer")
    if ans:
        if ans.get("explanation"):
            ep = doc.add_paragraph()
            ep.paragraph_format.left_indent = Cm(0.4)
            ep.paragraph_format.space_before = Pt(3)
            ep.paragraph_format.space_after = Pt(1)
            set_run_font(ep.add_run("▸ "), 9.5, bold=True, color=BLUE)
            set_run_font(ep.add_run(ans["explanation"]), 9.5, color=BLUE)
            paras.append(ep)
        if ans.get("answer_lines"):
            hp = doc.add_paragraph()
            hp.paragraph_format.left_indent = Cm(0.4)
            hp.paragraph_format.space_before = Pt(3)
            hp.paragraph_format.space_after = Pt(1)
            set_run_font(hp.add_run("▸ 답"), 9.5, bold=True, color=BLUE)
            paras.append(hp)
            for line in ans["answer_lines"]:
                ap_ = doc.add_paragraph()
                ap_.paragraph_format.left_indent = Cm(0.8)
                ap_.paragraph_format.space_after = Pt(1)
                set_run_font(ap_.add_run(f"- {line}"), 10, bold=True, color=BLUE)
                paras.append(ap_)
        if ans.get("uncertain"):
            up = doc.add_paragraph()
            up.paragraph_format.left_indent = Cm(0.4)
            up.paragraph_format.space_after = Pt(1)
            set_run_font(up.add_run("(정답 불확실 — 근거 자료에서 확정하지 못함)"),
                         9, italic=True, color=GRAY)
            paras.append(up)
    keep_block(paras)


def render_extra(doc, extra):
    """[교과서] 등 보충 블록 (강의자료 제공 시에만 사용)."""
    paras = []
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    set_run_font(p.add_run(extra.get("title", "[교과서]")), 10.5, bold=True)
    paras.append(p)
    for line in extra.get("lines", []):
        bp = doc.add_paragraph()
        bp.paragraph_format.left_indent = Cm(0.4)
        bp.paragraph_format.space_after = Pt(1)
        set_run_font(bp.add_run(line), 10)
        paras.append(bp)
    keep_block(paras)


def render_section_heading(doc, text, size=13):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    set_run_font(p.add_run(text), size, bold=True)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:color"), "999999")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


# ---------------------------------------------------------------- main build
def document_lines(doc, elements=None):
    """표 셀을 포함한 문서 순서로 텍스트를 읽는다. 빈 배치 문단은 제외."""
    roots = elements if elements is not None else list(doc.element.body)
    return [p.text for root in roots for element in root.iter(qn("w:p"))
            for p in [Paragraph(element, doc)] if p.text]


def expected_question_lines(q, author_label=False, page_mode=False):
    """원문 IR에서 기대 인쇄 문자열을 만든다. 렌더 결과에서 역으로 추출하지 않는다."""
    lines = []
    box = q.get("_context_box")
    if box:
        lines.append((f"{box['label']}\n" if box.get("label") else "") + box["text"])
    lead, stem_blocks = stem_parts(display_lines(q["stem"]))
    stem = "[" + ",".join(ytag(y) for y in q["years"]) + "] " + "\n".join(lead)
    if page_mode:
        stem += f"  ({fmt_pages(q.get('_pages'))})"
    if author_label and q.get("authors"):
        stem += f"  (출제: {', '.join(q['authors'])})"
    lines.extend([stem, source_history(q)])
    for line in [*(line for _, block in stem_blocks for line in block),
                 *(part for source in q["body"] for part in display_lines(source))]:
        text = IMG_RE.sub("", line).strip()
        if text:
            note = NOTE_RE.match(text)
            lines.append(f"[{note.group(1)}]" if note else text)
    if q.get("extra_choices"):
        lines.append("※ 선지 통합 — 아래는 다른 연도 복기에서만 등장한 선지")
        for year, boxed, block in year_statement_blocks(
                (year, part) for year, source in q["extra_choices"] for part in display_lines(source)):
            for line in block:
                if IMG_RE.search(line):
                    lines.append(f"[{ytag(year)}년 복기 이미지]")
                    text = IMG_RE.sub("", line).strip()
                    if text:
                        lines.append(text)
                else:
                    lines.append(f"[{ytag(year)}] {line.strip() if boxed else line}")
    if q.get("note"):
        lines.append(f"※ 병합 공시: {q['note']}")
    answer = q.get("_answer") or {}
    if answer.get("explanation"):
        lines.append("▸ " + answer["explanation"])
    if answer.get("answer_lines"):
        lines.append("▸ 답")
        lines.extend("- " + line for line in answer["answer_lines"])
    if answer.get("uncertain"):
        lines.append("(정답 불확실 — 근거 자료에서 확정하지 못함)")
    return lines


def print_report(args, report):
    path = getattr(args, "report", None)
    if not path:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return
    write_json(path, report)
    layout = report.get("layout", {})
    photos = report.get("shared_photos", {})
    summary = {key: report.get(key) for key in ("verdict", "output", "source_content_ok",
                                                "n_source_questions", "expected_blocks", "missing_images")}
    summary.update(report=path, word=layout.get("word", {}).get("status"),
                   pages=layout.get("word", {}).get("pages"),
                   photo_review=photos.get("review_file"), contact_sheets=photos.get("contact_sheets"),
                   photo_issue_count=len(photos.get("issues") or []),
                   photo_issues=(photos.get("issues") or [])[:6])
    print(json.dumps(summary, ensure_ascii=False))


def build(args):
    rawmd_path = Path(args.rawmd)
    rawmd_dir = rawmd_path.resolve().parent
    index = getattr(args, "index", None)
    selection = getattr(args, "selection", None)
    csv_arg = getattr(args, "csv", None)
    csv_input = Path(csv_arg) if csv_arg else find_csv(rawmd_path.resolve())
    if bool(index) != bool(selection):
        raise ValueError("--index and --selection must be supplied together.")
    if selection:
        if args.topic or args.include or args.exclude:
            raise ValueError("--selection already supplies the complete list of questions.")
        bank = source_index(index, rawmd_path, csv_arg)
        csv_input = index_csv_path(bank, rawmd_path, csv_arg)
        table, chapters = bank["table"], bank["chapters"]
        selected = select_chapters(bank, selection)
    else:
        if not args.topic:
            raise ValueError("Supply --topic or a reviewed --selection.")
        raw_text = rawmd_path.read_text(encoding="utf-8-sig")
        table = parse_matching_table(raw_text)
        parser = SimpleNamespace(**{name: globals()[name] for name in (
            "CHAPTER_RE", "Q_RE", "IMG_RE", "KY_OFFSET", "parse_chapter_title", "parse_chapters")})
        chapters = parse_bank(raw_text, parser)
        selected = []
        for want in args.topic:
            found = [c for c in chapters if c["questions"] and c["topic"] == want]
            if not found:
                found = [c for c in chapters if c["questions"] and want in c["topic"]]
            if not found:
                raise ValueError(f"Chapter not found: {want}. Available: "
                                 + " / ".join(c["topic"] for c in chapters if c["questions"]))
            if len({ch["topic"] for ch in found}) > 1:
                raise ValueError(f"Ambiguous chapter: {want}. Use an exact topic from "
                                 + " / ".join(c["topic"] for c in found))
            selected.extend(ch for ch in found if ch not in selected)

    # ---- 오배치 문항 편입/제외 (임의 분류 교차 탐색 결과 반영)
    #   임의 분류 문항은 주제 표기 없이 자동 배정된 것이라, 실제로는 다른 교수 파트에
    #   속하는 문항이 섞여 있다. 그대로 두면 그 교수 파트에서 통째로 누락된다.
    def locate(spec):
        """{"year":25,"no":"1"} 또는 {"year":25,"stem":"이명법"} → [(chapter, question)]"""
        hits = []
        for ch in chapters:
            for q in ch["questions"]:
                if q["year"] != yparse(spec["year"]):
                    continue
                if "no" in spec and q["no"] == str(spec["no"]):
                    hits.append((ch, q))
                elif "stem" in spec and spec["stem"] in q["stem"]:
                    hits.append((ch, q))
        return hits

    pulled, dropped, unresolved = [], [], []
    if args.include:
        for spec in json.loads(Path(args.include).read_text(encoding="utf-8")):
            hits = locate(spec)
            if not hits:
                unresolved.append({"op": "include", "spec": spec})
                continue
            for src_ch, q in hits:
                if src_ch in selected:
                    continue  # 이미 대상 챕터에 있음
                src_ch["questions"].remove(q)
                q["_pulled_from"] = src_ch["topic"]
                selected[0]["questions"].append(q)
                pulled.append({"tag": f"{q['year']:02d}-{q['no'] or '?'}",
                               "stem": q["stem"][:60], "from": src_ch["topic"]})
    if args.exclude:
        for spec in json.loads(Path(args.exclude).read_text(encoding="utf-8")):
            hits = [(ch, q) for ch, q in locate(spec) if ch in selected]
            if not hits:
                unresolved.append({"op": "exclude", "spec": spec})
                continue
            for ch, q in hits:
                ch["questions"].remove(q)
                dropped.append({"tag": f"{q['year']:02d}-{q['no'] or '?'}",
                                "stem": q["stem"][:60]})

    inputs = [
        rawmd_path, csv_input, index, selection, args.include, args.exclude, args.merge_map, args.year_prof,
        args.page_map, args.context_box, args.extras, args.answers, args.comments,
        getattr(args, "photo_review", None),
    ]
    inputs.extend(resolve_image(rawmd_dir, match.group(1))
                  for ch in selected for q in ch["questions"] for line in [q["stem"], *q["body"]]
                  for match in IMG_RE.finditer(line))
    extra_outputs = [getattr(args, "report", None), getattr(args, "preview_pdf", None)]
    check_output_paths(extra_outputs, inputs)
    preview_path = extra_outputs[1]
    if preview_path and Path(preview_path).suffix.lower() != ".pdf":
        raise ValueError("--preview-pdf requires a .pdf file.")
    source_ids = [q["_source_id"] for ch in selected for q in ch["questions"]]
    if not source_ids or len(source_ids) != len(set(source_ids)):
        raise ValueError("The selected chapters must contain at least one source question and each ID once.")
    photo_session, photo_status = prepare_review(args, selected, rawmd_dir, resolve_image, table)
    if photo_status and photo_status["verdict"] == "CHECK":
        result = {"verdict": "CHECK", "source_content_ok": False, "output": None,
                  "missing_images": [], "chapters": [], "shared_photos": photo_status}
        result["n_source_questions"] = len(source_ids)
        print_report(args, result)
        return result
    doc = Document()
    image_dir = rawmd_dir / "동봉 이미지"
    source_dirs = {image_dir} if image_dir.is_dir() else set()
    source_dirs.update(resolve_image(rawmd_dir, match.group(1)).parent
                       for chapter in selected for question in chapter["questions"]
                       for line in question["body"] for match in IMG_RE.finditer(line))
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin = sec.bottom_margin = Cm(2.2)
    sec.left_margin = sec.right_margin = Cm(2.2)
    layout = YamaLayout(doc, source_dirs=sorted(source_dirs))
    add_page_number(sec)

    # ---- 제목 블록
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(title.add_run(f"{args.subject} {args.exam} 기출 야마"), 16, bold=True)

    report = {
        "output": None, "chapters": [], "missing_images": [],
        "reassignment": {
            "pulled_in": pulled, "excluded": dropped, "unresolved_specs": unresolved,
            "inferred_candidates": [] if selection else scan_inferred_candidates(selected, chapters),
        },
    }
    subtitle_topics = []
    for ch in selected:
        info = table.get(ch["topic"], {})
        past = ch["past"] or (info.get("past") if info.get("past") not in ("미표기", "", "—") else None)
        current = ch["current"] or info.get("current") or "미상"
        ch["_past"], ch["_current"] = past, current
        subtitle_topics.append(f"{ch['topic']} ({current} 교수님)")

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(sub.add_run(" · ".join(subtitle_topics)), 12, bold=True)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(
        meta.add_run(f"원본: {rawmd_path.stem} · 생성일: {date.today().isoformat()}"),
        9, color=GRAY,
    )

    # ---- 병합·분리 선계산 (정리자 comment 통계에 필요)
    extras = []
    if args.extras:
        extras = json.loads(Path(args.extras).read_text(encoding="utf-8"))
    merge_map = []
    if args.merge_map:
        merge_map = json.loads(Path(args.merge_map).read_text(encoding="utf-8"))
    year_prof = {}
    if args.year_prof:
        year_prof = json.loads(Path(args.year_prof).read_text(encoding="utf-8"))
    answers = []
    if args.answers:
        answers = json.loads(Path(args.answers).read_text(encoding="utf-8"))
    page_map = []
    if args.page_map:
        page_map = json.loads(Path(args.page_map).read_text(encoding="utf-8"))

    context_boxes = []
    if args.context_box:
        context_boxes = json.loads(Path(args.context_box).read_text(encoding="utf-8"))

    # 정렬 기준은 페이지 라벨 표기와 독립이다. --page-map을 라벨용으로만 쓰고
    # 배열은 연도순으로 갈 수 있다.
    order = args.sort or ("page" if page_map else None)

    def prof_for(ch, year):
        """해당 연도에 이 주제를 출제한 교수. 매핑 우선순위: 주제 키 > default > 챕터 표기."""
        yk = ytag(year)
        for key, m in year_prof.items():
            if key != "default" and key in ch["topic"] and yk in m:
                return m[yk]
        if yk in year_prof.get("default", {}):
            return year_prof["default"][yk]
        return None

    for ch in selected:
        merged = merge_questions(ch["questions"], merge_map, rawmd_dir=rawmd_dir)
        ch["_merged"] = merged

        # 페이지 매핑: members(연도-번호) 우선, 없으면 문두 부분 문자열
        for entry in page_map:
            entry.setdefault("_used", 0)
        for q in ch["_merged"]:
            for entry in page_map:
                if q.get("_pages"):
                    break
                mem = {f"{ytag(yparse(m['year']))}-{m.get('no', '?')}" for m in entry.get("members", [])}
                hit = (mem and mem & set(q["nos"])) or (
                    entry.get("match") and entry["match"] in q["stem"])
                if hit:
                    q["_pages"] = entry.get("pages") or []
                    q["_page_note"] = entry.get("note")
                    entry["_used"] += 1

        # 공통지문 박스: 매칭 규칙은 page_map과 동일(members 우선, 없으면 문두 부분 문자열)
        for entry in context_boxes:
            entry.setdefault("_used", 0)
        for q in ch["_merged"]:
            for entry in context_boxes:
                if q.get("_context_box"):
                    break
                mem = {f"{ytag(yparse(m['year']))}-{m.get('no', '?')}" for m in entry.get("members", [])}
                if (mem and mem & set(q["nos"])) or (
                        entry.get("match") and entry["match"] in q["stem"]):
                    q["_context_box"] = entry
                    entry["_used"] += 1

        main, other = [], []
        for q in merged:
            authors = sorted({p for p in (
                raw.get("_explicit_prof") or prof_for(ch, raw["year"])
                for raw in q["_source_questions"]) if p})
            q["authors"] = authors
            # 연도 중 하나라도 현재 담당 출제면 본문, 출제자 미상도 본문(확인 불가는 유지)
            if not authors or any(professor_in_current(author, ch["_current"]) for author in authors):
                main.append(q)
            else:
                other.append(q)
        # 페이지 순서 모드: 교안 페이지 오름차순 → 같은 페이지면 최신 연도가 위로.
        # 페이지 미상 문항은 순서를 왜곡하지 않도록 맨 뒤로 보낸다.
        if order == "page":
            def page_key(q):
                pages = q.get("_pages") or []
                return (0, min(pages), -max(q["years"])) if pages else (1, 10 ** 6, -max(q["years"]))
            main.sort(key=page_key)
            other.sort(key=page_key)
        # 연도 순서 모드: 최신 연도 → 그 연도의 문항 번호 오름차순.
        # 공통지문 세트는 한 시험지에서 번호가 연속하므로 이 정렬에서 자연히 붙어 있다.
        elif order == "year":
            main.sort(key=year_no_key)
            other.sort(key=year_no_key)
        ch["_main"], ch["_other"] = main, other

        # 정답 매칭: members(연도-번호) 우선, 없으면 문두 부분 문자열
        for entry in answers:
            entry.setdefault("_used", 0)
        for q in ch["_merged"]:
            for entry in answers:
                if q.get("_answer"):
                    break
                mem = {f"{ytag(yparse(m['year']))}-{m.get('no', '?')}" for m in entry.get("members", [])}
                if mem and mem & set(q["nos"]):
                    q["_answer"] = entry
                    entry["_used"] += 1
                elif entry.get("match") and entry["match"] in q["stem"]:
                    q["_answer"] = entry
                    entry["_used"] += 1

    # ---- [정리자 comment]
    # 본체는 --comments로 전달되는 분석(개념 흐름, 빈출 경향, 교수 성향, 학습 태도).
    # 스크립트는 문서 구조 이해에 필요한 최소한의 안내만 뒤에 덧붙인다.
    render_section_heading(doc, "[정리자 comment]", 11)
    comment_lines = []
    if args.comments:
        comment_lines += [
            l.strip() for l in Path(args.comments).read_text(encoding="utf-8").splitlines()
            if l.strip()
        ]
    for ch in selected:
        years = sorted({q["year"] for q in ch["questions"]}, reverse=True)
        per_year = " · ".join(
            f"{ytag(y)}년 {sum(1 for q in ch['questions'] if q['year'] == y)}"
            for y in years
        )
        comment_lines.append(f"『{ch['topic']}』 연도별 출제량: {per_year}")
        if ch["_other"]:
            profs = sorted({p for q in ch["_other"] for p in q["authors"]})
            comment_lines.append(
                f"타 교수({', '.join(profs)}) 출제 문항 {len(ch['_other'])}개는 문서 하단에 "
                f"분리했습니다 (현재 담당: {ch['_current']})."
            )
    if order == "page":
        comment_lines.append(
            "문항은 교안 페이지 오름차순으로 배열했습니다 (같은 페이지면 최신 연도가 위). "
            "문두 옆 (p.N)은 근거 페이지이며, 근거를 특정하지 못한 문항은 p.미상으로 표기하고 맨 뒤에 두었습니다."
        )
    elif order == "year":
        comment_lines.append(
            "문항은 최신 연도부터, 연도 안에서는 시험지 문항 번호 오름차순으로 배열했습니다. "
            "공통지문을 쓰는 세트 문항이 원본 시험지처럼 붙어 있습니다."
            + (" 문두 옆 (p.N)은 교안 근거 페이지이며 정렬에는 쓰지 않았습니다." if page_map else "")
        )
    comment_lines.append(
        "표기 안내: [25,23]=복수 연도 출제(병합) · ※ 선지 통합=연도별 복기 선지를 합침 · "
        "(미복기)=복기 누락" + ("" if answers else " · 답안은 수록하지 않음") + "."
    )
    for line in comment_lines:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.4)
        p.paragraph_format.space_after = Pt(2)
        set_run_font(p.add_run(f"- {line}"), 10)

    # ---- 본문
    source_checks = []
    expected_images = 0
    def emit_one(q, author_label=False, block_start=None):
        nonlocal expected_images
        start = len(doc.element.body)
        expected = expected_question_lines(q, author_label, bool(page_map))
        expected_images += sum(len(IMG_RE.findall(line)) for line in q["body"])
        expected_images += sum(len(IMG_RE.findall(line)) for _, line in q["extra_choices"])
        render_question(doc, q, rawmd_dir, report["missing_images"],
                        author_label=author_label, page_mode=bool(page_map))
        for ex in extras:
            if ex.get("match") and ex["match"] in q["stem"]:
                render_extra(doc, ex)
                expected.append(ex.get("title", "[교과서]"))
                expected.extend(line for line in ex.get("lines", []) if line)
        # python-docx는 마지막 sectPr 앞에 새 본문을 삽입한다.
        elements = list(doc.element.body)[start - 1:-1]
        lead, stem_blocks = stem_parts(display_lines(q["stem"]))
        displayed_body = statement_blocks(
            part for source in q["body"] for part in display_lines(source))
        display_checks = [
            (q["stem"], lead + [line for _, block in stem_blocks for line in block]),
            ("\n".join(q["body"]), [line for _, block in displayed_body for line in block]),
            ("\n".join(source for _, source in q["extra_choices"]), [
                line for _, _, block in year_statement_blocks(
                    (year, part) for year, source in q["extra_choices"] for part in display_lines(source))
                for line in block]),
        ]
        source_checks.append({"members": q["nos"],
                              "text_ok": expected == document_lines(doc, elements),
                              "display_content_preserved": all(
                                  display_text_preserved(source, displayed)
                                  for source, displayed in display_checks)})
        layout.add_block(block_start if block_start is not None else start,
                         f"{len(source_checks)}:" + " · ".join(q["nos"]))

    photo_builder = SimpleNamespace(**{name: globals()[name] for name in
        ("set_run_font", "add_image", "resolve_image", "expected_question_lines", "source_history", "ytag")})
    if photo_session:
        photo_session.page_mode = bool(page_map)

    def emit_questions(merged_list, author_label=False, section_label=None):
        if photo_session:
            photo_session.section_label = section_label
            photo_session.emit(doc, layout, merged_list, emit_one, photo_builder, author_label)
        else:
            for question in merged_list:
                emit_one(question, author_label)

    for ch in selected:
        info = table.get(ch["topic"], {})
        expected = info.get("n")
        if ch.get("_selected_by_id"):
            baseline_expected = sum(len(raw["questions"]) for raw in chapters if raw["topic"] == ch["topic"])
        n_pulled = sum(1 for q in ch["questions"] if q.get("_pulled_from"))
        n_dropped = sum(1 for d in dropped if True) if ch is selected[0] else 0
        # 편입·제외를 되돌린 원래 문항 수로 매칭 표와 대조해야 검증이 의미가 있다
        baseline = len(ch["questions"]) - n_pulled + n_dropped
        report["chapters"].append(
            {
                "topic": ch["topic"],
                "past_prof": ch["_past"],
                "current_prof": ch["_current"],
                "n_raw": len(ch["questions"]),
                "n_pulled_in": n_pulled,
                "n_merged": len(ch["_merged"]),
                "n_main": len(ch["_main"]),
                "n_other_prof": len(ch["_other"]),
                "n_merged_groups": sum(1 for q in ch["_merged"] if q["dup_count"] > 1),
                "merged_groups": [
                    {"members": q["nos"], "stem": q["stem"][:60],
                     "choice_union": bool(q["extra_choices"]),
                     "image_dup": q["image_dup"], "image_duplicates": q["image_duplicates"],
                     "note": q.get("note")}
                    for q in ch["_merged"] if q["dup_count"] > 1
                ],
                "merge_candidates": find_merge_candidates(ch["_merged"]),
                **({
                    "pages": {
                        "n_mapped": sum(1 for q in ch["_merged"] if q.get("_pages")),
                        "n_unmapped": sum(1 for q in ch["_merged"] if not q.get("_pages")),
                        "unmapped_stems": [q["stem"][:50] for q in ch["_merged"]
                                           if not q.get("_pages")],
                    }
                } if page_map else {}),
                **({
                    "context_boxes": {
                        "n_applied": sum(1 for q in ch["_merged"] if q.get("_context_box")),
                        "unused_entries": [e.get("label") or e.get("match", "")[:40]
                                           for e in context_boxes if not e.get("_used")],
                    }
                } if context_boxes else {}),
                "expected_from_table": expected,
                "count_ok": (expected is None or str(
                    baseline_expected if ch.get("_selected_by_id") else baseline) == str(expected)),
            }
        )

    for ch in selected:
        if not ch["_main"]:
            continue
        if len(selected) > 1:
            render_section_heading(doc, f"{ch['topic']} — {ch['_current']}")
        emit_questions(ch["_main"], section_label=ch["topic"])

    # ---- 타 교수 야마
    if any(ch["_other"] for ch in selected):
        doc.add_page_break()
        render_section_heading(doc, "타 교수 야마", 14)
        for ch in selected:
            if not ch["_other"]:
                continue
            profs = sorted({p for q in ch["_other"] for p in q["authors"]})
            render_section_heading(
                doc, f"{ch['topic']} — 출제: {', '.join(profs)} (현재 담당: {ch['_current']})", 12
            )
            note = doc.add_paragraph()
            set_run_font(
                note.add_run(
                    "이번 연도 담당 교수와 기출 출제 교수가 다릅니다. "
                    "출제 스타일이 달라질 수 있으므로 참고용으로 활용하십시오."
                ),
                9, italic=True, color=GRAY,
            )
            emit_questions(ch["_other"], author_label=(len(profs) > 1), section_label="타 교수 야마 · " + ch["topic"])

    # ---- 저장
    if args.outname:
        fname = args.outname
    else:
        topic0 = re.sub(r'[\\/:*?"<>|·]+', "_", selected[0]["topic"])[:40]
        prof0 = selected[0]["_current"]
        suffix = ({"page": " 페이지순", "year": " 연도순"}.get(order, "")
                  + (" 정답포함" if answers else ""))
        fname = f"{args.subject}_{topic0}_{prof0} 야마{suffix}_ver{args.ver}.docx"
    if not fname.lower().endswith(".docx"):
        fname += ".docx"
    out = Path(args.out) / fname
    check_output_paths([out, out.with_suffix(".layout-qc.json"), *extra_outputs], inputs)
    out.parent.mkdir(parents=True, exist_ok=True)
    expected_blocks = sum(len(ch["_merged"]) for ch in selected)
    reconstructed_ids = [
        source["_source_id"] for ch in selected for q in ch["_merged"] for source in q["_source_questions"]
    ]
    report["n_source_questions"] = len(source_ids)
    report["source_ids"] = source_ids
    report["source_ids_preserved"] = sorted(source_ids) == sorted(reconstructed_ids)
    if photo_session:
        expected_images += photo_session.shared_image_count
    rendered_images = len(doc.inline_shapes)
    before_lines = document_lines(doc)
    report["layout"] = layout.save(out)
    if photo_session:
        report["layout"] = photo_session.add_continuations(doc, layout, out, report["layout"], photo_builder)
        before_lines = document_lines(doc)
    saved = Document(str(out))
    report["source_content_ok"] = (
        report["source_ids_preserved"]
        and all(check["text_ok"] and check["display_content_preserved"] for check in source_checks)
        and len(source_checks) == expected_blocks
        and document_lines(saved) == before_lines
        and sum(line.startswith("※ 출제 이력: ") for line in document_lines(saved)) == expected_blocks
    )
    report["source_checks"] = source_checks
    report["shared_photos"] = photo_session.finalize(out, report["layout"]) if photo_session else photo_status
    report["expected_blocks"] = expected_blocks
    report["expected_images_before_layout"] = expected_images
    report["images_rendered_before_layout"] = rendered_images
    report["verdict"] = "OK" if (
        report["layout"].get("verdict") == "OK" and report["source_content_ok"]
        and expected_images == rendered_images and not report["missing_images"]
        and all(ch["count_ok"] for ch in report["chapters"])
        and report["shared_photos"]["verdict"] in ("OK", "NOT_APPLICABLE")
        and not unresolved
    ) else "CHECK"
    report["output"] = str(out)
    if getattr(args, "preview_pdf", None):
        from yama_layout import measure_word

        preview = Path(args.preview_pdf).resolve()
        preview.parent.mkdir(parents=True, exist_ok=True)
        result = measure_word(out, [], [], [], export_pdf=preview)
        report["preview"] = {"path": str(preview) if result["status"] == "OK" else None,
                             "word_status": result["status"], "error": result.get("error")}
    if answers:
        answered = [q for ch in selected for q in ch["_merged"] if q.get("_answer")]
        report["answers"] = {
            "n_answered": len(answered),
            "n_unanswered": sum(len(ch["_merged"]) for ch in selected) - len(answered),
            "unanswered_stems": [
                q["stem"][:50] for ch in selected for q in ch["_merged"]
                if not q.get("_answer")
            ],
            "unused_entries": [
                e.get("match") or str(e.get("members")) for e in answers if not e["_used"]
            ],
            "entries_matched_multiple": [
                e.get("match") or str(e.get("members")) for e in answers if e["_used"] > 1
            ],
        }
    print_report(args, report)
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rawmd", required=True)
    ap.add_argument("--photo-review", "--shared-photos", dest="photo_review",
                    help="전체 사진 inventory 및 모든 고유 대표 쌍의 검토를 완료한 manifest JSON")
    ap.add_argument("--topic", action="append", help="챕터 주제(직접 선택 모드). 여러 번 지정 가능")
    ap.add_argument("--index", help="yama_bank.py index로 만든 로컬 파일")
    ap.add_argument("--csv", help="별도 폴더에 있는 색인 원본 CSV")
    ap.add_argument("--selection", help="확정한 원문 문항 ID 목록 JSON")
    ap.add_argument("--report", help="전체 QC 보고서 JSON. 지정하면 stdout에는 요약만 출력.")
    ap.add_argument("--photo-work", help="사진 inventory와 contact sheet 작업 폴더")
    ap.add_argument("--preview-pdf", help="페이지 시각 검수용 Word 렌더 PDF 경로")
    ap.add_argument("--out", default=".")
    ap.add_argument("--subject", default="과목")
    ap.add_argument("--exam", default="1차")
    ap.add_argument("--comments", help="정리자 comment 추가 줄 파일(줄당 1개)")
    ap.add_argument("--extras", help="[교과서]형 보충 블록 JSON")
    ap.add_argument("--merge-map", help="의미 판단 병합 그룹 JSON: "
                    '[{"members":[{"year":23,"no":"14"},{"year":21,"no":"5"}],"note":"..."}]')
    ap.add_argument("--year-prof", help="연도별 출제 교수 매핑 JSON: "
                    '{"default":{"25":"교수A"},"<주제 부분문자열>":{"23":"교수B"}}')
    ap.add_argument("--include", help="다른 챕터에서 끌어올 문항 JSON: "
                    '[{"year":25,"no":"1"},{"year":23,"stem":"이명법"}]')
    ap.add_argument("--exclude", help="이 챕터에서 뺄 문항 JSON(동일 형식)")
    ap.add_argument("--page-map", help="교안 페이지 매핑 JSON(지정 시 페이지 오름차순 정렬): "
                    '[{"match":"문두 부분문자열","pages":[3,4],"note":"근거 메모"}]')
    ap.add_argument("--sort", choices=["page", "year"],
                    help="배열 기준. page=교안 페이지 오름차순 / year=연도 내림차순→번호 오름차순. "
                         "생략하면 --page-map이 있을 때 page. --page-map을 페이지 라벨 표기용으로만 "
                         "쓰려면 --sort year를 함께 준다")
    ap.add_argument("--context-box", help="세트형 문항의 공통지문 박스 JSON: "
                    "[{\"members\":[{\"year\":16,\"no\":\"16\"}],\"label\":\"...\",\"text\":\"...\"}]")
    ap.add_argument("--answers", help="정답 포함 버전 JSON: "
                    '[{"match":"문두 부분문자열","correct":["가","다"],"explanation":"1-2줄"},'
                    '{"match":"...","answer_lines":["키워드 답안 줄", "..."]}]')
    ap.add_argument("--ver", default="1")
    ap.add_argument("--outname")
    args = ap.parse_args()
    try:
        report = build(args)
    except (ValueError, OSError, KeyError, TypeError) as exc:
        ap.exit(2, f"{exc}\n")
    if report["verdict"] != "OK":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
