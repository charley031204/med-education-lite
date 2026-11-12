"""Build an explained Obsidian question bank from a question-only Word index."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import sys

from question_index import HISTORY_YEAR, TYPES, materialize, question_text, read_json, verify_stamp, year_label

STATUSES = ("공식 답안", "교수 직접 답변", "자료 기반 풀이", "정답 불확실")
ROLES = {"lecture": "강의록", "script": "Script", "official": "공식 답안", "textbook": "표준교과서"}
CHILD_MARKER = re.compile(r"^\s*(?:>\s*)*\[![\w-]+\]")
COMMENTS = ("**수업 주제**", "**시험 포인트**", "**자료 안내**")


def lines(value, field: str, *, allow_empty=False) -> list[str]:
    if (not isinstance(value, list) or (not value and not allow_empty)
            or any(not isinstance(x, str) or not x.strip() or "\n" in x or "\r" in x
                   or CHILD_MARKER.match(x) for x in value)):
        raise ValueError(f"{field}에는 독립적인 Markdown 본문 줄을 넣으세요.")
    return value


def single(value, field: str) -> str:
    return lines([value], field)[0]


def selected_questions(index: dict, sort: str, min_year: int | None) -> list[dict]:
    def year_score(year):
        if year.startswith("건양대"):
            return int(year[3:]) - 1000
        return int(year) if year.isdigit() else 0

    def recent(question):
        return max((year_score(y) for y in question["years"]), default=0)

    selected = [q for q in index["questions"] if min_year is None or recent(q) >= min_year]
    if sort == "year-desc":
        selected.sort(key=lambda q: (q["section_order"], -recent(q)))
    if not selected:
        raise ValueError("선택 범위에 문항이 없습니다.")
    return selected


def validate_answers(document: dict, questions: list[dict]) -> dict:
    comments = lines(document.get("comment_lines"), "comment_lines")
    if not 3 <= len(comments) <= 8 or not all(line.startswith(COMMENTS) for line in comments):
        raise ValueError("comment_lines에는 수업 주제·시험 포인트·자료 안내 3–8줄이 필요합니다.")
    expected = {q["qid"]: q for q in questions}
    by_id = {}
    for answer in document["answers"]:
        qid = answer["qid"]
        if qid in by_id or qid not in expected:
            raise ValueError(f"문항 범위 또는 중복 ID를 확인하세요. {qid}")
        q = expected[qid]
        resolved_type = answer.get("type", q["type"])
        if resolved_type not in TYPES or (resolved_type != q["type"] and not q["type_uncertain"]):
            raise ValueError(f"문항 유형을 확인하세요. {qid}")
        if answer["status"] not in STATUSES:
            raise ValueError(f"답안 근거 유형을 확인하세요. {qid}")
        lines(answer["answer_lines"], f"{qid} 답안", allow_empty=answer["status"] == "정답 불확실")
        lines(answer["concept_lines"], f"{qid} 핵심 개념")
        options = answer.get("options", [])
        if [option["marker"] for option in options] != q["choice_markers"]:
            raise ValueError(f"Word의 선지 순서와 해설의 선지 순서가 다릅니다. {qid}")
        for option in options:
            if option["verdict"] not in ("O", "X", "?"):
                raise ValueError(f"선지 판정을 확인하세요. {qid} {option['marker']}")
            single(option["reason"], f"{qid} 선지 근거")
        roles = set()
        for evidence in answer["evidence"]:
            role = evidence["role"]
            if role not in ROLES:
                raise ValueError(f"근거 자료 유형을 확인하세요. {qid}")
            roles.add(role)
            for key in ("source", "location", "text"):
                single(evidence[key], f"{qid} 근거 {key}")
        gap = lines(answer.get("gap_lines", []), f"{qid} 제공 자료의 한계", allow_empty=True)
        conflicts = lines(answer.get("conflict_lines", []), f"{qid} 출처 충돌", allow_empty=True)
        status = answer["status"]
        if (not roles.intersection(("lecture", "script")) or "textbook" in roles) and not gap:
            raise ValueError(f"강의록·Script에서 확인할 수 없었던 범위를 적으세요. {qid}")
        if status == "공식 답안" and "official" not in roles:
            raise ValueError(f"공식 답안의 출처가 필요합니다. {qid}")
        if status == "교수 직접 답변" and "script" not in roles:
            raise ValueError(f"교수 직접 답변의 Script 근거가 필요합니다. {qid}")
        if status == "정답 불확실" and not (gap or conflicts):
            raise ValueError(f"정답을 확정할 수 없는 이유를 적으세요. {qid}")
        by_id[qid] = answer
    missing = set(expected) - set(by_id)
    if missing:
        raise ValueError(f"해설이 없는 문항 ID {', '.join(sorted(missing)[:10])} ({len(missing)}개)")
    return by_id


def quote(text: str, depth: int) -> str:
    marker = " ".join(">" for _ in range(depth))
    return marker + (" " + text if text else "")


def quoted(block: list[str], depth: int) -> list[str]:
    return [quote(line, depth) for line in block]


def render_items(items: list[dict], source: Path, attachments: Path, shared: set,
                 verify_only: bool) -> list[str]:
    output = []
    for item in items:
        kind = item["kind"]
        if kind == "text":
            output.append(item["text"])
        elif kind == "box":
            output.append("")
            output.extend(quoted([f"[!quote] {item['label']}",
                                  *render_items(item["items"], source, attachments, shared, verify_only)], 1))
            output.append("")
        elif kind == "image":
            if item.get("shared_photo") in shared:
                output.append(f"공통 사진 {item['shared_photo']} 참조")
            else:
                image = materialize(item, source, attachments, verify_only=verify_only)
                output.append(f"![[동봉 이미지/{image.name}|400]]")
        else:
            raise ValueError(f"지원하지 않는 문항 요소입니다. {kind}")
    return output


def render_question(q: dict, answer: dict, source: Path, attachments: Path,
                    depth: int, shared: set, verify_only: bool) -> list[str]:
    year_history = {year_label(match[0]) for match in HISTORY_YEAR.finditer(q["history"])}
    header = " · ".join(filter(None, (q["history"], *(
        y + ("년" if y != "연도미상" else "") for y in q["years"] if y not in year_history
    ))))
    question_type = answer.get("type", q["type"])
    block = [f"[!question] [{question_type}] {header}", f"<!-- yama-qid: {q['qid']} -->"]
    if context := q.get("context"):
        context_lines = render_items(context, source, attachments, shared, verify_only)
        block.extend(quoted([f"[!quote] {context_lines[0]}", *context_lines[1:]], 1))
        block.append("")
    block.append(q["stem"] or "<!-- yama-stem-in-body -->")
    block.extend(render_items(q["body"], source, attachments, shared, verify_only))
    while block[-1] == "":
        block.pop()

    answer_block = ["[!answer]- 해설", ""]
    if answer.get("conflict_lines") or answer["status"] == "정답 불확실" or answer.get("gap_lines"):
        answer_block.extend(quoted(["[!warning] 주의", "근거의 한계 또는 상충 내용은 답안에 적었습니다."], 1))
        answer_block.append("")
    tip = ["[!tip] 답안", f"**{answer['status']}**", "", *answer["answer_lines"]]
    if answer.get("options"):
        tip.append("")
        tip.extend(f"- {o['marker']} ({o['verdict']}) {o['reason']}" for o in answer["options"])
    for field, label in (("gap_lines", "제공 자료의 한계"), ("conflict_lines", "출처 충돌")):
        if answer.get(field):
            tip.extend(["", f"**{label}**", *answer[field]])
    answer_block.extend(quoted(tip, 1))
    answer_block.append("")
    answer_block.extend(quoted(["[!info] 핵심 개념", *answer["concept_lines"]], 1))
    if answer["evidence"]:
        answer_block.append("")
        answer_block.extend(quoted(["[!cite]- 근거", *(
            f"- **{ROLES[e['role']]}** · {e['source']} · {e['location']} 「{e['text']}」"
            for e in answer["evidence"]
        )], 1))
    return [*quoted(block, depth), quote("", depth - 1) if depth > 1 else "", *quoted(answer_block, depth)]


def render_bank(index: dict, answers: dict, questions: list[dict], out: Path, *,
                verify_only=False) -> str:
    source = verify_stamp(index["source"])
    attachments = out.parent / "동봉 이미지"
    by_id = validate_answers(answers, questions)
    for alias in index.get("original_aliases", []):
        materialize({"member": alias["part"]}, source, attachments, verify_only=verify_only)
    output = ["> [!quote] 공지", "> 본 자료의 저작권은 **KAISTUDY**에 있습니다.", ">",
              "> > [!abstract] 정리 Comment", *quoted(answers["comment_lines"], 2), ""]
    selected = {q["qid"] for q in questions}
    groups = {qid: group for group in index["groups"] for qid in group["question_ids"]}
    done = set()
    prior_section = None
    for question in questions:
        qid = question["qid"]
        if qid in done:
            continue
        section = question["section_order"]
        if prior_section != section:
            output.extend([f"## {question['section']}", ""])
            prior_section = section
        output.extend(["---", ""])
        group = groups.get(qid)
        shared_images = [image for image in group["images"]
                         if len(selected.intersection(image["question_ids"])) > 1] if group else []
        if not shared_images:
            output.extend(render_question(question, by_id[qid], source, attachments, 1, set(), verify_only))
            done.add(qid)
        else:
            members = [q for q in questions if q["qid"] in group["question_ids"]]
            if any(member["section_order"] != section for member in members):
                raise ValueError(f"공통사진 문항이 Word 섹션을 가로지릅니다. {group['id']}")
            shared = {image["photo_id"] for image in shared_images}
            output.append("> [!question] [공통사진 문항]")
            for image in shared_images:
                original = materialize(image, source, attachments, verify_only=verify_only)
                output.extend(quoted([f"공통 사진 {image['photo_id']}", f"![[동봉 이미지/{original.name}|400]]",
                                      *image["display_notes"]], 1))
            for member in members:
                member_id = member["qid"]
                output.append(">")
                output.extend(render_question(member, by_id[member_id], source, attachments, 2, shared, verify_only))
                done.add(member_id)
        output.append("")
    return "\n".join(output).rstrip() + "\n"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("list", "emit", "qc"))
    parser.add_argument("--index", required=True, type=Path)
    parser.add_argument("--answers", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--sort", choices=("year-desc", "word"), default="year-desc")
    parser.add_argument("--min-year", type=int, choices=range(0, 100), metavar="YY")
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--limit", type=int, choices=range(1, 11), default=5)
    args = parser.parse_args(argv)
    try:
        index = read_json(args.index)
        source = verify_stamp(index["source"])
        if source.suffix.lower() != ".docx":
            raise ValueError("Obsidian 해설집은 문항 전용 Word 인덱스가 필요합니다.")
        questions = selected_questions(index, args.sort, args.min_year)
        if args.command == "list":
            if args.offset < 0:
                raise ValueError("--offset은 0 이상이어야 합니다.")
            result = {"count": len(questions), "offset": args.offset,
                      "questions": [{**{k: q[k] for k in ("qid", "section", "history", "location", "type")},
                                     "stem": question_text(q)[:140]}
                                    for q in questions[args.offset:args.offset + args.limit]]}
        else:
            if not args.answers or not args.out:
                raise ValueError("--answers와 --out이 필요합니다.")
            if args.out.suffix.lower() != ".md" or args.out.resolve() in (
                    source, args.index.resolve(), args.answers.resolve()):
                raise ValueError("별도의 Markdown 출력 경로가 필요합니다.")
            if args.command == "emit" and args.out.exists():
                raise ValueError(f"이미 출력 파일이 있습니다. {args.out}")
            text = render_bank(index, read_json(args.answers), questions, args.out,
                               verify_only=args.command == "qc")
            if args.command == "qc":
                if args.out.read_text(encoding="utf-8") != text:
                    raise ValueError("저장된 Markdown과 원본에서 재현한 문항·해설이 일치하지 않습니다.")
            else:
                args.out.parent.mkdir(parents=True, exist_ok=True)
                with args.out.open("x", encoding="utf-8", newline="\n") as stream:
                    stream.write(text)
            result = {"status": "ok", "count": len(questions), "output": str(args.out.resolve())}
    except (KeyError, OSError, TypeError, ValueError) as exc:
        print(f"CHECK — {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
