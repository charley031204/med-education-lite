"""Exercise bounded Yama lookup, local evidence and Obsidian conversion."""

import json
from pathlib import Path
import subprocess
import sys
from tempfile import mkdtemp
import unittest

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
WORD = ROOT / "skills/yama-word-obsidian/scripts"
QA = ROOT / "skills/yama-question-answer/scripts"
QUESTION = WORD / "question_index.py"
BANK = WORD / "obsidian_bank.py"
EVIDENCE = QA / "local_evidence.py"


def word_fixture(root):
    doc = Document()
    heading = doc.add_paragraph()
    run = heading.add_run("합성 과목 — 합성 교수")
    run.bold, run.font.size = True, Pt(13)
    case = doc.add_table(rows=1, cols=1)
    case.cell(0, 0).text = "공통지문 — 1차\n분류 대상의 위치를 표시한 그림이다."
    for round_number, years, color, stem in [
        (1, "[25,23]", "navy", "상기 그림의 알파 장치를 분류하면?"),
        (2, "[25]", "orange", ""),
    ]:
        paragraph = doc.add_paragraph()
        paragraph.add_run(f"{years} {stem}").bold = True
        history = doc.add_paragraph(f"※ 출제 이력: 25년 {round_number}차 88번"
                                    + (" · 23년 3차 70번" if round_number == 1 else ""))
        history.paragraph_format.left_indent = Pt(10)
        doc.add_paragraph(f"[Figure {round_number}] 합성 알파 장치의 위치")
        image = root / f"photo-{round_number}.png"
        Image.new("RGB", (30, 18), color=color).save(image)
        doc.add_picture(str(image))
        if not stem:
            doc.add_paragraph("뒤 회차에서 확인하는 별도 발문이다.")
        box = doc.add_table(rows=1, cols=1)
        caption = OxmlElement("w:tblCaption")
        caption.set(qn("w:val"), "YamaStatementBox")
        box._tbl.tblPr.append(caption)
        box.cell(0, 0).text = "가) 합성 조건\n나) 비교 조건"
        doc.add_paragraph("① 가")
        doc.add_paragraph("② 나")
    source = root / "exam.docx"
    doc.save(source)
    return source


class QaLiteTests(unittest.TestCase):
    def setUp(self):
        self.folder = Path(mkdtemp(prefix="qa-lite-"))

    def cli(self, script, *args, code=0):
        result = subprocess.run([sys.executable, "-X", "utf8", str(script), *map(str, args)],
                                capture_output=True, encoding="utf-8", check=False)
        self.assertEqual(result.returncode, code, result.stdout + result.stderr)
        return json.loads(result.stdout) if result.stdout else result.stderr

    def test_lookup_keeps_duplicate_candidates_small_and_extracts_only_selected_figure(self):
        doc = word_fixture(self.folder)
        index = self.folder / "questions.json"
        summary = self.cli(QUESTION, "build", "--source", doc, "--out", index)
        self.assertEqual(summary["count"], 2)
        self.assertNotIn("합성 조건", json.dumps(summary, ensure_ascii=False))
        ambiguous = self.cli(QUESTION, "lookup", "--index", index, "--year", 25, "--number", 88, code=1)
        self.assertEqual(ambiguous["status"], "ambiguous")
        self.assertEqual([c["qid"] for c in ambiguous["candidates"]], ["q001", "q002"])
        self.assertTrue(all("body" not in c for c in ambiguous["candidates"]))
        figures = self.folder / "selected"
        packet = self.cli(QUESTION, "lookup", "--index", index, "--year", 25, "--round", "1차",
                          "--number", 88, "--figures", figures)
        question = packet["question"]
        self.assertEqual(question["choice_markers"], ["①", "②"])
        self.assertEqual(question["context"][0]["text"], "공통지문 — 1차")
        self.assertEqual(question["history"], "25년 1차 88번 · 23년 3차 70번")
        self.assertEqual(len(list(figures.iterdir())), 1)
        figure = next(i for i in question["body"] if i["kind"] == "image")
        self.assertEqual(Path(figure["local_file"]).read_bytes(), (self.folder / "photo-1.png").read_bytes())
        other = self.cli(QUESTION, "lookup", "--index", index, "--text", "별도 발문")["question"]
        self.assertEqual(other["qid"], "q002")
        self.assertEqual(other["stem"], "")

    def test_current_organizer_raw_preserves_qid_and_local_history(self):
        raw = self.folder / "organizer.md"
        raw.write_text(
            "#### 합성 과목\n"
            "<!-- question qid=q-fixture topic=fixture year=2025 -->\n"
            "**Q. 88.** [25본1_1차]\n"
            "<!-- body q-fixture start -->\n"
            "합성 문항\n① 가\n② 나\n"
            "<!-- body q-fixture end -->\n<!-- question q-fixture end -->\n",
            encoding="utf-8",
        )
        index = self.folder / "raw-index.json"
        self.cli(QUESTION, "build", "--source", raw, "--out", index)
        packet = self.cli(QUESTION, "lookup", "--index", index, "--qid", "q-fixture")
        self.assertEqual(packet["question"]["history"], "25년 1차 88번")
        self.assertEqual(packet["question"]["stem"], "합성 문항")
        self.assertEqual(packet["question"]["location"], "줄 2–9")

    def test_evidence_role_caps_keep_conflict_and_hide_part_two(self):
        lecture = self.folder / "raw.md"
        script = self.folder / "script.md"
        figure = self.folder / "plot.png"
        figure.write_bytes(b"synthetic figure bytes")
        lecture.write_text("<!-- page 11 -->\n알파 판정은 A다.\n![[plot.png|600]]\n"
                           "<!-- page 12 -->\n" + "무관한 세부 내용\n" * 50
                           + "…" * 900 + "알파 예외 설명\n" + "알파 추가 자료\n" * 12, encoding="utf-8")
        script.write_text("## PART 1 — 교정 스크립트\n[Page 11?]\n"
                          "[00:03:12] 알파는 정정합니다. 최종 판정은 B입니다.\n"
                          "교수의 다음 설명\n"
                          "## PART 2 — 정리본\n알파 C라는 오해.\n", encoding="utf-8")
        index = self.folder / "evidence.json"
        self.cli(EVIDENCE, "build", "--lecture", lecture, "--script", script, "--out", index)
        hits = self.cli(EVIDENCE, "search", "--index", index, "--term", "알파", "--limit", 2)
        raw_result = hits["lecture"]["results"]
        script_result = hits["script"]["results"]
        self.assertEqual(len(raw_result), 2)
        self.assertTrue(hits["lecture"]["matches_limited"])
        self.assertEqual(raw_result[0]["page"], "11")
        self.assertEqual(Path(raw_result[0]["figures"][0]["file"]), figure)
        self.assertTrue(any(item.get("clipped") for item in raw_result[1]["excerpts"]))
        self.assertIn("p.11? / 00:03:12 / 줄 3", script_result[0]["location"])
        self.assertIn("최종 판정은 B", script_result[0]["excerpts"][0]["text"])
        self.assertNotIn("C라는 오해", json.dumps(hits, ensure_ascii=False))
        self.assertLess(len(json.dumps(hits, ensure_ascii=False)), 3500)
        targeted = self.cli(EVIDENCE, "search", "--index", index, "--file", script.name, "--at", 3)
        self.assertEqual(targeted["lecture"]["status"], "out_of_scope")
        self.assertEqual(targeted["script"]["status"], "found")
        script.write_text("## PART 2 — 정리본\n알파 C라는 오해.\n", encoding="utf-8")
        self.cli(EVIDENCE, "search", "--index", index, "--term", "알파", code=2)
        self.cli(EVIDENCE, "build", "--lecture", lecture, "--script", script, "--out", index)
        fresh = self.cli(EVIDENCE, "search", "--index", index, "--term", "알파")
        self.assertTrue(fresh["script"]["sources"][0]["summary_only"])
        self.assertEqual(fresh["script"]["status"], "no_match")

    def test_bank_emits_source_callouts_and_qc_detects_missing_figure(self):
        doc = word_fixture(self.folder)
        index = self.folder / "questions.json"
        self.cli(QUESTION, "build", "--source", doc, "--out", index)
        scope = self.cli(BANK, "list", "--index", index, "--limit", 1)
        self.assertEqual(scope["count"], 2)
        self.assertEqual(len(scope["questions"]), 1)
        answers = self.folder / "answers.json"
        answer = {"comment_lines": ["**수업 주제** 합성 분류", "**시험 포인트** 분류 조건",
                                    "**자료 안내** 원문 연도 유지"], "answers": []}
        for qid in ("q001", "q002"):
            answer["answers"].append(
                {"qid": qid, "status": "정답 불확실", "answer_lines": [], "gap_lines": ["합성 자료 부족."],
                 "options": [{"marker": "①", "verdict": "?", "reason": "첫 조건을 확인할 수 없다."},
                             {"marker": "②", "verdict": "?", "reason": "둘째 조건을 확인할 수 없다."}],
                 "concept_lines": ["분류 기준과 관찰 내용을 대응시킨다."],
                 "evidence": [{"role": "lecture", "source": "합성 교안.md",
                               "location": "p.3 / 줄 4", "text": "분류 예시"}]}
            )
        answers.write_text(json.dumps(answer, ensure_ascii=False), encoding="utf-8")
        md = self.folder / "bank.md"
        self.cli(BANK, "emit", "--index", index, "--answers", answers, "--out", md)
        output = md.read_text(encoding="utf-8")
        self.assertIn("> > [!quote] 공통지문 — 1차", output)
        self.assertIn("> > [!quote] 보기\n> > 가) 합성 조건", output)
        self.assertIn("\n\n> [!answer]- 해설", output)
        self.assertIn("> > **정답 불확실**", output)
        self.assertIn("> > [!cite]- 근거", output)
        self.assertEqual(output.count("|400]]"), 2)
        self.cli(BANK, "qc", "--index", index, "--answers", answers, "--out", md)
        md_index = self.folder / "md-index.json"
        self.cli(QUESTION, "build", "--source", md, "--out", md_index)
        second = self.cli(QUESTION, "lookup", "--index", md_index, "--qid", "q002")["question"]
        self.assertEqual(second["body"][0]["text"], "[Figure 2] 합성 알파 장치의 위치")
        self.assertEqual(second["body"][1]["kind"], "image")
        self.assertEqual(second["body"][2]["text"], "뒤 회차에서 확인하는 별도 발문이다.")
        attachment = next((md.parent / "동봉 이미지").iterdir())
        displaced = md.parent / "keep-qa-figure.png"
        attachment.rename(displaced)
        self.cli(BANK, "qc", "--index", index, "--answers", answers, "--out", md, code=2)
        self.assertFalse(attachment.exists())
        self.assertTrue(displaced.exists())

    def test_index_refuses_to_replace_an_unrelated_existing_file(self):
        source = word_fixture(self.folder)
        out = self.folder / "unrelated.json"
        out.write_text('{"keep": true}', encoding="utf-8")
        self.cli(QUESTION, "build", "--source", source, "--out", out, code=2)
        self.assertEqual(out.read_text(encoding="utf-8"), '{"keep": true}')


if __name__ == "__main__":
    unittest.main()
