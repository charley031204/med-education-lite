import contextlib
import csv
from hashlib import sha256
import io
import json
import os
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from docx import Document
from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
SCRIPT_DIR = ROOT / "skills" / "yama-chapter-word" / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

import build_yama_docx as builder
import shared_photos
from yama_bank import candidate_packet, index_bank, main as bank_main, select_chapters, write_json
from yama_layout import iter_paragraphs


def record(qid, topic, year, number, body, prof="원문 미표기"):
    return "\n".join([
        f"<!-- question qid={qid} topic={topic} year={year} -->",
        f"**Q. {number}.** [{year} 예시시험] [LLM 임의분류]",
        "", f"*과거 교수: {prof}*", "", f"<!-- body {qid} start -->",
        body, f"<!-- body {qid} end -->", f"<!-- question {qid} end -->", "",
    ])


def word_required():
    return os.environ.get("YAMA_TEST_WORD") == "1"


class YamaLiteTests(unittest.TestCase):
    def setUp(self):
        work = ROOT / ".work" / "yama-word" / "tests"
        work.mkdir(parents=True, exist_ok=True)
        self.work = Path(tempfile.mkdtemp(prefix="case-", dir=work))
        self.raw = self.work / "fixture_전연도_rawmd.md"
        self.csv = self.work / "fixture_questions.csv"
        self.index_path = self.work / "index.json"
        self.selection_path = self.work / "selection.json"
        self.output = self.work / "output"
        self.args = SimpleNamespace(
            rawmd=str(self.raw), index=str(self.index_path), selection=str(self.selection_path),
            report=str(self.work / "report.json"), preview_pdf=None, topic=None,
            include=None, exclude=None, year_prof=None, photo_review=None,
            extras=None, comments=None, merge_map=None, answers=None, page_map=None,
            context_box=None, sort=None, out=str(self.output), outname="fixture.docx",
            subject="예시 과목", exam="1차", ver="1",
        )

    def index(self, *selected):
        bank = index_bank(self.raw)
        write_json(self.index_path, bank)
        write_json(self.selection_path, {
            "raw_sha256": bank["raw_sha256"], "csv_sha256": bank["csv_sha256"],
            "chapters": [{"topic": "주제A", "current": "담당자 / 공동담당자", "source_ids": list(selected)}],
        })
        return bank

    def build(self):
        with contextlib.redirect_stdout(io.StringIO()):
            if word_required():
                self.args.preview_pdf = str(self.work / "preview.pdf")
                return builder.build(self.args)
            with patch("yama_layout.measure_word", return_value={"available": False, "status": "UNAVAILABLE"}):
                return builder.build(self.args)

    def test_scoped_retrieval_and_content_verification(self):
        stem = "**특정 성분의 검사 조건은?**"
        no = lambda text: f"{stem}\n\n1) {text}\n2) 원문 유지"
        raw = "\n".join([
            "#### 주제A — 담당자",
            record("q1", "T01", 2025, "11", no("표본 A")),
            record("q2", "T01", 2023, "12", no("표본 A")),
            record("q3", "T01", 2024, "13", no("표본 B"), "공동담당자"),
            "#### 주제B — 이전담당자",
            record("q4", "T02", 2025, "11", "교안 고유어를 고르시오.\n1) 다른 선택지", "이전담당자"),
            record("q5", "T02", 2024, "15", "별도 범위 원문"),
        ])
        self.raw.write_text(raw, encoding="utf-8")
        with self.csv.open("w", encoding="utf-8", newline="") as handle:
            fields = ["qid", "topic_id", "topic", "topic_candidates", "exam_round", "prof_past"]
            rows = [("q1", "T01", "주제A", ["T01"], "", ""),
                    ("q2", "T01", "주제A", [], "2차", ""),
                    ("q3", "T01", "주제A", ["T01"], "", ""),
                    ("q4", "T02", "주제B", ["T01"], "", "이전담당자"),
                    ("q5", "T02", "주제B", ["T02"], "", "")]
            writer = csv.writer(handle)
            writer.writerow(fields)
            writer.writerows((qid, tid, topic, json.dumps(hints), round_name, prof)
                             for qid, tid, topic, hints, round_name, prof in rows)
        bank = self.index("q1", "q2", "q3", "q4")
        packet = candidate_packet(bank, "T01", [], ["교안 고유어"], [], False, 0, 2)
        self.assertEqual(packet["n_bank"], 5)
        self.assertEqual(packet["n_candidates"], 4)
        self.assertEqual(packet["next_offset"], 2)
        packet = candidate_packet(bank, "T01", [], ["교안 고유어"], [], False, 2, 2)
        self.assertEqual([q["id"] for q in packet["candidates"]], ["q3", "q4"])
        self.assertIn("csv_candidate", packet["candidates"][1]["reasons"])
        self.assertIn("lecture_term", packet["candidates"][1]["reasons"])
        self.assertEqual(bank["csv_links"], {"linked": 5})
        result = self.build()
        self.assertTrue(result["source_content_ok"], result)
        self.assertEqual(result["n_source_questions"], 4)
        self.assertEqual(result["expected_blocks"], 3)
        self.assertEqual(result["chapters"][0]["n_other_prof"], 1)
        doc = Document(result["output"])
        lines = [p.text for p in iter_paragraphs(doc)]
        self.assertIn("[25,23] 특정 성분의 검사 조건은?", lines)
        self.assertIn("[24] 특정 성분의 검사 조건은?", lines)
        self.assertIn("※ 출제 이력: 25년 11번 · 23년 2차 12번", lines)
        self.assertIn("1) 표본 A", lines)
        self.assertIn("1) 표본 B", lines)
        self.assertIn("타 교수 야마", lines)
        self.assertNotIn("별도 범위 원문", lines)
        if word_required():
            self.assertEqual(result["verdict"], "OK", result["layout"]["word"])
            self.assertTrue(Path(result["preview"]["path"]).is_file())
        else:
            self.assertEqual(result["verdict"], "CHECK")
        self.raw.write_text(raw + "\n변경", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "bank changed"):
            self.build()

    def test_source_paths_and_stale_csv_are_rejected(self):
        raw = "#### 주제A — 담당자\n**Q. 1. 예시** [25본1_1차]\n\n① 유지\n"
        self.raw.write_text(raw, encoding="utf-8")
        with contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit):
                bank_main(["index", "--rawmd", str(self.raw), "--out", str(self.raw)])
        self.assertEqual(self.raw.read_text(encoding="utf-8"), raw)
        self.csv.write_text("topic,topic_id\n", encoding="utf-8")
        self.index("raw-c0001-q00001")
        self.args.report = str(self.csv)
        with self.assertRaisesRegex(ValueError, "Output path conflicts"):
            self.build()
        self.assertEqual(self.csv.read_text(encoding="utf-8"), "topic,topic_id\n")
        self.csv.write_text("topic,topic_id,stem\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "CSV changed"):
            self.build()

    def test_bad_header_and_repeated_selected_id_are_rejected(self):
        self.raw.write_text("#### 주제A — 담당자\n**Q. 1. 예시** [25본1_1차]\n\n① 유지\n"
                            "**Q. 2. 예시** [알 수 없음]\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "headers did not all parse"):
            index_bank(self.raw)
        self.raw.write_text("#### 주제A — 담당자\n**Q. 1. 예시** [25본1_1차]\n\n① 유지\n", encoding="utf-8")
        bank = self.index("raw-c0001-q00001", "raw-c0001-q00001")
        with self.assertRaisesRegex(ValueError, "repeated selected question"):
            select_chapters(bank, self.selection_path)

    def test_photo_approval_binds_original_leaves_and_adjacent_images(self):
        image_dir = self.work / "동봉 이미지"
        image_dir.mkdir()
        images = {}
        for name, color in [("a.png", "red"), ("b.png", "green"), ("c.png", "blue")]:
            path = image_dir / name
            graphic = Image.new("RGB", (400, 180), "white")
            draw = ImageDraw.Draw(graphic)
            draw.rectangle((30, 30, 320, 150), fill=color)
            draw.text((80, 80), name, fill="white")
            graphic.save(path)
            images[name] = sha256(path.read_bytes()).hexdigest()
        alias = image_dir / "copy.png"
        alias.write_bytes((image_dir / "a.png").read_bytes())
        raw = "\n".join([
            "#### 주제A — 담당자",
            record("q1", "T01", 2025, "11", "같은 사진의 첫 문항\n[Figure A]\n![[동봉 이미지/a.png]]\n① 보존"),
            record("q2", "T01", 2023, "11", "같은 사진의 다른 문항\n[Figure A]\n![[동봉 이미지/copy.png]]\n① 유지"),
            record("q3", "T01", 2022, "11", "별도 사진 두 장을 비교\n[Figure B]\n![[동봉 이미지/b.png]]\n"
                                            "[Figure C]\n![[동봉 이미지/c.png]]\n② 원문"),
        ])
        self.raw.write_text(raw, encoding="utf-8")
        self.index("q1", "q2", "q3")
        preflight = self.build()
        self.assertEqual(preflight["verdict"], "CHECK")
        self.assertIsNone(preflight["output"])
        review = Path(preflight["shared_photos"]["review_file"])
        draft = json.loads(review.read_text(encoding="utf-8"))
        all_ids = shared_photos._byte_representatives(draft["inventory"]["files"])
        signature = dict(reviewer="Fixture primary", independent_reviewer="Fixture secondary",
                         rationale="Known synthetic colors", independent_rationale="Known synthetic labels")
        draft["discovery_review"].update(signature, viewed_file_ids=all_ids, independent_viewed_file_ids=all_ids)
        for pair in draft["pairs"]:
            pair.update(signature, verdict="different", independent_verdict="different")
        approved = self.work / "fixture-photo-review.json"
        write_json(approved, draft)
        self.args.photo_review = str(approved)
        result = self.build()
        self.assertTrue(result["source_content_ok"], result["shared_photos"])
        if word_required():
            self.assertTrue(result["shared_photos"]["source_reconstruction_ok"], result)
        else:
            self.assertIn("A shared photo header is not measured with its own first source leaf",
                          result["shared_photos"]["issues"])
        self.assertEqual(result["n_source_questions"], 3)
        self.assertEqual(result["expected_blocks"], 3)
        self.assertEqual(result["images_rendered_before_layout"], 3)
        doc = Document(result["output"])
        self.assertEqual(len(doc.inline_shapes), 3)
        self.assertTrue(any(len(table.rows) == 1 and len(table.columns) == 2
                            and table.cell(0, 0).paragraphs[0].text == "[Figure B]"
                            and table.cell(0, 1).paragraphs[0].text == "[Figure C]" for table in doc.tables))
        self.assertEqual([q["id"] for q in draft["inventory"]["questions"]], ["q1", "q2", "q3"])
        self.assertEqual(sha256((image_dir / "a.png").read_bytes()).hexdigest(), images["a.png"])
        self.assertEqual(sha256(alias.read_bytes()).hexdigest(), images["a.png"])
        if word_required():
            self.assertEqual(result["verdict"], "OK", result["layout"]["word"])
            self.assertTrue(Path(result["preview"]["path"]).is_file())
        else:
            self.assertEqual(result["verdict"], "CHECK")


if __name__ == "__main__":
    unittest.main()
