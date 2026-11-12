"""Exercise transcript chunking, focused retrieval, and assembly with synthetic text."""

import json
import re
import subprocess
import sys
import unittest
from pathlib import Path
from tempfile import mkdtemp


ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "skills/med-transcript-polish/scripts/transcript.py"
P1 = "## PART 1 — 교정 스크립트 (순수 정제)"
P2 = "## PART 2 — 강의 정리본 (녹음본 요약)"


class TranscriptCLI(unittest.TestCase):
    def setUp(self):
        test_root = ROOT / ".work/transcript-lite"
        test_root.mkdir(parents=True, exist_ok=True)
        self.run_dir = Path(mkdtemp(dir=test_root, prefix="run-"))

    def command(self, *args, code=0):
        proc = subprocess.run(
            [sys.executable, "-X", "utf8", str(CLI), *map(str, args)],
            capture_output=True,
            encoding="utf-8",
            timeout=30,
        )
        self.assertEqual(proc.returncode, code, proc.stdout + proc.stderr)
        return json.loads(proc.stdout) if code == 0 else proc.stderr

    def write(self, name, body, encoding="utf-8"):
        path = self.run_dir / name
        with path.open("w", encoding=encoding, newline="") as handle:
            handle.write(body)
        return path

    def test_long_transcript_reconstructs_exactly_and_keeps_timestamps_intact(self):
        timestamps = []
        rows = []
        for minute in range(60):
            stamp = f"[00:{minute:02d}:00 - 00:{minute:02d}:59]"
            timestamps.append(stamp)
            rows.append(f"{stamp}\r\n이 말의 의미는 같지만 사례 {minute}의 조건은 다릅니다.\r\n")
        standalone_stamp = "01:00:00.000 --> 01:00:30.000"
        timestamps.append(standalone_stamp)
        rows.append(f"{standalone_stamp}\r\n" + "인" * 627 + "\ufeff마무리\r\n")
        original = "".join(rows)
        source = self.write("긴 녹음.txt", original, encoding="utf-8-sig")
        work = self.run_dir / "chunks"
        result = self.command("split", source, work, "--max-chars", 137)
        pieces = [path.read_bytes().decode() for path in sorted((work / "raw").glob("*.txt"))]
        self.assertGreater(result["chunks"], 30)
        self.assertTrue(all(0 < len(piece) <= 137 for piece in pieces))
        self.assertTrue(all(not piece.endswith("\r") for piece in pieces))
        self.assertEqual("".join(pieces), original)
        for stamp in timestamps:
            self.assertEqual(sum(piece.count(stamp) for piece in pieces), 1, stamp)
        self.command("split", source, work, code=1)
        self.assertEqual(
            pieces, [path.read_bytes().decode() for path in sorted((work / "raw").glob("*.txt"))]
        )

    def test_lecture_search_reports_ties_and_disambiguates_with_additional_term(self):
        lecture = self.write(
            "강의.md",
            "# 교안\n\n<!-- page 7 -->\n## 분류\nAT1과 핵심 분류\n"
            "<!-- page 9 -->\n## 분류\nAT1과 다른 핵심 분류\n"
            "<!-- page 12 -->\nAT10\n![[AT1-file.png]]\n"
            "<!-- page 15 -->\n" + "고유 근거의 확인\n" * 10,
        )
        tied = self.command("lecture", lecture, "--terms", "AT1", "at1", "--limit", 1)
        self.assertEqual(tied["top_score_page_count"], 2)
        self.assertEqual(tied["total_matches"], 2)
        self.assertEqual(tied["candidates"][0]["matched_terms"], ["AT1"])
        clear = self.command("lecture", lecture, "--terms", "AT1", "다른 핵심 분류")
        self.assertEqual(clear["top_score_page_count"], 1)
        self.assertEqual(clear["candidates"][0]["page"], 9)
        absent = self.command("lecture", lecture, "--terms", "누락 후보")
        self.assertEqual(absent["candidates"], [])
        first = self.command("lecture", lecture, "--page", 15, "--max-chars", 32)
        rest = self.command("lecture", lecture, "--page", 15, "--offset", first["next_offset"])
        self.assertEqual(first["body"] + rest["body"], ("고유 근거의 확인\n" * 10).strip())
        self.assertIn("does not exist", self.command("lecture", lecture, "--page", 8, code=1))

    def assembly_inputs(self):
        source = self.write(
            "녹음.txt",
            "[00:00 - 00:10]\n여기서 이 수치가\n[00:10 - 00:20]\n낮을 수 있어요. 다음 예시는 3 mmol/L예요.",
        )
        work = self.run_dir / "work"
        self.command("split", source, work)
        self.write(
            "work/polished/0001.md",
            "<!-- join: start -->\n여기서 이 수치가 낮을 수 있어요.\n\n다음 예시는 3 mmol/L예요.\n",
        )
        lecture = self.write(
            "강의.md",
            "<!-- page 4 -->\n## 첫 예시\n<!-- page 5 -->\n## 추가 예시\n<!-- page 7 -->\n## 다음 예시\n",
        )
        part2 = self.write("파트2.md", "## 핵심\n\n---\n\n| 수치 | 설명 |\n| --- | --- |\n| 3 mmol/L | 예시 |\n")
        return source, work, part2, lecture

    def test_assembly_preserves_part2_and_body_on_off(self):
        source, work, part2, lecture = self.assembly_inputs()
        self.write(
            "work/pages/0001.tsv",
            "# unresolved\t다음 발화의 페이지 전환 위치를 특정하기 어려움\n"
            "1\t4\tconfirmed\t수치 변화 설명이 첫 예시와 일치\n"
            "3\t5\tskipped\t교수가 5쪽을 생략한다고 발언\n"
            "3\t7\tuncertain\t예시 주제는 맞으나 문장 경계가 모호함\n",
        )
        tagged = self.run_dir / "정제.md"
        untagged = self.run_dir / "정제-태그없음.md"
        on = self.command("assemble", source, work, part2, tagged, "--title", "강의", "--lecture", lecture)
        off = self.command(
            "assemble", source, work, part2, untagged, "--title", "강의", "--lecture", lecture, "--no-page-tags"
        )
        self.assertEqual(on["page_annotations"], {"confirmed": 1, "skipped": 1, "uncertain": 1})
        self.assertEqual(on["page_review"], "partial")
        self.assertEqual(on["first_unresolved_chunks"], ["0001"])
        self.assertEqual(off["page_review"], "off")
        text = tagged.read_text(encoding="utf-8")
        self.assertIn(f"{P1}\n\n[Page 4]\n여기서 이 수치가 낮을 수 있어요.", text)
        self.assertIn("[Page 5]\n(이 페이지는 교수님이 넘어가심)\n[Page 7?]", text)
        removed = re.sub(r"(?m)^(?:\[Page \d+\??\]|\(이 페이지는 교수님이 넘어가심\))\n", "", text)
        self.assertEqual(removed, untagged.read_text(encoding="utf-8"))
        self.assertEqual(text.split(P2, 1)[1], "\n\n" + part2.read_text(encoding="utf-8"))
        self.command("assemble", source, work, part2, untagged, "--title", "다른 제목", "--no-page-tags", code=1)
        self.assertEqual(untagged.read_text(encoding="utf-8"), removed)

    def test_assembly_rejects_missing_and_invalid_reviews_and_reports_no_match(self):
        source, work, part2, lecture = self.assembly_inputs()
        output = self.run_dir / "result.md"
        args = ("assemble", source, work, part2, output, "--title", "강의", "--lecture", lecture)
        self.assertIn("Missing page review", self.command(*args, code=1))
        self.write("work/pages/0001.tsv", "# reviewed\n1\t8\tconfirmed\t없는 번호\n")
        self.assertIn("unknown or repeated page", self.command(*args, code=1))
        self.write("work/pages/0001.tsv", "# reviewed\n2\t7\tconfirmed\t빈 줄\n")
        self.assertIn("nonempty body line", self.command(*args, code=1))
        self.write("work/pages/0001.tsv", "# reviewed\n1\t4\tconfirmed\t근거\n3\t4\tconfirmed\t중복\n")
        self.assertIn("unknown or repeated page", self.command(*args, code=1))
        self.assertFalse(output.exists())
        self.write("work/pages/0001.tsv", "# unresolved\t발화가 교안 어느 쪽인지 확인하기 어려움\n")
        result = self.command(*args)
        self.assertEqual(result["page_review"], "no-matches")
        self.assertEqual(result["unresolved_chunk_count"], 1)
        self.assertNotIn("[Page", output.read_text(encoding="utf-8"))

    def test_assembly_joins_cross_chunk_sentences_and_rejects_tampered_raw(self):
        start = "처음부터 충분히 긴 앞 문장을 쓴 뒤에 나오는 말입니다. " * 4
        middle = "음절을 이어서 정제할 수 있습니다. " * 6
        original = start + middle
        source = self.write("녹음.txt", original)
        work = self.run_dir / "work"
        report = self.command("split", source, work, "--max-chars", 128)
        self.assertGreater(report["chunks"], 1)
        for index in range(1, report["chunks"] + 1):
            raw = (work / "raw" / f"{index:04d}.txt").read_text(encoding="utf-8")
            mode = "start" if index == 1 else "space"
            self.write(f"work/polished/{index:04d}.md", f"<!-- join: {mode} -->\n{raw.strip()}\n")
        part2 = self.write("파트2.md", "## 강의 개요\n\n주제 메모\n")
        output = self.run_dir / "result.md"
        result = self.command("assemble", source, work, part2, output, "--title", "강의")
        self.assertEqual(result["page_review"], "off")
        self.assertIn(f"{P1}\n\n{original.strip()}\n\n---\n", output.read_text(encoding="utf-8"))
        self.write("work/raw/0001.txt", "원문 누락")
        error = self.command(
            "assemble", source, work, part2, self.run_dir / "tampered.md", "--title", "강의", code=1
        )
        self.assertIn("do not reconstruct", error)
        self.assertFalse((self.run_dir / "tampered.md").exists())


if __name__ == "__main__":
    unittest.main()
