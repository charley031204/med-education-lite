"""Exercise the local CLI against small PDF and PowerPoint fixtures."""

from pathlib import Path
from tempfile import mkdtemp
import json
import subprocess
import sys
import unittest

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.util import Inches, Pt
import pymupdf


ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "skills/pdf-to-rawmd/scripts/rawmd.py"
FIGURE_BLOCK = ("> **[Figure {figure}]**\n"
                "> Fixture image.\n>\n> **구성 요소:**\n> - 없음\n>\n"
                "> **흐름/관계:**\n> - 없음\n>\n> **범례/수치:**\n> - 없음")


def make_pdf(path, image):
    with pymupdf.open() as pdf:
        first = pdf.new_page(width=320, height=240)
        first.insert_text((20, 28), "Title", fontname="hebo", fontsize=20)
        first.insert_text((20, 55), "# literal * _ | <test> $5", fontsize=10)
        first.insert_text((20, 72), "1. literal", fontsize=10)
        first.insert_text((20, 89), "---", fontsize=10)
        for x in (20, 100, 180):
            first.draw_line((x, 105), (x, 165))
        for y in (105, 135, 165):
            first.draw_line((20, y), (180, y))
        for text, point in (("A | B", (26, 125)), ("Q", (106, 125)),
                            ("C", (26, 155)), ("2", (106, 155))):
            first.insert_text(point, text, fontsize=10)
        first.insert_image(pymupdf.Rect(200, 105, 300, 205), filename=image)
        pdf.new_page(width=240, height=240).insert_image((0, 0, 240, 240), filename=image)
        rotated = pdf.new_page(width=320, height=240)
        rotated.insert_image((12, 20, 50, 70), filename=image)
        rotated.set_rotation(90)
        for _ in range(14):
            pdf.new_page()
        pdf.save(path)


def make_pptx(path, image):
    deck = Presentation()
    first = deck.slides.add_slide(deck.slide_layouts[5])
    first.shapes.title.text = "Native title"
    group = first.shapes.add_group_shape()
    textbox = group.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
    paragraph = textbox.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "Styled"
    run.font.bold = True
    run.font.italic = True
    run.font.size = Pt(20)
    literal_line = textbox.text_frame.add_paragraph()
    literal_line.text = "+ literal _ [brackets] | <tag>"
    formula = textbox.text_frame.add_paragraph()
    equation = etree.fromstring(
        b'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        b'<m:f><m:num><m:r><m:t>1</m:t></m:r></m:num>'
        b'<m:den><m:r><m:t>x</m:t></m:r></m:den></m:f></m:oMath>'
    )
    formula._p.append(equation)
    table = first.shapes.add_table(2, 2, Inches(1), Inches(4), Inches(5), Inches(1)).table
    table.cell(0, 0).merge(table.cell(0, 1))
    table.cell(0, 0).text = "A | B"
    table.cell(1, 0).text = "C\nD"
    table.cell(1, 1).text = "E"
    chart = CategoryChartData()
    chart.categories = ["item"]
    chart.add_series("Series", (2,))
    first.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6), Inches(1),
                           Inches(3), Inches(2), chart)
    picture = first.shapes.add_picture(str(image), Inches(6), Inches(4), Inches(1), Inches(1))
    picture.crop_left = 0.1
    first.notes_slide.notes_text_frame.text = "Speaker-only text"
    second = deck.slides.add_slide(deck.slide_layouts[5])
    second.shapes.title.text = "Hidden title"
    second._element.set("show", "0")
    # Presentation relationship order can differ from package member filenames.
    slide_ids = deck.slides._sldIdLst
    slide_ids.insert(0, slide_ids[-1])
    deck.save(path)


class RawmdCLI(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Keep the generated fixtures and outputs for inspection.
        test_root = ROOT / ".work" / "test-runtime"
        test_root.mkdir(parents=True, exist_ok=True)
        cls.run_dir = Path(mkdtemp(dir=test_root, prefix="run-"))
        cls.png = cls.run_dir / "fixture.png"
        Image.new("RGB", (30, 30), "#1e62a6").save(cls.png)
        cls.pdf = cls.run_dir / "sample [title].pdf"
        make_pdf(cls.pdf, cls.png)
        cls.pptx = cls.run_dir / "sample.pptx"
        make_pptx(cls.pptx, cls.png)

    def command(self, *args, code=0):
        run = subprocess.run([sys.executable, str(CLI), *map(str, args)], capture_output=True,
                             encoding="utf-8", timeout=30)
        self.assertEqual(run.returncode, code, run.stdout + run.stderr)
        return run

    def test_plan_metadata_and_range_bounds(self):
        result = self.command("plan", self.pdf)
        summary = json.loads(result.stdout)
        self.assertEqual(summary["total_pages"], 17)
        self.assertEqual([len(chunk["pages"]) for chunk in summary["chunks"]], [15, 2])
        self.assertNotIn("literal", result.stdout)
        invalid = self.run_dir / "invalid-draft.md"
        for spec in ("1-16", "0", "17-16", "18", "1,"):
            with self.subTest(spec=spec):
                self.command("draft", self.pdf, "--pages", spec, "--out", invalid, code=2)
        self.assertFalse(invalid.exists())

    def test_pdf_native_draft_and_scan_warning(self):
        job = self.run_dir / "pdf-native"
        draft = job / "chunk_01.md"
        packet = job / "chunk_01.txt"
        result = self.command("draft", self.pdf, "--pages", "1-2", "--out", draft, "--packet", packet)
        summary = json.loads(result.stdout)
        self.assertEqual(summary["status"], "UNREVIEWED")
        self.assertIn("possible scanned or image-only page", summary["review"]["2"])
        self.assertNotIn("literal", result.stdout)
        text = draft.read_text(encoding="utf-8")
        self.assertIn("## **Title**", text)
        self.assertIn(r"\# literal \* \_ | &lt;test&gt; \$5", text)
        self.assertIn(r"1\. literal", text)
        self.assertIn(r"\---", text)
        self.assertIn(r"| A \| B | Q |", text)
        self.assertIn("<!-- page 2 -->", text)
        self.assertIn("[불명확:", text)
        self.assertIn("NATIVE EXTRACTION — UNREVIEWED", packet.read_text(encoding="utf-8"))
        self.assertNotIn("raw:block", text)
        self.assertEqual({path.name for path in job.iterdir()}, {"chunk_01.md", "chunk_01.txt"})

    def test_native_slide_order_group_table_math_and_notes(self):
        draft = self.run_dir / "pptx-draft.md"
        packet = self.run_dir / "pptx-packet.txt"
        summary = json.loads(self.command("draft", self.pptx, "--pages", "1-2",
                                          "--out", draft, "--packet", packet).stdout)
        self.assertEqual(summary["hidden_pages"], [1])
        self.assertIn("chart needs visual review", summary["review"]["2"])
        body = draft.read_text(encoding="utf-8")
        self.assertLess(body.index("Hidden title"), body.index("Native title"))
        self.assertIn("***Styled***", body)
        self.assertIn(r"\+ literal \_ \[brackets\] | &lt;tag&gt;", body)
        self.assertIn(r"$\frac{1}{x}$", body)
        self.assertIn(r"| A \| B | A \| B |", body)
        self.assertIn("| C<br>D | E |", body)
        self.assertIn("<!-- 병합: R1C1–R1C2 -->", body)
        compact = packet.read_text(encoding="utf-8")
        self.assertIn("[group ", compact)
        self.assertIn("[text local ", compact)
        self.assertNotIn("Speaker-only text", compact + body)
        notes = self.command("packet", self.pptx, "--pages", "2", "--notes").stdout
        self.assertIn("[speaker notes — separate]\nSpeaker-only text", notes)
        plan = json.loads(self.command("plan", self.pptx).stdout)
        self.assertEqual(plan["total_pages"], 2)
        native_figures = self.run_dir / "pptx-candidates"
        native = json.loads(self.command("figures", self.pptx, "--pages", "2",
                                         "--out-dir", native_figures).stdout)
        self.assertEqual(len(native["candidates"]), 1)
        with Image.open(native_figures / "page_002_a.png") as image:
            self.assertEqual(image.size, (30, 30))

    def test_render_candidate_and_crop(self):
        folder = self.run_dir / "rotated-render"
        self.command("render", self.pdf, "--pages", "3", "--box", "180,10,225,60",
                     "--dpi", "72", "--out-dir", folder)
        rendered = folder / "page_003.png"
        with Image.open(rendered) as image:
            self.assertEqual(image.size, (45, 50))
            self.assertEqual(image.convert("RGB").getpixel((17, 22)), (30, 98, 166))
        original = rendered.read_bytes()
        self.command("render", self.pdf, "--pages", "3", "--out-dir", folder, code=2)
        self.assertEqual(rendered.read_bytes(), original)
        self.command("render", self.pdf, "--pages", "1,3", "--box", "1,1,10,10",
                     "--out-dir", folder, code=2)
        cropped = folder / "sample_p003a.png"
        self.command("crop", rendered, "--box", "2,2,12,16", "--out", cropped)
        with Image.open(cropped) as image:
            self.assertEqual(image.size, (10, 14))
        self.command("crop", rendered, "--box", "1,1,99,90", "--out", folder / "out.png", code=2)
        result = self.command("figures", self.pdf, "--pages", "1,2", "--out-dir", self.run_dir / "pdf-candidates")
        candidates = json.loads(result.stdout)["candidates"]
        self.assertTrue(candidates)
        self.assertEqual({item["page"] for item in candidates}, {1})

    def test_merge_and_validate_with_existing_output_and_invalid_embed(self):
        folder = self.run_dir / "merge"
        assets = folder / "동봉 이미지"
        assets.mkdir(parents=True)
        first = folder / "chunk_01.md"
        second = folder / "chunk_02.md"
        final = folder / "sample_rawmd.md"
        body = "<!-- page 3 -->\n\n"
        for letter, caption in (("a", "3a / Caption [A]"), ("b", "3b")):
            name = f"sample_p003{letter}.png"
            (assets / name).write_bytes(self.png.read_bytes())
            body += FIGURE_BLOCK.format(figure=caption) + f"\n\n![[동봉 이미지/{name}|600]]\n\n"
        first.write_text(body, encoding="utf-8")
        second.write_text("<!-- page 5 -->\n\n*(빈 페이지)*\n", encoding="utf-8")
        args = ("merge", self.pdf, "--pages", "3,5", "--chunk", second, "--chunk", first,
                "--out", final, "--slug", "sample", "--date", "2026-11-12")
        merged = json.loads(self.command(*args).stdout)
        self.assertTrue(merged["mechanical_only"])
        self.assertEqual(merged["pages"], [3, 5])
        self.assertEqual((merged["figures"], merged["embeds"]), (2, 2))
        self.assertIn(r"원본 파일: sample \[title\].pdf", final.read_text(encoding="utf-8"))
        original = final.read_bytes()
        self.command(*args, code=2)
        self.assertEqual(final.read_bytes(), original)
        self.command("validate", self.pdf, final, "--pages", "3,5", "--slug", "sample")
        self.command("validate", self.pdf, final, "--pages", "2,5", code=1)
        (assets / "sample_p003b.png").write_bytes(b"broken")
        invalid = self.command("validate", self.pdf, final, "--pages", "3,5", code=1)
        self.assertIn("invalid_png", invalid.stdout)
        self.assertEqual(final.read_bytes(), original)
        wrong_link = folder / "wrong_link.md"
        wrong_link.write_text(body.replace("sample_p003a.png", "../sample_p003a.png"), encoding="utf-8")
        result = self.command("validate", self.pdf, wrong_link, "--pages", "3", code=1)
        self.assertIn("invalid_asset_filename", result.stdout)
        bad_marker = folder / "bad_marker.md"
        bad_marker.write_text("<!-- page 3 -->\n\n*(빈 페이지)*\n\n<!--Page 4-->\n", encoding="utf-8")
        result = self.command("validate", self.pdf, bad_marker, "--pages", "3", code=1)
        self.assertIn("malformed_page_marker", result.stdout)

    def test_draft_protects_both_existing_outputs(self):
        draft = self.run_dir / "do-not-overwrite.md"
        packet = self.run_dir / "do-not-overwrite.txt"
        packet.write_text("original packet\n", encoding="utf-8")
        self.command("draft", self.pdf, "--pages", "1", "--out", draft, "--packet", packet, code=2)
        self.assertFalse(draft.exists())
        self.assertEqual(packet.read_text(encoding="utf-8"), "original packet\n")
        self.command("draft", self.pdf, "--pages", "1", "--out", draft,
                     "--packet", draft / "nested.txt", code=2)
        self.assertFalse(draft.exists())

    def test_original_pdf_bytes_remain_unchanged(self):
        original = self.pdf.read_bytes()
        folder = self.run_dir / "immutable-source"
        commands = [
            ("plan", self.pdf),
            ("packet", self.pdf, "--pages", "1"),
            ("draft", self.pdf, "--pages", "1", "--out", folder / "chunk.md"),
            ("render", self.pdf, "--pages", "1", "--out-dir", folder / "render"),
            ("figures", self.pdf, "--pages", "1", "--out-dir", folder / "figures"),
        ]
        for command in commands:
            with self.subTest(command=command[0]):
                self.command(*command)
                self.assertEqual(self.pdf.read_bytes(), original)
        self.command("draft", self.pdf, "--pages", "1", "--out", self.pdf, code=2)
        self.command("merge", self.pdf, "--chunk", folder / "chunk.md",
                     "--out", self.pdf, "--pages", "1", code=2)
        self.assertEqual(self.pdf.read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
