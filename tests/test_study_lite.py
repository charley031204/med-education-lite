"""Synthetic study fixtures; artifacts remain under .work/study-lite for review."""
import copy
import json
import os
from pathlib import Path
import sys
import unittest
import uuid
from unittest.mock import patch
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / 'skills/transcript-pdf-study/scripts'
sys.path.insert(0, str(SCRIPTS))
import prepare_study
import render_study


class StudyTests(unittest.TestCase):
    def setUp(self):
        self.work = ROOT / '.work/study-lite' / uuid.uuid4().hex
        self.work.mkdir(parents=True)
        (self.work / 'transcript.md').write_text('# Example\n## PART 1 — 교정 스크립트\n교수: 검체 A와 B의 차이는 시험에 출제합니다.\n## PART 2 — 강의 정리본\n검체 A는 조건 1에서 관찰한다.\n\n---\n\n### 비교\n검체 B는 조건 2에서 관찰할 수 있다.\n', encoding='utf-8')
        (self.work / 'lecture.md').write_text('조건 1과 조건 2를 비교한다.\n그림 1. 검체 비교 도식\n', encoding='utf-8')
        (self.work / 'questions.md').write_text('2024 Q1. 검체 A와 B의 차이는?\n', encoding='utf-8')
        self.yellow = {'kind': 'question', 'source': 'questions.md', 'quote': '검체 A와 B의 차이는?', 'locator': '2024 Q1', 'match_reason': '문항이 검체 A와 B의 조건 차이를 묻는다.'}
        self.green = {'kind': 'professor_exam', 'source': 'transcript.md', 'quote': '교수: 검체 A와 B의 차이는 시험에 출제합니다.', 'locator': 'PART 1 첫 발언', 'explicit': True, 'match_reason': '교수가 같은 비교를 시험 범위로 직접 지목했다.'}
        ground = {'source': 'transcript.md', 'quote': '검체 A는 조건 1에서 관찰한다.', 'locator': 'PART 2'}
        self.data = {'title': '검체 비교 학습 정리', 'transcript': 'transcript.md', 'lecture': 'lecture.md', 'question_sources': ['questions.md'], 'include_images': False,
                     'keywords': [{'runs': [{'text': '검체 A'}], 'grounding': [ground]}],
                     'key_summary': [{'runs': [{'text': '검체 A는 조건 1에서 관찰한다.'}], 'grounding': [ground]}],
                     'sections': [{'title': '검체 관찰 조건', 'blocks': [
                         {'type': 'paragraph', 'runs': [{'text': '검체 A는 조건 1에서 관찰한다.', 'evidence': [self.yellow, self.green]}]},
                         {'type': 'paragraph', 'runs': [{'text': '검체 B는 조건 2에서 관찰할 수 있다.'}]},
                         {'type': 'table', 'rows': [[[{'text': '검체'}], [{'text': '관찰 조건'}]], [[{'text': '검체 A'}], [{'text': '조건 1'}]], [[{'text': '검체 B'}], [{'text': '조건 2'}]]]}]}]}

    def test_pt2_keeps_internal_rules(self):
        extracted = prepare_study.extract_pt2((self.work / 'transcript.md').read_text(encoding='utf-8'))
        self.assertIn('---', extracted)
        self.assertIn('### 비교', extracted)
        self.assertNotIn('교수:', extracted)
        with self.assertRaises(ValueError):
            prepare_study.extract_pt2('## PART 2\nA\n## PART 2\nB')
        with self.assertRaises(ValueError):
            prepare_study.extract_pt2('## PART 2\nA\n## PART 3\nB')

    def test_evidence_and_grounding_fail_closed(self):
        render_study.validate(self.data, self.work)
        for key, value in [('quote', '없는 근거'), ('explicit', False), ('source', 'lecture.md')]:
            broken = copy.deepcopy(self.data)
            broken['sections'][0]['blocks'][0]['runs'][0]['evidence'][1][key] = value
            with self.assertRaises(ValueError):
                render_study.validate(broken, self.work)
        broken = copy.deepcopy(self.data)
        broken['question_sources'] = []
        with self.assertRaises(ValueError):
            render_study.validate(broken, self.work)
        broken = copy.deepcopy(self.data)
        broken['keywords'][0]['grounding'] = []
        with self.assertRaises(ValueError):
            render_study.validate(broken, self.work)

    def test_docx_editable_highlights_and_font(self):
        render_study.validate(self.data, self.work)
        path = self.work / 'study.docx'
        render_study.build_docx(self.data, self.work, path)
        with ZipFile(path) as archive:
            xml = archive.read('word/document.xml').decode()
            self.assertIn('FFF8D9', xml)
            self.assertIn('E7F3E7', xml)
            self.assertIn('[기출]', xml)
            self.assertIn('[교수 시험 언급]', xml)
            self.assertIn('검체 B는 조건 2에서 관찰할 수 있다.', xml)
            self.assertIn('Malgun Gothic', xml)
            self.assertNotIn('w:documentProtection', archive.read('word/settings.xml').decode())

    def test_image_permission_and_missing_caption(self):
        from PIL import Image
        Image.new('RGB', (600, 200), 'white').save(self.work / 'figure.png')
        block = {'type': 'image', 'path': 'figure.png', 'caption': '그림 1. 검체 비교 도식', 'alt': '검체 비교', 'source': {'source': 'lecture.md', 'quote': '그림 1. 검체 비교 도식', 'locator': '그림 1'}}
        self.data['sections'][0]['blocks'].append(block)
        with self.assertRaises(ValueError):
            render_study.validate(self.data, self.work)
        self.data['include_images'] = True
        render_study.validate(self.data, self.work)
        del block['caption']
        with self.assertRaises(ValueError):
            render_study.validate(self.data, self.work)

    def test_input_and_original_paths_are_protected(self):
        input_path = self.work / 'study.json'
        self.data['originals'] = ['original.pdf']
        with self.assertRaises(ValueError):
            render_study.protect_sources(self.data, input_path, self.work / 'original', self.work / 'qa')
        with self.assertRaises(ValueError):
            render_study.protect_sources(self.data, input_path, self.work / 'out', self.work)
        render_study.protect_sources(self.data, input_path, self.work / 'out', self.work / 'qa')

    def test_old_pdf_cannot_pass_failed_conversion(self):
        from types import SimpleNamespace
        old_pdf = self.work / 'study.pdf'
        old_pdf.write_bytes(b'previous output' * 20)
        before = old_pdf.read_bytes()
        with patch.object(render_study.subprocess, 'run', return_value=SimpleNamespace(returncode=0, stdout='conversion failed', stderr='')):
            with self.assertRaises(RuntimeError):
                render_study.convert_pdf(self.work / 'study.docx', old_pdf, self.work, 'libreoffice', 'unused')
        self.assertEqual(old_pdf.read_bytes(), before)

    def test_forced_line_breaks_are_rejected(self):
        self.data['sections'][0]['blocks'][0]['runs'][0]['text'] = '잘린\n문장'
        with self.assertRaises(ValueError):
            render_study.validate(self.data, self.work)

    @unittest.skipUnless(os.environ.get('STUDY_TEST_WORD') == '1', 'Set STUDY_TEST_WORD=1 for installed Word integration')
    def test_real_pdf_and_docx(self):
        from PIL import Image, ImageDraw
        image = Image.new('RGB', (900, 320), 'white')
        draw = ImageDraw.Draw(image)
        draw.rectangle((20, 40, 370, 280), outline='teal', width=6)
        draw.rectangle((530, 40, 880, 280), outline='navy', width=6)
        draw.text((160, 150), 'A', fill='black', font_size=50)
        draw.text((670, 150), 'B', fill='black', font_size=50)
        image.save(self.work / 'figure.png')
        self.data['include_images'] = True
        self.data['sections'].append({'title': '관찰 조건의 비교', 'blocks': [
            {'type': 'paragraph', 'runs': [{'text': '긴 영문과 한글이 같은 문단에 있어도 문장 흐름을 유지한다. Laboratory specimen observation protocol and measurement conditions는 문맥에 맞게 줄바꿈한다. ' * 4}]},
            {'type': 'image', 'path': 'figure.png', 'caption': '그림 1. 검체 비교 도식', 'alt': '왼쪽 검체 A, 오른쪽 검체 B', 'source': {'source': 'lecture.md', 'quote': '그림 1. 검체 비교 도식', 'locator': '그림 1'}}]})
        render_study.validate(self.data, self.work)
        (self.work / 'study.json').write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding='utf-8')
        docx, pdf = self.work / 'study.docx', self.work / 'study.pdf'
        render_study.build_docx(self.data, self.work, docx)
        render_study.convert_pdf(docx, pdf, self.work, 'word')
        report = render_study.inspect_pdf(pdf, self.work)
        report['pdf_sha256'] = render_study.sha(pdf)
        (self.work / 'qc.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        self.assertEqual(report['issues'], [])
        self.assertEqual(report['pages'], 2)
        import fitz
        import re
        from docx import Document
        with fitz.open(pdf) as rendered:
            actual_text = re.sub(r'\s+', '', ''.join(page.get_text() for page in rendered))
        editable = Document(docx)
        paragraphs = list(editable.paragraphs)
        for table in editable.tables:
            for row in table.rows:
                for cell in row.cells:
                    paragraphs.extend(cell.paragraphs)
        for paragraph in paragraphs:
            self.assertIn(re.sub(r'\s+', '', paragraph.text), actual_text)
        self.assertFalse(list(self.work.glob('*.html')))
        print('Study real-artifact QA:', self.work)


if __name__ == '__main__':
    unittest.main()
