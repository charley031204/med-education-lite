"""Render a reviewed study JSON to editable DOCX and matching PDF. No HTML."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import uuid

from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

FONT = 'Malgun Gothic'
YELLOW = 'FFF8D9'
GREEN = 'E7F3E7'
SKILL = Path(__file__).resolve().parents[1]


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def protect_sources(data, input_path, output, work):
    base = input_path.parent
    protected = {input_path.resolve()}
    for key in ('transcript', 'lecture'):
        protected.add((base / data[key]).resolve())
    for path in data.get('question_sources', []) + data.get('originals', []):
        protected.add((base / path).resolve())
    def collect(value):
        if isinstance(value, dict):
            for key, item in value.items():
                if key in ('source', 'path') and isinstance(item, str):
                    protected.add((base / item).resolve())
                collect(item)
        elif isinstance(value, list):
            for item in value:
                collect(item)
    collect(data)
    destinations = {output.with_suffix('.docx').resolve(), output.with_suffix('.pdf').resolve()}
    if destinations & protected or any(work.resolve() == path or work.resolve() in path.parents for path in protected):
        raise ValueError('Output or work directory overlaps an input/original source; choose a separate path')


def source_evidence(ref, base, cache):
    if not isinstance(ref, dict) or not all(ref.get(k) for k in ('source', 'quote', 'locator')):
        raise ValueError('Evidence requires source, exact quote, and locator')
    path = (base / ref['source']).resolve()
    if path not in cache:
        cache[path] = path.read_text(encoding='utf-8-sig')
    if ref['quote'] not in cache[path]:
        raise ValueError(f'Evidence quote is absent from {path.name}: {ref["locator"]}')


def check_runs(runs, base, cache, context):
    if not isinstance(runs, list) or not runs:
        raise ValueError('runs must be a nonempty list')
    for run in runs:
        if not isinstance(run.get('text'), str) or not run['text'].strip():
            raise ValueError('Every run requires text')
        if '\n' in run['text'] or '\r' in run['text']:
            raise ValueError('Use separate paragraphs, not forced line breaks')
        for mark in run.get('evidence', []):
            source_evidence(mark, base, cache)
            if mark.get('kind') not in ('question', 'professor_exam'):
                raise ValueError('Evidence kind must be question or professor_exam')
            if not mark.get('match_reason'):
                raise ValueError('Highlighted content requires a semantic match_reason')
            source_path = (base / mark['source']).resolve()
            if mark['kind'] == 'question' and source_path not in context['questions']:
                raise ValueError('Yellow evidence must come from a declared past-question source')
            if mark['kind'] == 'professor_exam':
                if mark.get('explicit') is not True or not re.search(r'시험|출제|exam|test', mark['quote'], re.I):
                    raise ValueError('Green requires an explicit professor exam statement')
                if source_path != context['transcript'] or mark['quote'] not in context['pt1']:
                    raise ValueError('Green evidence must quote the professor from PART 1')


def validate(data, base):
    if not data.get('title') or not data.get('sections'):
        raise ValueError('title and sections are required')
    cache = {}
    # Validate the original transcript boundary before accepting the transformed input.
    from prepare_study import extract_pt2
    transcript_path = (base / data['transcript']).resolve()
    transcript = transcript_path.read_text(encoding='utf-8-sig')
    pt2 = extract_pt2(transcript)
    if not data.get('lecture'):
        raise ValueError('A corresponding lecture text source is required')
    (base / data['lecture']).read_text(encoding='utf-8-sig')
    context = {'questions': {(base / path).resolve() for path in data.get('question_sources', [])},
               'transcript': transcript_path, 'pt1': transcript[:-len(pt2)], 'pt2': pt2}
    grounding_paths = {transcript_path, (base / data['lecture']).resolve()}
    for key in ('keywords', 'key_summary'):
        if not data.get(key):
            raise ValueError(f'{key} must contain grounded entries')
        for entry in data[key]:
            check_runs(entry['runs'], base, cache, context)
            if not entry.get('grounding'):
                raise ValueError(f'{key} entry needs grounding')
            for ref in entry['grounding']:
                source_evidence(ref, base, cache)
                source_path = (base / ref['source']).resolve()
                if source_path not in grounding_paths or (source_path == transcript_path and ref['quote'] not in pt2):
                    raise ValueError('Keywords and summary must be grounded in PT2 or the corresponding lecture')
    for section in data['sections']:
        if not section.get('title') or not section.get('blocks'):
            raise ValueError('Each section requires title and blocks')
        for block in section['blocks']:
            kind = block.get('type')
            if kind in ('paragraph', 'bullet', 'heading'):
                check_runs(block['runs'], base, cache, context)
            elif kind == 'table':
                rows = block['rows']
                if not rows or len(rows[0]) < 2 or len(rows[0]) > 4:
                    raise ValueError('Tables require 2 to 4 columns; split dense tables')
                if any(len(row) != len(rows[0]) for row in rows):
                    raise ValueError('Table rows must have equal column counts')
                for row in rows:
                    for cell in row:
                        check_runs(cell, base, cache, context)
            elif kind == 'image':
                if data.get('include_images') is not True:
                    raise ValueError('Images require include_images=true from the user request')
                if not block.get('caption') or not block.get('alt') or not block.get('source'):
                    raise ValueError('Image requires caption, alt, and source')
                with Image.open(base / block['path']) as img:
                    img.verify()
                source_evidence(block['source'], base, cache)
            else:
                raise ValueError(f'Unsupported block type: {kind}')
    return cache


def set_font(run, size=None):
    run.font.name = FONT
    if size:
        run.font.size = Pt(size)
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), FONT)


def shade(run, color):
    element = OxmlElement('w:shd')
    element.set(qn('w:fill'), color)
    run._element.get_or_add_rPr().append(element)


def add_runs(paragraph, runs):
    for item in runs:
        kinds = {e['kind'] for e in item.get('evidence', [])}
        run = paragraph.add_run(item['text'])
        set_font(run)
        run.bold = bool(item.get('bold'))
        if 'question' in kinds:
            shade(run, YELLOW)
        elif 'professor_exam' in kinds:
            shade(run, GREEN)
        # The content keeps yellow; the adjacent green label preserves overlap.
        if kinds:
            for kind, label, color in [('question', ' [기출]', YELLOW), ('professor_exam', ' [교수 시험 언급]', GREEN)]:
                if kind in kinds:
                    label_run = paragraph.add_run(label)
                    set_font(label_run, 8)
                    shade(label_run, color)


def build_docx(data, base, path):
    global FONT
    FONT = data.get('font_family', 'Malgun Gothic')
    document = Document()
    document.core_properties.author = ''
    document.core_properties.title = data['title']
    page = document.sections[0]
    page.page_width, page.page_height = Mm(210), Mm(297)
    page.top_margin = page.bottom_margin = Mm(17)
    page.left_margin = page.right_margin = Mm(18)
    for name, size in [('Normal', 10.6), ('Title', 21), ('Heading 1', 16), ('Heading 2', 12), ('Caption', 9)]:
        style = document.styles[name]
        style.font.name = FONT
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), FONT)
        style.paragraph_format.space_after = Pt(7)
        style.paragraph_format.line_spacing = 1.5
    document.styles['Normal'].paragraph_format.widow_control = True
    for fonts in document.styles.element.iter(qn('w:rFonts')):
        for attribute in list(fonts.attrib):
            if 'theme' in attribute.lower():
                del fonts.attrib[attribute]
        for attribute in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            fonts.set(qn('w:' + attribute), FONT)
    for border in list(document.styles['Title'].element.iter(qn('w:pBdr'))):
        border.getparent().remove(border)
    document.add_paragraph(data['title'], 'Title')
    document.add_paragraph('핵심 키워드', 'Heading 2')
    for item in data['keywords']:
        add_runs(document.add_paragraph(), item['runs'])
    document.add_paragraph('핵심 정리', 'Heading 2')
    for item in data['key_summary']:
        add_runs(document.add_paragraph(), item['runs'])
    legend = document.add_paragraph()
    for label, color in [('기출 대응 내용  ', YELLOW), ('교수의 명시적 시험 언급', GREEN)]:
        run = legend.add_run(label)
        set_font(run, 9)
        shade(run, color)
    for index, section in enumerate(data['sections']):
        title = document.add_paragraph(section['title'], 'Heading 1')
        title.paragraph_format.page_break_before = index > 0 or bool(section.get('new_page'))
        for block in section['blocks']:
            kind = block['type']
            if kind in ('paragraph', 'bullet', 'heading'):
                paragraph = document.add_paragraph(style='Heading 2' if kind == 'heading' else 'Normal')
                if kind == 'bullet':
                    paragraph.paragraph_format.left_indent = Mm(4)
                    paragraph.add_run('• ')
                add_runs(paragraph, block['runs'])
            elif kind == 'table':
                table = document.add_table(rows=0, cols=len(block['rows'][0]))
                table.style = 'Table Grid'
                table.autofit = False
                for row_index, row in enumerate(block['rows']):
                    cells = table.add_row().cells
                    for cell, content in zip(cells, row):
                        cell.width = Mm(174 / len(row))
                        add_runs(cell.paragraphs[0], content)
                    trpr = cells[0]._tc.getparent().get_or_add_trPr()
                    trpr.append(OxmlElement('w:cantSplit'))
                    if row_index == 0:
                        trpr.append(OxmlElement('w:tblHeader'))
                document.add_paragraph().paragraph_format.space_after = Pt(2)
            elif kind == 'image':
                paragraph = document.add_paragraph()
                paragraph.paragraph_format.keep_with_next = True
                image_path = base / block['path']
                with Image.open(image_path) as img:
                    width = min(160, 170 * img.width / img.height)
                shape = paragraph.add_run().add_picture(str(image_path), width=Mm(width))
                shape._inline.docPr.set('descr', block['alt'])
                document.add_paragraph(block['caption'], 'Caption')
    document.save(path)


def convert_pdf(docx, pdf, work, converter, soffice=None):
    conversion_dir = work / ('conversion-' + uuid.uuid4().hex)
    conversion_dir.mkdir()
    fresh_pdf = conversion_dir / (docx.stem + '.pdf')
    if converter == 'word':
        if os.name != 'nt':
            raise RuntimeError('Word conversion requires Windows')
        command = [sys.executable, str(SKILL / 'scripts/word_to_pdf.py'), str(docx), str(fresh_pdf)]
    else:
        executable = soffice or shutil.which('soffice')
        if not executable:
            raise RuntimeError('LibreOffice is missing; supply --soffice or choose --converter word')
        profile = (work / ('lo-profile-' + uuid.uuid4().hex)).resolve()
        command = [str(executable), '-env:UserInstallation=' + profile.as_uri(), '--headless', '--convert-to', 'pdf', '--outdir', str(conversion_dir), str(docx)]
    result = subprocess.run(command, capture_output=True, text=True, timeout=180, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    (work / 'conversion.log').write_text(result.stdout + '\n' + result.stderr, encoding='utf-8')
    if result.returncode or not fresh_pdf.exists() or fresh_pdf.stat().st_size < 100:
        raise RuntimeError('DOCX to PDF conversion failed; inspect conversion.log. DOCX remains available.')
    shutil.copyfile(fresh_pdf, pdf)


def inspect_pdf(pdf, work, font_family='Malgun Gothic'):
    import fitz
    issues = []
    fonts = set()
    with fitz.open(pdf) as document:
        for index, page in enumerate(document):
            page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(work / f'page-{index + 1}.png')
            words = page.get_text('words')
            if not words and not page.get_images():
                issues.append(f'page {index + 1}: blank page')
            for word in words:
                if word[0] < 35 or word[2] > page.rect.width - 35 or word[1] < 20 or word[3] > page.rect.height - 20:
                    issues.append(f'page {index + 1}: text outside safe page area')
                    break
            fonts.update(font[3] for font in page.get_fonts())
        pages = len(document)
    normalized = re.sub(r'[^a-z0-9]', '', font_family.lower())
    if not any(normalized in re.sub(r'[^a-z0-9]', '', font.lower()) for font in fonts):
        issues.append(f'{font_family} font not found in PDF; check font substitution')
    return {'pages': pages, 'fonts': sorted(fonts), 'issues': issues, 'visual_review': 'pending', 'semantic_review': 'pending'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('--output', type=Path, required=True, help='Output path without extension')
    parser.add_argument('--work-dir', type=Path, required=True)
    parser.add_argument('--converter', choices=('word', 'libreoffice'), default='libreoffice')
    parser.add_argument('--soffice', type=Path)
    parser.add_argument('--overwrite', action='store_true')
    args = parser.parse_args()
    args.input, args.output, args.work_dir = args.input.resolve(), args.output.resolve(), args.work_dir.resolve()
    docx, pdf = args.output.with_suffix('.docx'), args.output.with_suffix('.pdf')
    if not args.overwrite and (docx.exists() or pdf.exists()):
        parser.error('Output exists; --overwrite requires a user-requested revision')
    data = json.loads(args.input.read_text(encoding='utf-8-sig'))
    protect_sources(data, args.input, args.output, args.work_dir)
    validate(data, args.input.parent)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.work_dir.mkdir(parents=True, exist_ok=True)
    build_docx(data, args.input.parent, docx)
    convert_pdf(docx, pdf, args.work_dir, args.converter, args.soffice)
    report = inspect_pdf(pdf, args.work_dir, data.get('font_family', 'Malgun Gothic'))
    report.update({'docx_sha256': sha(docx), 'pdf_sha256': sha(pdf), 'input_sha256': sha(args.input), 'drive_upload': 'not_attempted'})
    (args.work_dir / 'qc.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False))
    return 2 if report['issues'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
