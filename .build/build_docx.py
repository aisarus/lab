import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = ROOT.parent


def set_paragraph_rtl(paragraph, rtl=True):
    p_pr = paragraph._p.get_or_add_pPr()
    bidi = p_pr.find(qn('w:bidi'))
    if bidi is None:
        bidi = OxmlElement('w:bidi')
        p_pr.append(bidi)
    bidi.set(qn('w:val'), '1' if rtl else '0')


def set_run_language(run, language, rtl=False):
    r_pr = run._r.get_or_add_rPr()
    fonts = r_pr.rFonts
    if fonts is None:
        fonts = OxmlElement('w:rFonts')
        r_pr.insert(0, fonts)
    for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        fonts.set(qn(f'w:{attr}'), 'Arial')

    lang = r_pr.find(qn('w:lang'))
    if lang is None:
        lang = OxmlElement('w:lang')
        r_pr.append(lang)
    lang.set(qn('w:val'), language)
    if rtl:
        lang.set(qn('w:bidi'), language)
        rtl_node = r_pr.find(qn('w:rtl'))
        if rtl_node is None:
            rtl_node = OxmlElement('w:rtl')
            r_pr.append(rtl_node)
        rtl_node.set(qn('w:val'), '1')


def configure_document(document):
    section = document.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    normal = document.styles['Normal']
    normal.font.name = 'Arial'
    for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        normal._element.rPr.rFonts.set(qn(f'w:{attr}'), 'Arial')
    normal.font.size = Pt(11)
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.widow_control = True

    for style_name, font_size in [('Heading 1', 14), ('Heading 2', 12)]:
        style = document.styles[style_name]
        style.font.name = 'Arial'
        for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
            style._element.rPr.rFonts.set(qn(f'w:{attr}'), 'Arial')
        style.font.size = Pt(font_size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.keep_with_next = True


def add_heading(document, text, level, rtl):
    paragraph = document.add_paragraph(style=f'Heading {level}')
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_rtl(paragraph, rtl)
    run = paragraph.add_run(text)
    set_run_language(run, 'he-IL' if rtl else 'ru-RU', rtl)


def add_body_paragraph(document, text, rtl):
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_paragraph_rtl(paragraph, rtl)
    paragraph.paragraph_format.first_line_indent = Cm(0.5)
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(text)
    run.font.size = Pt(11)
    set_run_language(run, 'he-IL' if rtl else 'ru-RU', rtl)


def add_play_paragraph(document, text, rtl):
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT if rtl else WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_rtl(paragraph, rtl)

    if ':' in text:
        speaker, speech = text.split(':', 1)
        speaker_run = paragraph.add_run(speaker + ':')
        speaker_run.bold = True
        set_run_language(speaker_run, 'he-IL' if rtl else 'ru-RU', rtl)
        speech_run = paragraph.add_run(speech)
        set_run_language(speech_run, 'he-IL' if rtl else 'ru-RU', rtl)
    else:
        run = paragraph.add_run(text)
        set_run_language(run, 'he-IL' if rtl else 'ru-RU', rtl)


def build_document(data_path, output_path, rtl):
    data = json.loads(data_path.read_text(encoding='utf-8'))
    document = Document()
    configure_document(document)

    add_heading(document, data['question_1_title'], 1, rtl)
    page_break_after = data.get('question_1_page_break_after')
    for index, text in enumerate(data['question_1_paragraphs'], start=1):
        if page_break_after and index == page_break_after + 1:
            document.add_page_break()
        add_body_paragraph(document, text, rtl)

    document.add_page_break()
    add_heading(document, data['question_2_title'], 1, rtl)
    add_heading(document, data['historical_intro_title'], 2, rtl)
    for text in data['historical_intro_paragraphs']:
        add_body_paragraph(document, text, rtl)

    if data['play_paragraphs']:
        add_heading(document, data['play_title'], 2, rtl)
        for text in data['play_paragraphs']:
            add_play_paragraph(document, text, rtl)

    properties = document.core_properties
    properties.title = data['document_title']
    properties.author = 'Arseny Perel'
    document.save(output_path)
    print(output_path)


build_document(ROOT / 'content_he.json', PROJECT_ROOT / 'KVS_Hebrew.docx', True)
build_document(ROOT / 'content_ru.json', PROJECT_ROOT / 'KVS_Russian.docx', False)
