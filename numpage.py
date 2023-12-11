from docx import Document
from docx.oxml import OxmlElement, ns
from docx.shared import Inches, Cm, Mm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
from datetime import datetime
def create_element(name):
    return OxmlElement(name)


def create_attribute(element, name, value):
    element.set(ns.qn(name), value)


def add_page_number(paragraph, field):
    page_num_run = paragraph.add_run()
    # обозначаем начало позиции вывода
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')
    # задаем вывод текущего значения страницы PAGE (всего страниц NUMPAGES)
    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:separate', 'preserve')
    instrText.text = field
    # обозначаем конец позиции вывода
    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')
    # добавляем все в наш параграф (который формируется динамически)
    page_num_run._r.append(fldChar1)
    page_num_run._r.append(instrText)
    page_num_run._r.append(fldChar2)


