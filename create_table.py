"""Файл содержащий функцию создания таблицы для каждого вида испытаний"""

from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.text import font
from docx.oxml.ns import qn
from borders import set_cell_border

class Electric_test_table:
    
    def create_el_resist(table_name, requrement: str, method: str):
        """Сопротивление жилы
        requrement - номер пункта требований
        method - номер пункта метода
        """
        items = [
            "Электрическое сопротивление токопроводящих жил, пересчи-танное на 1 км длины и темпера-туру 20 С, Ом",
            requrement,
            method,
            'Значение требования',
            'Допуск',
            'Измеренное значение',
            'Соответствует',
        ]
        cell = table_name.add_row().cells
        for i, item in enumerate(items):
            cell[i].text = str(item)
            cell[i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY