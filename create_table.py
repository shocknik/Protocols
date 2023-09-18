"""Файл содержащий функцию создания таблицы для каждого вида испытаний"""

from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.text import font
from docx.oxml.ns import qn
from borders import set_cell_border

class Electric_test_table:
    
    
    def __init__(self, table_name, requrement: str, method: str, mean_req: str, mean_method: str, test_record: str) -> None:
        self.table_name = table_name
        self.requrement = requrement
        self.method = method
        self.mean_req = mean_req
        self.mean_method = mean_method
        self.test_record = test_record
    
    def create_row_for_test(self):
        """метод создающий строку для вида испытаний"""
        items = [
            self.test_record,
            self.requrement,
            self.method,
            self.mean_req,
            self.mean_method,
        ]
        cell = self.table_name.add_row().cells
        for i, item in enumerate(items):
            cell[i].text = str(item)
            cell[i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            cell[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    