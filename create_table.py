"""Файл содержащий функцию создания таблицы испытаний"""

from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.text import font
from docx.oxml.ns import qn
from borders import set_cell_border
from backend_read import border_form, border_around_cell, table_inner_border_vertical, func_calculate_cells

class Test_Table_Row:
    """Класс для создания строк в таблице испытаний
    """
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
        set_cell_border(cell[6], end={"sz": 6, "val": 'double', "space": "0"})
        set_cell_border(cell[0], start={"sz": 6, "val": "double", "space": "0"})
        
    
    def create_row_param(self):
        """Метод для создания строк с параметрами и условиями испытаний"""
        cell_param = self.table_name.add_row().cells
        cell_param[0].text = "Параметры образцов и условия испытаний\n\
- длина образца\n\
- температура выдержки в климатической камере\n\
- время выдержки в климатической камере\n\
- диаметр бухты"
        cell_criteria = self.table_name.add_row().cells
        cell_criteria[0].text = "Критерии годности:\n\
- внешний вид\n"

        set_cell_border(cell_param[1], top={"sz": 12, "val": "none", "color": "black", "space": "0"})
        set_cell_border(cell_criteria[1], top={"sz": 12, "val": 'hidden', "color": "black", "space": "0"})
        set_cell_border(cell_param[6], end={"sz": 6, "val": "double", "space": "0"})
        set_cell_border(cell_criteria[6], end={"sz": 6, "val": 'double', "space": "0"})
        set_cell_border(cell_param[0], start={"sz": 6, "val": "double", "space": "0"})
        set_cell_border(cell_criteria[0], start={"sz": 6, "val": 'double', "space": "0"})
        
class Test_Table:
    """Класс таблицы"""
    
    def __init__(self, table_name, test_record, requrement, method, mean_req, limit) -> None:
        self.table_name = table_name
        self.test_record = test_record
        self.requrement = requrement
        self.method = method
        self.mean_req = mean_req
        self.limit = limit
        
    def row_for_navigation(self):
        cell = self.table_name.add_row().cells
        func_calculate_cells(cell)
        
        
        
        
    
    
    
    def title_row(self, num, name):
        """
        Метод, который создает строку раздела испытаний
        num - порядковый номер раздела по ходу протокола
        name - наименование категории испытаний
        """
        cell = self.table_name.add_row().cells # Создание строки в таблице
        cell[0].merge(cell[6]) # объединений всех ячеек
        cell[0].text = str(num) + " " + str(name) # вставка текста
        cell[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # выравнивание текста по центру
        cell[0].paragraphs[0].runs[0].bold = True # выделение жирным
        cell[0].paragraphs[0].paragraph_format.space_after = Pt(0) # убрать отступ внизу
        cell[0].paragraphs[0].runs[0].font.size = Pt(12) # установить размер шрифта
        border_around_cell(cell[0])
        
    
    def create_simple_row(self, num):
        """Метод, который создает единичную строку под простые испытания"""      
        cell = self.table_name.add_row().cells
        cell[0].text = str(num)
        cell[0].text += str(self.test_record)
    
    
    def create_sample_par_row(self):
        """Метод, который создает доп строку под параметры образца"""
        pass
    
    
    def create_validity_criteria(self):
        """Метод, который создает доп строку под критерии-годности"""
        pass
    