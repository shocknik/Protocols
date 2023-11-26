from setting import test_center, ratification, title_test_center, adress
from borders import set_cell_border
from docx import Document
from docx.shared import Inches, Cm, Mm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
class Protocol:
    
    
    def __init__(self, path, test_center) -> None:
        self.path = path
        self.test_center = test_center
        
    def create_new_file(self):
        doc = Document()
        doc.save(self.path)
        
        
    def create_title_list(self):
        doc = Document(self.path)
        if self.test_center == 'KT':
            i = 0
        elif self.test_center =='SK':
            i = 1
        """Задание альбомной формы, шрифта и отступов"""
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        title = doc.sections[0]
        title.orientation = WD_ORIENTATION.LANDSCAPE
        title.page_height = Cm(21.0)
        title.page_width = Cm(29.7)
        """формилование отступов полей"""
        title.left_margin = Mm(15)
        title.right_margin = Mm(13)
        title.top_margin = Mm(15)
        title.bottom_margin = Mm(13)
        """Создание таблицы для рамки листа"""
        title_table = doc.add_table(rows=7, cols=3)
        cell_header = title_table.cell(0, 1)
        """Название Испытательного центра"""
        cell_header.paragraphs[0].add_run(title_test_center[i]).font.size = Pt(14)
        cell_header.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell_header.width = Cm(23)
        cell_header.paragraphs[0].paragraph_format.space_after = Pt(0)
        if i == 0:
            cell_header_2 = title_table.cell(1, 1)
            cell_header_2.paragraphs[0].add_run('ИСПЫТАТЕЛЬНЫЙ ЦЕНТР').bold = True
            cell_header_2.paragraphs[0].paragraph_format.space_before = Pt(0)
            cell_header_2.paragraphs[0].paragraph_format.space_after = Cm(1)
            cell_header_2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            set_cell_border(cell_header_2, top={"sz": 12, "val": "single", "color": "black", "space": "0"})
        elif i == 1:
            cell_header_2 = title_table.cell(1, 0)
            
        """Утверждающая надпись"""
        title_table.cell(1, 2).merge(title_table.cell(2, 2))
        cell_seo = title_table.cell(1, 2)
        cell_seo.paragraphs[0].add_run("УТВЕРЖДАЮ\n").bold = True
        cell_seo.paragraphs[0].add_run(ratification[i])
        cell_seo.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        """Адрес осуществления деятельности"""
        title_table.cell(2, 0).merge(title_table.cell(2, 1))
        cell_adress = title_table.cell(2, 0)
        cell_adress.paragraphs[0].add_run(adress[i])
        cell_adress.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        cell_adress.paragraphs[0].paragraph_format.line_spacing = Pt(12)
        cell_adress.width = Cm(15)
        
        doc.save(self.path)
        
        
        
        
obj = Protocol(path = 'D:\\My_projects\\Protoсols\\tests_1.docx', test_center=test_center[0])
obj.create_new_file()
obj.create_title_list()