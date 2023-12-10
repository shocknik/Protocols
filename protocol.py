import os
from setting import test_center, ratification, title_test_center, adress
from borders import set_cell_border
from numpage import add_page_number
from docx import Document
from docx.oxml import OxmlElement, ns
from docx.shared import Inches, Cm, Mm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
from datetime import datetime
from backend_read import get_cable_mark,\
    get_specifications,\
    get_name_specifications,\
    get_list_text_par,\
    create_marker_list,\
    change_font, border_form,\
    func_union_cells,\
    get_list_par_from_tables,\
    filling_table_heads_all,\
    get_list_requarements,\
    get_list_methods,\
    table_inner_border_vertical,\
    func_def_test_by_program
class Protocol:
    
    def __init__(self, path, test_center) -> None:
        self.path = path
        self.test_center = test_center
        
    def create_new_file(self):
        """Метод создания нового файла"""
        doc = Document()
        doc.save(self.path)
        
    def create_title_list(self):
        """Метод создания титульного листа"""
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
        chp = cell_header.paragraphs[0]
        chp.add_run(title_test_center[i]).font.size = Pt(14)
        chp.runs[0].bold = True
        chp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell_header.width = Cm(23)
        chp.paragraph_format.space_after = Pt(0)
        set_cell_border(cell_header, bottom={"sz": 12, "val": "single", "color": "black", "space": "0"})
        if i == 0:
            cell_header_2 = title_table.cell(1, 1)
            chp2 = cell_header_2.paragraphs[0]
            chp2.add_run('ИСПЫТАТЕЛЬНЫЙ ЦЕНТР').bold = True
            chp2.paragraph_format.space_before = Pt(0)
            chp2.paragraph_format.space_after = Cm(1)
            chp2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        elif i == 1:
            cell_header_2 = title_table.cell(1, 0)
            cell_header_2.paragraphs[0].add_run().add_picture('logoSK.png', width=Inches(1.25))
        """Утверждающая надпись"""
        title_table.cell(1, 2).merge(title_table.cell(2, 2))
        cell_seo = title_table.cell(1, 2)
        cs = cell_seo.paragraphs[0]
        cs.add_run("УТВЕРЖДАЮ\n").bold = True
        cs.add_run(f'{ratification[i]}\n')
        cs.add_run(f'«___» _____________ {datetime.now().year}')
        cs.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        """Адрес осуществления деятельности"""
        title_table.cell(2, 0).merge(title_table.cell(2, 1))
        cell_adress = title_table.cell(2, 0)
        ca = cell_adress.paragraphs[0]
        ca.add_run(adress[i])
        ca.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        ca.paragraph_format.line_spacing = Pt(12)
        cell_adress.width = Cm(15)
        """Номер протокола"""
        cell_prot = title_table.cell(3, 1)
        cpp = cell_prot.paragraphs[0]
        cpp.add_run("ПРОТОКОЛ №            \n").bold = True
        cpp.add_run("от              \nприемочных испытаний")
        cpp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cpp.paragraph_format.space_after = Pt(0)
        """Описание протокола испытаний"""
        cell_discription = title_table.cell(4, 1)
        cell_discription.width = Cm(30)
        cd = cell_discription.paragraphs[0]
        cd.add_run("кабеля                                      марки ")
        cd.add_run("{},\n".format(get_cable_mark()[0])).bold = True
        cd.add_run("изготовленного ООО НПП «СПЕЦКАБЕЛЬ» на соответствие требованиям {} {}".\
            format(get_specifications(), get_name_specifications()))
        cd.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY_HI
        cd.paragraph_format.line_spacing = Pt(12)
        cd.paragraph_format.space_after = Pt(6)
        """Доп информация о протоколе"""
        title_table.cell(5, 0).merge(title_table.cell(5, 1))
        title_table.cell(5, 1).merge(title_table.cell(5, 2))
        cell_remarks = title_table.cell(5, 0)
        crp = cell_remarks.paragraphs[0]
        crp.add_run(f'1 Листов всего: ')
        add_page_number(crp)
        crp.add_run(f'\n2 Результаты испытаний распространяются только на предоставленный (е) заказчиком образец (цы).\n\
3 Протокол испытаний не может быть частично или полностью воспроизведен без письменного разрешения Испытательного центра."')
        crp.paragraph_format.space_before = Cm(2)
        """Надпись Москва 2023"""
        cell_moscow = title_table.cell(6, 1)
        cym = cell_moscow.paragraphs[0]
        cym.add_run(f"Москва\n{datetime.now().year}")
        cym.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cym.paragraph_format.space_before = Cm(2)
        doc.save(self.path)
    
    def create_two_list(self):
        doc = Document(self.path)
        doc.add_section(start_type=WD_SECTION_START.NEW_PAGE)
        doc.add_page_break()
        footer_1 = doc.sections[0].footer.paragraphs[0]
        footer_1.style.font.size = Pt(10)
        footer_1.add_run('Нижний колонтитул для первой секции')
        footer_2 = doc.sections[1].footer.paragraphs[0]
        footer_2.style.font.size = Pt(10)
        footer_2.add_run('Нижний колонтитул для второй секции')
        doc.save(self.path)
    
        
        
   
        
obj = Protocol(path = 'D:\\My_projects\\Protoсols\\tests_1.docx', test_center=test_center[1])
obj.create_new_file()
obj.create_title_list()
obj.create_two_list()

os.startfile("D:\\My_projects\\Protoсols\\tests_1.docx")



