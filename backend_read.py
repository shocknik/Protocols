"""
TODO
0 Попытаться научиться распаковывать ДОК-файлы и парсить xml
import xml.etree.ElementTree as ET
import lxml
import zipfile
from bs4 import BeautifulSoup

doc_zip = zipfile.ZipFile("Программы квалификационных испытаний\\Программы квалификационных испытаний\\020\\Программа типовых испытаний ТУ 16.К99-020-2009.docx")
doc_xml = doc_zip.read('word/document.xml')

soup_xml = BeautifulSoup(doc_xml, features='xml')
pretty_xml = soup_xml.prettify()
root = ET.fromstring(pretty_xml)
print(root)
1 Создать шаблон протокола испытаний
    - автоматизировать генерацию шаблона с присвоением номер +1
    - делать запрос на чтение программы
    - создавать таблицу приборов
    - создавать таблицу испытаний
2 Изучить программы испытаний
    - выявить ключевые слова, к которым могут относиться:
        - марка кабеля
        - номер ТУ
        - наименование ТУ
    - научится переходить к таблице и читать ее
    - читать испытания в таблице и собирать их в словарь или лист
    - заполнять по полученому словарю таблицу в протоколе
3 Структура проекта:
    3.1 main.py (главный исполняемый файл, который генерит запросы)
    3.2 backend.py (файл содержащий функции)
    3.3 model.py (файл который реализует логику, используя функции backend.py)
    3.4 view.py (отображает информацию через терминал)
    3.5 controller.py (осуществляет взаимодействие между model.py и veiw.py)
"""
import re
import os
import docx
from docx import Document
from docx.shared import Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING

document = Document()

def get_paths_filesdoc() -> list:
    """Получить пути всех файлов doc в директории"""
    paths = []
    folder = os.getcwd()
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('docx') and not file.startswith('~'):
                paths.append(os.path.join(root, file))
    return paths


doc = Document('Программы квалификационных испытаний\\Программы квалификационных испытаний\\064\\Программа типовых испытаний ФЖТК.357400.064ТУ.docx')

list_par: list = doc.paragraphs # список объектов параграфов


def get_list_text_par() -> list:
    """получаем список тесктов параграфов"""
    list_text_par: list = []
    for par in list_par:
        list_text_par.append(par.text)
    return list_text_par

        
def get_cable_mark() -> list:
    """Получение списка марок кабелей из программы"""
    for par in get_list_text_par():
        result = re.findall(r'[0-9А-Яа-я\S]+\b\s\dх\dх\d[,]*\d*', par) # марка кабеля
        if len(result) > 0:
            return result # Печать списка марок
       
        
def get_specifications() -> str:
    """Получение ТУ на кабели из программы"""
    for par in get_list_text_par():
        result = re.findall(r'[ФЖТК]+\W[0-9]+\W[0-9ТУ]+|[ТУ]+[\s]\d{1,3}\S[К1-9\S]+', par)
        if 0 < len(result) < 2:
            return result[0]


def get_name_specifications() -> str:
    for par in get_list_text_par():
        result = re.findall(r'\«Кабели\s*[а-яА-Я\s,0-9]+\.\s[Техническиеусловия\s]+\»', par) # название ТУ
        if 0 < len(result) < 2:
            return result[0]


def get_list_par_from_tables() -> str:
    """
    Получает список параметров из таблицы в программе
    """
    list_par = []
    for table in doc.tables:
        for i in range(2, len(table.rows)):
            list_par.append(table.cell(i, 1).text)
    return list_par
            
def get_list_requarements() -> str:
    """Полчает список пунктов требований"""
    list_req = []
    for table in doc.tables:
        for i in range(2, len(table.rows)):
            list_req.append(table.cell(i, 2).text)
    return list_req
                

def get_list_methods() -> str:
    """Полчает список пунктов методов"""
    list_methods = []
    for table in doc.tables:
        for i in range(2, len(table.rows)):
            list_methods.append(table.cell(i, 3).text)
    return list_methods
            
