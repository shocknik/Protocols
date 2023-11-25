
import logging_config
import peewee
import logging
from connect_to_db import TestCategory, SubCategory, db
from setting import *

logger = logging.getLogger(__name__)

def insert_category(test_category: dict[str, str]):
    """Функция, которая заполняет записи в БД таблицу категорий из словаря"""
    for key, value in test_category.items():
        try:
            category = TestCategory(category=value)
            category.save()
            logger.info(f"Запись {value} OK")
        except peewee.IntegrityError as e:
            print(f'Исключение {e}: Запись "{value}" уже есть в базе')
            
            
           
def insert_subcategory(subcategory: dict[str, list]):
    """Функция, которая заполняет записи в БД таблицу подкатегорий из словаря"""
    for key, value in subcategory.items():
        logger.info(f"Ключ категории {key}")
        cat = TestCategory.get(TestCategory.category == key)
        for sub in value:
            try:
                subcat = SubCategory(name = sub, id_category = cat)
                subcat.save()
            except peewee.IntegrityError as e:
                print(f'Подкатегория "{sub}" уже существует в категории "{cat.category}"')
                continue


def insert_pmi(dict_test: dict[str, str]):
    for k, v in dict_test:
        try:
            test = PMITest(name = k, form = v)
            test.save()
        except peewee.IntegrityError as e:
            print(f'Запись {k}: {v} уже существует')
            continue
    
def insert_pmitests(tests_from_pmi: dict[str, dict[str, dict[str, str]]]):
    """Функция, которая заполняет несколько таблиц из словаря с 
    ПМИ испытаниями и формами из протокола,
    учитывая категории и подкатегории"""
    i = 1
    dict_cat = {}
    dict_sub = {}
    dict_test = {}
    for category, subcategory in tests_from_pmi.items():
        subs = []
        dict_cat[i] = category
        i += 1
        for key, value in subcategory.items():
            subs.append(key)
            for k, v in value.items():
                dict_test[k] = v
        dict_sub[category] = subs
    # insert_category(dict_cat)
    # insert_subcategory(dict_sub)
    return dict_cat, dict_sub, dict_test
print(insert_pmitests(PMITest))   
