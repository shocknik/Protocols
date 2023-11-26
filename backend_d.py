
import logging_config
import peewee
import logging
from connect_to_db import TestCategory, SubCategory, PMITests, db
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

    
def insert_pmitests(tests_from_pmi: dict[str, dict[str, dict[str, str]]]):
    """Функция, которая заполняет несколько таблиц из словаря с 
    ПМИ испытаниями и формами из протокола,
    учитывая категории и подкатегории"""
    for category, subcategory in tests_from_pmi.items():
        cat = TestCategory.get_or_none(TestCategory.category == category)
        if cat is None:
            cat = TestCategory(category=category)
            cat.save()
        for key, value in subcategory.items():
            sub = SubCategory.get_or_none(SubCategory.name == key)
            if sub is None:
                sub = SubCategory(name = key, id_category = cat)
                sub.save()
            for k, v in value.items():
                try:
                    pmitest = PMITests(name=k, form=v, id_category=cat, id_subcategory=sub)
                    pmitest.save()
                except peewee.IntegrityError as e:
                    print(f'Запись {k}: {v} уже существует')
                continue    
                   
insert_pmitests(PMITest)