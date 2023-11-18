import logging_config
import logging
from peewee import *

logger = logging.getLogger(__name__)
logger.info("An info")


db = SqliteDatabase('D:\\My_projects\\Protoсols\\tests.db')

class BaseModel(Model):    
    class Meta:
        database = db


class TestCategory(BaseModel):
    ID = PrimaryKeyField(null=False)
    category = CharField(unique=True)
    
        
class PMITests(BaseModel):
    ID = PrimaryKeyField(null=False)
    name = CharField(unique=True)
    form = CharField(unique=True)
    id_category=ForeignKeyField(TestCategory)
    
    
class TestCriteria(BaseModel):
    ID = PrimaryKeyField(null=False)
    criteria = CharField()
    mean_criteria = CharField()

class RelationshipCategoryCriteria(BaseModel):
    id_category = ForeignKeyField(TestCriteria)
    id_criteria = ForeignKeyField(TestCategory)
    
    
def create_tables():
    try:
        with db:
            db.create_tables([TestCategory, PMITests, TestCriteria, RelationshipCategoryCriteria])
    except Exception as e:
        logger.error(f'Ошибка в создании таблицы: {e}', exc_info=True)
create_tables()