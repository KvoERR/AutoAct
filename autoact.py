from docx import Document
import pandas as pd
import pymorphy3
from datetime import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH

morph = pymorphy3.MorphAnalyzer()  

def extend_date(date):
    months = [
        "января", "февраля", "марта", "апреля",
        "мая", "июня", "июля", "августа",
        "сентября", "октября", "ноября", "декабря"
    ]
    parts = date.split('.')
    return f"« {parts[0]} » {months[int(parts[1])-1]} {parts[2]} г."

def to_genitive_name(full_name):
    parts = full_name.strip().split()
    if len(parts) != 3:
        raise ValueError("Ожидалось 3 части: фамилия имя отчество")

    last_name, first_name, patronymic = parts

    # Определяем пол по отчеству
    patronymic_clean = patronymic.lower()
    if patronymic_clean.endswith(('ович', 'евич', 'ич')):
        gender = 'masc'
    elif patronymic_clean.endswith(('овна', 'евна', 'ична')):
        gender = 'femn'
    else:
        gender = 'masc'

    def decline(word, case):
        parses = morph.parse(word)
        
        # 1. Сначала ищем парс с пометкой Surn (фамилия) и правильным родом
        for p in parses:
            if 'Surn' in p.tag:
                if hasattr(p.tag, 'gender') and p.tag.gender == gender:
                    inflected = p.inflect({case})
                    if inflected:
                        return inflected.word.capitalize()

        # 2. Если не нашли — ищем прилагательное (для -ов/-ев/-ин/-ын, -ова/-ева/-ина)
        for p in parses:
            if 'ADJF' in p.tag or 'Anum' in p.tag:
                if hasattr(p.tag, 'gender') and p.tag.gender == gender:
                    inflected = p.inflect({case})
                    if inflected:
                        return inflected.word.capitalize()

        # 3. Обычные имена/фамилии без метки Surn
        for p in parses:
            if 'Name' in p.tag or 'NOUN' in p.tag:
                if hasattr(p.tag, 'gender') and p.tag.gender == gender:
                    inflected = p.inflect({case})
                    if inflected:
                        return inflected.word.capitalize()

        # 4. Без учёта пола
        for p in parses:
            inflected = p.inflect({case})
            if inflected:
                return inflected.word.capitalize()

        return word.capitalize()

    last_g = decline(last_name, 'gent')
    first_g = decline(first_name, 'gent')
    patronymic_g = decline(patronymic, 'gent')

    return f"{last_g} {first_g} {patronymic_g}"

def to_genitive_dep(dep):
    words = dep.strip().split()
    if not words:
        return dep

    first_word = words[0]
    rest_words = words[1:]

    parses = morph.parse(first_word)
    for p in parses:
        inflected = p.inflect({'gent'})
        if inflected:
            new_first = inflected.word.capitalize()
            break
    else:
        new_first = first_word
    
    new_first=new_first.lower()

    return ' '.join([new_first] + rest_words)

def to_inst_name(name):
    parts = name.strip().split()
    if len(parts) != 3:
        raise ValueError("Ожидалось 3 части: фамилия имя отчество")

    last_name, first_name, patronymic = parts

    patronymic_clean = patronymic.lower()
    if patronymic_clean.endswith(('ович', 'евич', 'ич')):
        gender = 'masc'
    elif patronymic_clean.endswith(('овна', 'евна', 'ична')):
        gender = 'femn'
    else:
        gender = 'femn'

    def decline(word, case):
        parses = morph.parse(word)
        for p in parses:
            if hasattr(p.tag, 'gender') and p.tag.gender == gender:
                inflected = p.inflect({case})
                if inflected:
                    return inflected.word.capitalize()
        for p in parses:
            inflected = p.inflect({case})
            if inflected:
                return inflected.word.capitalize()
        return word.capitalize()

    last_d = decline(last_name, 'ablt')
    first_d = decline(first_name, 'ablt')
    patronymic_d = decline(patronymic, 'ablt')

    return f"{last_d} {first_d} {patronymic_d}"

def shorten_name(full_name):
    parts = full_name.strip().split()
    if len(parts) != 3:
        raise ValueError("Ожидалось 3 части: Фамилия Имя Отчество")
    
    last_name = parts[0].strip()
    first_name_initial = parts[1][0].upper() + '.'
    patronymic_initial = parts[2][0].upper() + '.'

    return f"{last_name} {first_name_initial}{patronymic_initial}"

class AutoAct:
    def __init__(self,file_path,options):
        self.table = pd.read_excel(file_path)
        self.options = options
        self.catalog = "samples"
        
    def parce_cells(self, row):
        self.act_number = str(row['Номер акта'])
        self.name=row['Имя']
        self.gent_name=to_genitive_name(self.name)
        self.inst_name=to_inst_name(self.name)
        self.short_name=shorten_name(self.name)
        self.snils = row
        self.job_title = row['Должность']
        self.department = row['Отдел']
        self.gent_department = to_genitive_dep(self.department)
        self.control = row['Управление']
        self.date_obj = row['Дата']
        self.date = row['Дата'].strftime('%d.%m.%Y')
        self.extended_date=extend_date(self.date)
        
        self.room=row['Помещение']
        self.adress = row['Адрес']
        self.comp_name = row['Наименование ЭВМ']
        self.serial_number = row['S/N ЭВМ']
        self.ip = row['IP-адрес']
        self.disk_number = row['Номер диска']
        self.os=row['ОС']

        self.id_vipnet = row['ID VipNet']
        self.crypto_number = str(row['Номер криптосредства'])
        self.crypto_date = row['Дата криптосредства']
        self.sticker_number = str(row['Номер стикера'])

        self.boss=row['Начальник']
    
    def replace_in_doc(self, doc):
        for paragraph in doc.paragraphs:
            if '{' in paragraph.text:  # Только если есть потенциальные плейсхолдеры
                for key, value in self.__dict__.items():
                    if f"{{{key}}}" in paragraph.text:
                        paragraph.text = paragraph.text.replace(f"{{{key}}}", str(value))

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    self.replace_in_doc(cell)

    def generate_doc(self,name1,name2):
        doc = Document(f'{self.catalog}/{name1}.docx')
        self.replace_in_doc(doc)
        doc.save(f"{self.name} {name2}.docx")
    
    def table_processing(self):
        for index, row in self.table.iterrows():
            self.parce_cells(row)
            if self.options["act"]:
                self.generate_doc('act','CSP акт')
            if self.options["bid"]:
                self.generate_doc('bid','CSP заяка')
            if self.options["csp"]:
                self.generate_doc('csp','CSP уничтожение')
            yield (f"Документы на {self.name} готовы")

if __name__ == '__main__':
    autoact=AutoAct()
    autoact.table_processing()