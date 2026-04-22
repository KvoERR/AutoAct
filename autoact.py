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
    def __init__(self):
        self.table = pd.read_excel('Инфо.xlsx')
        
    def parce_cells(self, row):
        self.act_number = str(row['Номер акта'])
        self.name=row['Имя']
        self.gent_name=to_genitive_name(self.name)
        self.inst_name=to_inst_name(self.name)
        self.short_name=shorten_name(self.name)
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
    def generate_act(self):
        doc = Document('Акт.docx')

        paragraph = doc.paragraphs[0]
        info = paragraph.text
        info=info.replace(f"{{act_number}}",self.act_number)
        paragraph.text = info

        cell = doc.tables[1].rows[0].cells[1]
        info=cell.text
        info=info.replace(f"{{extended_date}}",self.extended_date)
        cell.text = info
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        paragraph = doc.paragraphs[4]
        info = paragraph.text
        info=info.replace(f"{{name}}",self.name
            ).replace(f"{{job_title}}",self.job_title
            ).replace(f"{{department}}",self.gent_department
            ).replace(f"{{date}}",self.date
            ).replace(f"{{room}}",self.room 
            ).replace(f"{{adress}}",self.adress
            ).replace(f"{{serial_number}}",self.serial_number
            ).replace(f"{{disk_number}}",self.disk_number
            ).replace(f"{{gent_name}}",self.gent_name
            )                    
        paragraph.text=info

        paragraph = doc.paragraphs[6]
        info = paragraph.text
        paragraph.runs[7].text=self.id_vipnet

        paragraph = doc.paragraphs[11]
        info = paragraph.text
        info=info.replace(f"{{gent_name}}",self.gent_name)
        paragraph.text=info

        paragraph = doc.paragraphs[12]
        info = paragraph.text
        info=info.replace(f"{{room}}",self.room
        ).replace(f"{{crypto_number}}",self.crypto_number
        ).replace(f"{{date}}",self.date)
        paragraph.text=info
            
        paragraph = doc.paragraphs[14]
        info = paragraph.text
        paragraph.runs[5].text=self.sticker_number
        paragraph.runs[7].text=self.inst_name

        table=doc.tables[2]
        info=table.rows[0].cells[0].paragraphs[5].text
        info=info.replace(f"{{extended_date}}",self.extended_date)
        table.rows[0].cells[0].paragraphs[5].text=info

        table=doc.tables[2]
        info=table.rows[0].cells[1].paragraphs[0].text
        info=info.replace(f"{{gent_department}}",self.gent_department)
        table.rows[0].cells[1].paragraphs[0].text=info

        table=doc.tables[2]
        info=table.rows[0].cells[1].paragraphs[3].text
        info=info.replace(f"{{short_name}}",self.short_name)
        table.rows[0].cells[1].paragraphs[3].text=info

        table=doc.tables[2]
        info=table.rows[0].cells[1].paragraphs[4].text
        info=info.replace(f"{{extended_date}}",self.extended_date)
        table.rows[0].cells[1].paragraphs[4].text=info

        doc.save(f"Акт_{self.name}.docx")
    def generate_bid(self):
        doc = Document('Заявка.docx')

        table=doc.tables[0]

        info=table.rows[0].cells[2].text
        info=info.replace(f"{{name}}",self.name)
        table.rows[0].cells[2].text=info

        info=table.rows[1].cells[2].text
        jt1=self.job_title.capitalize()
        info=info.replace(f"{{job_title}}",jt1)
        table.rows[1].cells[2].text=info

        info=table.rows[2].cells[2].text
        info=info.replace(f"{{department}}",self.department)
        table.rows[2].cells[2].text=info

        info=table.rows[3].cells[2].text
        info=info.replace(f"{{control}}",self.control)
        table.rows[3].cells[2].text=info

        info=table.rows[4].cells[2].text
        info=info.replace(f"{{adress}}",self.adress
        ).replace(f"{{room}}",self.room)
        table.rows[4].cells[2].text=info

        info=table.rows[5].cells[2].text
        info=info.replace(f"{{comp_name}}",self.comp_name)
        table.rows[5].cells[2].text=info

        info=table.rows[6].cells[2].text
        info=info.replace(f"{{serial_number}}",self.serial_number)
        table.rows[6].cells[2].text=info

        info=table.rows[7].cells[2].text
        info=info.replace(f"{{ip}}",self.ip)
        table.rows[7].cells[2].text=info

        info=table.rows[8].cells[2].text
        info=info.replace(f"{{disk_number}}",self.disk_number)
        table.rows[8].cells[2].text=info

        info=table.rows[9].cells[2].text
        info=info.replace(f"{{os}}",self.os)
        table.rows[9].cells[2].text=info

        info=table.rows[10].cells[2].text
        info=info.replace(f"{{sticker_number}}",self.sticker_number)
        table.rows[10].cells[2].text=info

        table=doc.tables[1]

        info=table.rows[0].cells[2].paragraphs[0].text
        info=info.replace(f"{{boss}}",self.boss)
        table.rows[0].cells[2].paragraphs[0].text=info

        doc.save(f"Заявка_{self.name}.docx")
    def generate_csp(self):
        doc = Document('CSP уничтожение.docx')
        paragraph = doc.paragraphs[0]
        info = paragraph.text
        info=info.replace(f"{{act_number}}",self.act_number)
        paragraph.text = info
    def table_processing(self):
        for index, row in self.table.iterrows():
            self.parce_cells(row)
            self.generate_act()
            self.generate_bid()
            #self.generate_csp()

            



            print(f"Документы на {self.name} готовы")

if __name__ == '__main__':
    autoact=AutoAct()
    autoact.table_processing()