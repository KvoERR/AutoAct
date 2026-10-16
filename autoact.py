import os
from docx import Document
from abc import ABC, abstractmethod
import pymorphy3


morph = pymorphy3.MorphAnalyzer()  

class Functions:
    def extend_date(date):
        months = [
            "января", "февраля", "марта", "апреля",
            "мая", "июня", "июля", "августа",
            "сентября", "октября", "ноября", "декабря"
        ]
        
        # Проверяем, что это Timestamp или datetime
        if hasattr(date, 'year') and hasattr(date, 'month') and hasattr(date, 'day'):
            day = date.day
            month = date.month
            year = date.year
        else:
            # На всякий случай — fallback
            date_str = str(date).strip()
            if '-' in date_str:
                year, month, day = map(int, date_str.split()[0].split('-'))
            else:
                raise ValueError(f"Неподдерживаемый формат даты: {date}")

        return f"« {day} » {months[month - 1]} {year} г."

    def shorten_date(date):
        return date.strftime('%d.%m.%Y')

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

    def to_genitive_word(word):
        words = word.strip().split()
        if not words:
            return word

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

class AutoAct(ABC):
    def __init__(self, filename: str, data: str, templates: str, changes: str, features: str) -> dict:
        self.filename = filename
        self.data = data
        self.template = templates
        self.changes = changes
        self.features = features

        print(self.__dict__)

    @abstractmethod
    def run(self):
        ...
    
        """Выполнить обработку и вернуть результат."""
    '''def __init__(self, data_path, template_path, changes, features):
        self.data_dir = data_path
        self.template_dir = template_path
        self.changes = changes
        self.features = features
    
        self.apply_features()
    
    def extend_changes(self,name,key,base):
        "Добавляет измененные features в changes"
        if name == "Слово в Р.п.":
            for change in self.changes:
                change[key]=to_genitive_word(change[base])
        elif name == "Имя в Р.п.":
            for change in self.changes:
                change[key]=to_genitive_name(change[base])
        elif name == "Имя в Д.п.":
            for change in self.changes:
                change[key]=to_inst_name(change[base])
        elif name == "Сокращенное имя":
            for change in self.changes:
                change[key]=shorten_name(change[base])
        elif name == "Сокращённая дата":
            for change in self.changes:
                change[key]=shorten_date(change[base]) 
        elif name == "Расширенная дата":
            for change in self.changes:
                change[key]=extend_date(change[base]) 


    def apply_features(self):
        "Проходимся по features и расширяем changes"
        for feature in self.features:
            name = feature.get('name')
            rules = feature.get('data', {})
            for base, key in rules.items():
                self.extend_changes(name,key,base)
                

    def replace_in_doc(self, doc, change):
        def do_replace(paragraph):
            for key, value in change.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(value))

        for paragraph in doc.paragraphs:
            do_replace(paragraph)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        do_replace(paragraph)
    
    def table_processing(self):
        iter = 1
        for filename in os.listdir(self.template_dir):
            if filename.endswith('.docx') and not filename.startswith('~$'):
                filepath = os.path.join(self.template_dir, filename)
                
                try:
                    for change in self.changes:
                        doc = Document(filepath)
                        self.replace_in_doc(doc, change)

                        # Генерация имени файла
                        name_value = change.get('{name}', f"объект_{iter}")
                        base_name = os.path.splitext(filename)[0]
                        safe_name = "".join(c for c in name_value if c.isalnum() or c in " _-")[:50]
                        output_filename = f"{base_name} {safe_name}.docx"

                        doc.save(output_filename)
                        yield f"✅ {output_filename}"
                        iter += 1

                except Exception as e:
                    yield f"❌ Ошибка при обработке {filename}: {str(e)}"
'''

class ManualMode(AutoAct):
    def run(self):
        iter = 1
        for template in os.listdir(self.templates):
            if not template.endswith('.docx') or template.startswith('~$'):
                continue

            filepath = os.path.join(self.templates, template)
            base_name = os.path.splitext(template)[0]

            for change in self.changes:
                try:
                    doc = Document(filepath)
                    self.replace_in_doc(doc, change)

                    name_value = change.get(self.filename, '').strip() or f"объект_{iter}"
                    safe_name = "".join(
                        c for c in name_value if c.isalnum() or c in " _-"
                    )[:50].strip() or f"объект_{iter}"

                    output_filename = f"{base_name} {safe_name}.docx"
                    doc.save(output_filename)

                    yield f"✅ {output_filename}"
                    iter += 1

                except Exception as e:
                    yield f"❌ Ошибка ({template}, строка {iter}): {e}"
                    iter += 1

class AutoMode(AutoAct):
    def run(self):
        ...

if __name__ == '__main__':
    autoact=AutoAct()
    #autoact.table_processing()