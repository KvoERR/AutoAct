from docx import Document
import pandas as pd
import pymorphy3

morph = pymorphy3.MorphAnalyzer()

def extend_date(date):
    months = [
        "января", "февраля", "марта", "апреля",
        "мая", "июня", "июля", "августа",
        "сентября", "октября", "ноября", "декабря"
    ]
    parts = date.split('-')
    return f"« {parts[2].split(' ')[0]} » {months[int(parts[1])-1]} {parts[0]} г."

def to_genitive_name(full_name):
    parts = full_name.strip().split()
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

    def decline(word, required_case):
        parses = morph.parse(word)
        for p in parses:
            if hasattr(p.tag, 'gender') and p.tag.gender == gender:
                inflected = p.inflect({required_case})
                if inflected:
                    return inflected.word.capitalize()
        for p in parses:
            inflected = p.inflect({required_case})
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

table = pd.read_excel('Инфо.xlsx')
for index, row in table.iterrows():
    act_number = str(row['Номер акта'])
    name=row['Имя']
    gent_name=to_genitive_name(name)
    inst_name=to_inst_name(name)
    short_name=shorten_name(name)
    job_title = row['Должность']
    department = row['Отдел']
    gent_department = to_genitive_dep(department)
    control = row['Управление']
    date=str(row['Дата'])
    extended_date=extend_date(date)
    room=row['Помещение']
    adress = row['Адрес']

    comp_name = row['Наименование ЭВМ']
    serial_number = row['S/N ЭВМ']
    ip = row['IP-адрес']
    disk_number = row['Номер диска']
    os=row['ОС']

    id_vipnet = row['ID VipNet']
    crypto_number = str(row['Номер криптосредства'])
    crypto_date = row['Дата криптосредства']
    sticker_number = str(row['Номер стикера'])

    boss=row['Начальник']

    doc = Document('Акт.docx')

    paragraph = doc.paragraphs[0]
    info = paragraph.text
    info=info.replace(f"{{act_number}}",act_number)
    paragraph.text = info

    cell = doc.tables[1].rows[0].cells[1]
    info=cell.text
    info=info.replace(f"{{extended_date}}",extended_date)
    cell.text = info

    paragraph = doc.paragraphs[4]
    info = paragraph.text
    info=info.replace(f"{{name}}",name
        ).replace(f"{{job_title}}",job_title
        ).replace(f"{{department}}",department
        ).replace(f"{{date}}",date
        ).replace(f"{{room}}",room 
        ).replace(f"{{adress}}",adress
        ).replace(f"{{serial_number}}",serial_number
        ).replace(f"{{disk_number}}",disk_number
        ).replace(f"{{gent_name}}",gent_name
        )                    
    paragraph.text=info

    paragraph = doc.paragraphs[6]
    info = paragraph.text
    paragraph.runs[7].text=id_vipnet

    paragraph = doc.paragraphs[11]
    info = paragraph.text
    info=info.replace(f"{{gent_name}}",gent_name)
    paragraph.text=info

    paragraph = doc.paragraphs[12]
    info = paragraph.text
    info=info.replace(f"{{room}}",room
    ).replace(f"{{crypto_number}}",crypto_number
    ).replace(f"{{date}}",date)
    paragraph.text=info
    
    paragraph = doc.paragraphs[14]
    info = paragraph.text
    paragraph.runs[5].text=sticker_number
    paragraph.runs[7].text=inst_name

    table=doc.tables[2]
    info=table.rows[0].cells[0].paragraphs[5].text
    info=info.replace(f"{{extended_date}}",extended_date)
    table.rows[0].cells[0].paragraphs[5].text=info

    table=doc.tables[2]
    info=table.rows[0].cells[1].paragraphs[0].text
    info=info.replace(f"{{gent_department}}",gent_department)
    table.rows[0].cells[1].paragraphs[0].text=info

    table=doc.tables[2]
    info=table.rows[0].cells[1].paragraphs[3].text
    info=info.replace(f"{{short_name}}",short_name)
    table.rows[0].cells[1].paragraphs[3].text=info

    table=doc.tables[2]
    info=table.rows[0].cells[1].paragraphs[4].text
    info=info.replace(f"{{extended_date}}",extended_date)
    table.rows[0].cells[1].paragraphs[4].text=info

    doc.save(f"Акт_{name}.docx")

    doc = Document('Заявка.docx')

    table=doc.tables[0]

    info=table.rows[0].cells[2].text
    info=info.replace(f"{{name}}",name)
    table.rows[0].cells[2].text=info

    info=table.rows[1].cells[2].text
    jt1=job_title.capitalize()
    info=info.replace(f"{{job_title}}",jt1)
    table.rows[1].cells[2].text=info

    info=table.rows[2].cells[2].text
    info=info.replace(f"{{department}}",department)
    table.rows[2].cells[2].text=info

    info=table.rows[3].cells[2].text
    info=info.replace(f"{{control}}",control)
    table.rows[3].cells[2].text=info

    info=table.rows[4].cells[2].text
    info=info.replace(f"{{adress}}",adress
    ).replace(f"{{room}}",room)
    table.rows[4].cells[2].text=info

    info=table.rows[5].cells[2].text
    info=info.replace(f"{{comp_name}}",comp_name)
    table.rows[5].cells[2].text=info

    info=table.rows[6].cells[2].text
    info=info.replace(f"{{serial_number}}",serial_number)
    table.rows[6].cells[2].text=info

    info=table.rows[7].cells[2].text
    info=info.replace(f"{{ip}}",ip)
    table.rows[7].cells[2].text=info

    info=table.rows[8].cells[2].text
    info=info.replace(f"{{disk_number}}",disk_number)
    table.rows[8].cells[2].text=info

    info=table.rows[9].cells[2].text
    info=info.replace(f"{{os}}",os)
    table.rows[9].cells[2].text=info

    info=table.rows[10].cells[2].text
    info=info.replace(f"{{sticker_number}}",sticker_number)
    table.rows[10].cells[2].text=info

    table=doc.tables[1]

    info=table.rows[0].cells[2].paragraphs[0].text
    info=info.replace(f"{{boss}}",boss)
    table.rows[0].cells[2].paragraphs[0].text=info

    doc.save(f"Заявка_{name}.docx")
    print(f"Документы на {name} готовы")
    
    


'''
number = input("серийный номер: ")
name = input("имя: ")

doc = Document('Акт.docx')

paragraph = doc.paragraphs[0]
info = paragraph.text

info = insert_string(info, number + " ", info.find("номер") + 6)
info = insert_string(info, " " + name, info.find("пользователя") + 12)

paragraph.text = info

doc.save(f"Акт_{name}.docx")

print("Акт составлен")

doc = Document('Заявка.docx')

table = doc.add_table(rows=2, cols=2)

table.cell(0,0).text = "Серийный номер АРМ"
table.cell(0,1).text = number
table.cell(1,0).text = "ФИО Пользователя"
table.cell(1,1).text = name

doc.save(f"Заявка_{name}.docx")

print("Заявка составлена")
'''