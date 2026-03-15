from docx import Document

def insert_string(original, new_text, index):
    return original[:index] + new_text + original[index:]

number = input("Введите серийный номер: ")
name = input("Введите имя: ")

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
