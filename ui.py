import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
from autoact import AutoAct

class App:
    def __init__(self, root):
        self.root = root
        self.file_path = None
        self.setup_ui()

    def setup_ui(self):
        self.root.title("AutoAct")
        self.root.geometry("1050x600")

        # Прямо используем root как контейнер
        self.setup_main(self.root)  # Передаём root вместо tab
        
    class Feature:
        def __init__(self, tables_frame, name, number, col1, col2):
            self.name = name
            self.auto_add_row = True  # Добавляем этот атрибут

            # Рамка для таблицы
            frame = tk.LabelFrame(tables_frame, text=name, padx=5, pady=5)
            frame.grid(row=0, column=number, sticky='nsew', padx=(0, 5), pady=0)

            # Контейнер с прокруткой
            tree_container = tk.Frame(frame)
            tree_container.pack(fill='both', expand=True)

            # Стиль
            style = ttk.Style()
            style.configure("Custom.Treeview", 
                            background="#f0f0f0",    # Серый фон вне строк
                            fieldbackground="#f0f0f0", 
                            foreground="black",
                            rowheight=22)

            # Создаём Treeview
            self.tree = ttk.Treeview(tree_container, show='headings', height=8, style="Custom.Treeview")
            self.tree.pack(side='left', fill='both', expand=True)

            # Прокрутка
            scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
            scroll.pack(side='right', fill='y')
            self.tree.configure(yscrollcommand=scroll.set)

            # Настройка колонок
            self.tree["columns"] = ("col1", "col2")
            self.tree.heading("col1", text=col1)
            self.tree.heading("col2", text=col2)
            self.tree.column("col1", width=30, anchor='w')  # Увеличил ширину
            self.tree.column("col2", width=50, anchor='w')

            # Привязываем двойной клик для редактирования
            self.tree.bind('<Double-1>', self.on_double_click)

            # Словарь для хранения редактируемых виджетов
            self.edit_widgets = {}

        def insert(self, col1="", col2=""):
            """Вставка строки в таблицу"""
            self.tree.insert("", "end", values=(col1, col2))       
        
        def on_double_click(self, event):
            """Обработка двойного клика для редактирования ячейки"""
            # Получаем элемент, на который кликнули
            region = self.tree.identify_region(event.x, event.y)
            
            if region == "cell":
                # Получаем строку и колонку
                row_id = self.tree.identify_row(event.y)
                column = self.tree.identify_column(event.x)
                
                if not row_id:
                    return
                
                # Получаем текущее значение
                current_value = self.tree.item(row_id, 'values')
                col_index = int(column[1:]) - 1  # column имеет формат "#1", "#2"
                
                if col_index < len(current_value):
                    # Создаем поле для ввода
                    x, y, width, height = self.tree.bbox(row_id, column)
                    
                    # Создаем Entry для редактирования
                    entry = tk.Entry(self.tree, font=("Arial", 10))
                    entry.place(x=x, y=y, width=width, height=height)
                    entry.insert(0, current_value[col_index])
                    entry.focus()
                    
                    # Сохраняем ID строки и индекс колонки в entry
                    entry.row_id = row_id
                    entry.col_index = col_index
                    
                    # Сохраняем контекст для обработки завершения редактирования
                    entry.bind('<Return>', self.save_edit)
                    entry.bind('<FocusOut>', self.save_edit)
                    entry.bind('<Escape>', lambda e, ent=entry: ent.destroy())
                    
                    # Сохраняем ссылку на entry для возможного использования
                    self.edit_widgets[row_id] = entry
        
        def save_edit(self, event):
            """Сохранение отредактированного значения"""
            entry = event.widget
            new_value = entry.get()
            row_id = entry.row_id
            col_index = entry.col_index
            
            # Получаем текущие значения строки
            current_values = list(self.tree.item(row_id, 'values'))
            
            # Обновляем нужную колонку
            if col_index < len(current_values):
                current_values[col_index] = new_value
                self.tree.item(row_id, values=current_values)
            
            # Удаляем поле для ввода
            entry.destroy()
            if row_id in self.edit_widgets:
                del self.edit_widgets[row_id]   

            if self.auto_add_row:
                self.check_and_add_row(row_id) 

        def check_and_add_row(self, edited_row_id):
            """Проверяет, нужно ли добавить новую строку"""
            # Получаем все строки
            all_items = self.tree.get_children()
            
            if not all_items:
                return
            
            # Проверяем последнюю строку
            last_item = all_items[-1]
            last_values = self.tree.item(last_item, 'values')
            
            # Если последняя строка - отредактированная и она не пустая
            if last_item == edited_row_id:
                # Проверяем, заполнена ли хотя бы одна колонка
                is_filled = any(str(v).strip() for v in last_values if v)
                
                if is_filled:
                    # Добавляем новую пустую строку
                    self.add_row()
                    # Прокручиваем к новой строке
                    self.tree.see(last_item)
    
        def get_dict(self):
            result = {}
            for item in self.tree.get_children():
                values = self.tree.item(item, 'values')
                if len(values) >= 2:
                    key = values[0].strip()
                    value = values[1].strip()
                    if key:
                        result[key] = value
            return result
        
    def clear_tree_selection(self, event=None):
        """Снимает выделение со всех таблиц при клике вне их"""
        # Проверяем, кликнули ли не по Treeview
        if event and isinstance(event.widget, ttk.Treeview):
            return  # Если кликнули по таблице, не снимаем выделение
        
        for feature in [self.f1, self.f2, self.f3, self.f4]:
            if hasattr(feature, 'tree'):
                # Убираем выделение со всех элементов
                for item in feature.tree.selection():
                    feature.tree.selection_remove(item)
            
    def setup_main(self, parent):
        # Пути
        path_frame = tk.Frame(parent)
        path_frame.pack(anchor='w', padx=20, pady=5, fill='x')

        tk.Label(path_frame, text="Папка с шаблонами:", font=("Arial", 10)).grid(row=0, column=0, sticky='w')
        self.template_path_entry = tk.Entry(path_frame, width=100, state='readonly', font=("Arial", 10))
        self.template_path_entry.grid(row=0, column=1, padx=5)
        tk.Button(path_frame, text="Обзор", command=self.select_template_folder).grid(row=0, column=2, padx=5)

        tk.Label(path_frame, text="Таблица замен:", font=("Arial", 10)).grid(row=1, column=0, sticky='w')
        self.table_path_entry = tk.Entry(path_frame, width=100, state='readonly', font=("Arial", 10))
        self.table_path_entry.grid(row=1, column=1, padx=5)
        tk.Button(path_frame, text="Обзор", command=self.select_table_file).grid(row=1, column=2, padx=5)

        self.status_label = tk.Label(parent, text="", fg="gray", font=("Arial", 9))
        self.status_label.pack(anchor='w', padx=20, pady=0)

        # Заголовок для таблиц
        tk.Label(parent, text="Дополнительные настройки", font=("Arial", 10)).pack(anchor='w', padx=20, pady=(10, 5))

        # Контейнер для таблиц
        tables_frame = tk.Frame(parent)
        tables_frame.pack(pady=0, padx=20, fill='both', expand=True)

        # Создание таблиц доп. ф-ий
        self.f1 = self.Feature(tables_frame, "Слово в Р.п.", 0, "Слово", "Р.п.")
        self.f1.insert("{department}", "{gent_department}")

        self.f2 = self.Feature(tables_frame, "Имя в Р.п.", 1, "Имя", "Р.п.")
        self.f2.insert("{name}", "{gent_name}")

        self.f3 = self.Feature(tables_frame, "Имя в Д.п.", 2, "Имя", "Д.п.")
        self.f3.insert("{name}", "{inst_name}")

        self.f4 = self.Feature(tables_frame, "Сокращенное имя", 3, "Имя", "Сокращённое")
        self.f4.insert("{name}", "{short_name}")

        self.f5 = self.Feature(tables_frame, "Сокращённая дата", 4, "Дата", "Сокращённая")
        self.f5.insert("{date}", "{short_date}")

        self.f6 = self.Feature(tables_frame, "Расширенная дата", 5, "Дата", "Расширенная")
        self.f6.insert("{date}", "{extended_date}")

        self.f1.insert()
        self.f2.insert()
        self.f3.insert()
        self.f4.insert()
        self.f5.insert()
        self.f6.insert()

        # Растягивание колонок
        for i in range(6):
            tables_frame.columnconfigure(i, weight=1)

        # Снятие выделения
        self.root.bind('<Button-1>', self.clear_tree_selection)

        # Кнопка запуска
        tk.Button(parent, text="Начать", command=self.process_file).pack(pady=10, anchor='w', padx=20)

        # Логи
        log_container = tk.Frame(parent)
        log_container.pack(pady=10, anchor='w', padx=20, fill='both', expand=True)

        self.log_text = tk.Text(
            log_container,
            height=10,
            width=70,
            state='disabled',
            bg="white",
            font=("Arial", 10)
        )
        self.log_text.pack(fill='both', expand=True)

    def select_table_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл таблицы",
            filetypes=[
                ("Excel файлы", "*.xlsx"),
                ("CSV файлы", "*.csv"),
                ("Старые Excel файлы", "*.xls"),
            ],
        )
        if file_path:
            self.table_path_entry.config(state='normal')
            self.table_path_entry.delete(0, tk.END)
            self.table_path_entry.insert(0, file_path)
            self.table_path_entry.config(state='readonly')

    def select_template_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с шаблонами")
        if folder:
            self.template_path_entry.config(state='normal')
            self.template_path_entry.delete(0, tk.END)
            self.template_path_entry.insert(0, folder)
            self.template_path_entry.config(state='readonly')

    def process_file(self):
        table_path = self.table_path_entry.get().strip()
        template_path = self.template_path_entry.get().strip()
        
        if not table_path:
            self.status_label.config(text="❗ Сначала выберите файл таблицы!", fg="red")
            return

        if not template_path:
            self.status_label.config(text="❗ Сначала выберите папку с шаблонами!", fg="red")
            return

        changes = pd.read_excel(table_path).to_dict('records')
        features = []
        for feature in [self.f1, self.f2, self.f3, self.f4, self.f5, self.f6]:
            feature_data = feature.get_dict()
            feature_name = feature.name
            features.append({
                'name': feature_name,
                'data': feature_data
            })
        
        auto_act = AutoAct(template_path, changes, features)

        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        
        try:
            for status in auto_act.table_processing():
                self.log_text.insert(tk.END, status + "\n")
                self.log_text.see(tk.END)
                self.log_text.update_idletasks()
            
            self.log_text.insert(tk.END, "🎉 Все документы сгенерированы!\n")
        except Exception as e:
            self.log_text.insert(tk.END, f"❌ Ошибка: {str(e)}\n")
        
        self.log_text.config(state='disabled')


if __name__ == "__main__":
    window = tk.Tk()
    app = App(window)
    window.mainloop()