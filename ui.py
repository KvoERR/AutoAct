import tkinter as tk
from tkinter import filedialog
from autoact import AutoAct


class App:
    def __init__(self, root):
        self.root = root
        self.file_path = None

        self.setup_ui()

    def setup_ui(self):
        self.root.title("AutoAct")
        self.root.geometry("600x400")

        # Переменные для чекбоксов
        self.var1 = tk.IntVar()
        self.var2 = tk.IntVar()
        self.var3 = tk.IntVar()

        # Кнопки "Выбрать всё" / "Снять всё"
        tk.Button(self.root, text="Выбрать всё", command=self.select_all).pack(pady=10)
        tk.Button(self.root, text="Снять всё", command=self.deselect_all).pack(pady=5)

        # Чекбоксы
        tk.Checkbutton(self.root, text="Акт", variable=self.var1).pack(pady=5)
        tk.Checkbutton(self.root, text="Заявка", variable=self.var2).pack(pady=5)
        tk.Checkbutton(self.root, text="CSP", variable=self.var3).pack(pady=5)

        # Кнопка выбора файла
        tk.Button(self.root, text="Выбрать файл", command=self.select_file).pack(pady=20)

        # Метка для отображения пути
        self.label = tk.Label(self.root, text="")
        self.label.pack()

        # Кнопка для запуска обработки
        tk.Button(self.root, text="Начать", command=self.process_file).pack(pady=10)

        # Результат
        self.log_text = tk.Text(self.root, height=10, width=70, state='disabled', bg="white", font=("Arial", 10))
        self.log_text.pack(pady=10)

    def select_all(self):
        self.var1.set(1)
        self.var2.set(1)
        self.var3.set(1)

    def deselect_all(self):
        self.var1.set(0)
        self.var2.set(0)
        self.var3.set(0)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Выберите файл таблицы",
            filetypes=[
                ("Excel файлы", "*.xlsx"),
                ("CSV файлы", "*.csv"),
                ("Старые Excel файлы", "*.xls"),
            ],
        )
        if self.file_path:
            self.label.config(text=f"Выбрано: {self.file_path}")
        else:
            self.label.config(text="Файл не выбран")

    def process_file(self):
        if not self.file_path:
            self.label.config(text="Сначала выберите файл!")
            return

        options = {
            "act": bool(self.var1.get()),
            "bid": bool(self.var2.get()),
            "csp": bool(self.var3.get()),
        }

        auto_act = AutoAct(self.file_path, options)

        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        
        for status in auto_act.table_processing():
            self.log_text.insert(tk.END, status + "\n")
            self.log_text.see(tk.END) 
            self.log_text.update_idletasks() 
        
        self.log_text.insert(tk.END, "🎉 Все документы сгенерированы!\n")
        self.log_text.config(state='disabled')


if __name__ == "__main__":
    window = tk.Tk()
    app = App(window)
    window.mainloop()