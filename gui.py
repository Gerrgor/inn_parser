import tkinter as tk
from tkinter import filedialog, messagebox
from parser import Parser
from utils import validate_inn_file, validate_save_file


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Парсер данных по ИНН")
        self.root.geometry("700x500")
        self.root.configure(bg="lightblue")
        self.root.attributes("-alpha", 0.98)
        self.center_window()
        self.current_step = 1
        self.inn_file = ""  # Сохраняем путь к файлу с ИНН
        self.save_file_path = ""  # Сохраняем путь для сохранения результата
        self.source_option = ""  # Сохраняем выбранный источник данных
        self.selected_data = []  # Сохраняем выбранные данные для парсинга
        self.checkbox_vars = {}  # Сохраняем состояние чекбоксов
        self.frame = tk.Frame(self.root, bg="lightblue")
        self.frame.pack(expand=True)
        self.parser = Parser()  # Инициализация парсера
        self.step1()

    def center_window(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (700 // 2)
        y = (screen_height // 2) - (500 // 2)
        self.root.geometry(f"700x500+{x}+{y}")

    def step1(self):
        self.clear_frame()
        self.current_step = 1

        tk.Label(
            self.frame, text="Выберите файл с ИНН:", bg="lightblue", font=("Arial", 16)
        ).pack(pady=20)

        self.entry_inn = tk.Entry(self.frame, width=50, font=("Arial", 14))
        self.entry_inn.pack(pady=50)
        self.entry_inn.insert(0, self.inn_file)  # Вставляем сохраненный путь

        tk.Button(
            self.frame,
            text="Обзор",
            command=self.load_inn_file,
            font=("Arial", 14),
            width=15,
        ).pack(pady=20)

        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=40)

        tk.Button(
            button_frame, text="Далее", command=self.step2, font=("Arial", 14), width=15
        ).pack()

    def load_inn_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.entry_inn.delete(0, tk.END)
            self.entry_inn.insert(0, file_path)
            self.inn_file = file_path  # Сохраняем путь к файлу

    def step2(self):
        if not validate_inn_file(self.inn_file):
            return  # Прерываем выполнение функции, если файл не валиден

        self.clear_frame()
        self.current_step = 2

        tk.Label(
            self.frame,
            text="Выберите путь для сохранения результата:",
            bg="lightblue",
            font=("Arial", 16),
        ).pack(pady=20)

        self.entry_save = tk.Entry(self.frame, width=50, font=("Arial", 14))
        self.entry_save.pack(pady=50)
        self.entry_save.insert(0, self.save_file_path)  # Вставляем сохраненный путь

        tk.Button(
            self.frame,
            text="Обзор",
            command=self.save_file,
            font=("Arial", 14),
            width=15,
        ).pack(pady=20)

        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=40)

        tk.Button(
            button_frame, text="Назад", command=self.step1, font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=40)
        tk.Button(
            button_frame, text="Далее", command=self.step3, font=("Arial", 14), width=15
        ).pack(side=tk.RIGHT, padx=40)

    def save_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.entry_save.delete(0, tk.END)
            self.entry_save.insert(0, file_path)
            self.save_file_path = file_path  # Сохраняем путь для сохранения

    def step3(self):
        if not validate_save_file(self.save_file_path):
            return  # Прерываем выполнение функции, если файл не валиден

        self.clear_frame()
        self.current_step = 3

        tk.Label(
            self.frame,
            text="Выберите данные для парсинга:",
            bg="lightblue",
            font=("Arial", 16),
        ).pack(pady=20)

        # Список данных для парсинга
        self.data_options = {
            "Полное юридическое наименование": True,
            "Руководитель": True,
            "Уставной капитал": False,
            "Численность персонала": False,
            "Статус": False,
            "Адрес": False,
            "Юридический адрес": False,
            "Телефон": True,
            "E-mail": True,
            "Сайт": True,
        }

        # Создаем чекбоксы для каждого варианта
        for option, enabled in self.data_options.items():
            # Если состояние чекбокса уже сохранено, используем его
            var = self.checkbox_vars.get(option, tk.BooleanVar(value=enabled))
            self.checkbox_vars[option] = var
            cb = tk.Checkbutton(
                self.frame,
                text=option,
                variable=var,
                bg="lightblue",
                font=("Arial", 13),
                state=(
                    tk.NORMAL if enabled else tk.DISABLED
                ),  # Отключаем неактивные чекбоксы
            )
            cb.pack(anchor=tk.W)

        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=40)

        tk.Button(
            button_frame, text="Назад", command=self.step2, font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=40)
        tk.Button(
            button_frame, text="Далее", command=self.step4, font=("Arial", 14), width=15
        ).pack(side=tk.RIGHT, padx=40)

    def step4(self):
        # Сохраняем выбранные данные для парсинга
        self.selected_data = [
            option for option, var in self.checkbox_vars.items() if var.get()
        ]
        print(f"Выбранные данные для парсинга: {self.selected_data}")

        # Переходим к выбору источника данных
        self.clear_frame()
        self.current_step = 4
        tk.Label(
            self.frame,
            text="Выберите источник данных:",
            bg="lightblue",
            font=("Arial", 16),
        ).pack(pady=80)

        # Пример нескольких вариантов выбора
        self.var_source = tk.StringVar(value=self.source_option)
        options = [
            ("List-org", "https://www.list-org.com/"),
            ("Источник 2", "https://example.com/"),
            ("Источник 3", "https://another-example.com/"),
        ]

        for text, value in options:
            rb = tk.Radiobutton(
                self.frame,
                text=text,
                variable=self.var_source,
                value=value,
                bg="lightblue",
                font=("Arial", 14),
            )
            rb.pack(anchor=tk.W)

        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=80)

        tk.Button(
            button_frame, text="Назад", command=self.step3, font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=60)
        tk.Button(
            button_frame,
            text="Завершить",
            command=self.finish,
            font=("Arial", 14),
            width=15,
        ).pack(side=tk.RIGHT, padx=60)

    def finish(self):
        # Используем сохраненные данные
        inn_file = self.inn_file
        save_file = self.save_file_path
        source = self.var_source.get()

        # Проверяем, поддерживается ли источник
        if source not in ["https://www.list-org.com/"]:
            messagebox.showerror("Ошибка", "Источник пока не поддерживается.")
            return  # Прерываем выполнение функции

        # Передаем выбранные данные в парсер
        self.parser.selected_data = self.selected_data

        # Вызываем функцию для обработки данных
        success = self.parser.process_data(inn_file, save_file, source)

        if success:
            messagebox.showinfo(
                "Готово", f"Данные успешно обработаны и сохранены в файл:\n{save_file}"
            )
            self.root.destroy()  # Закрываем программу после успешного завершения
        else:
            messagebox.showerror("Ошибка", "Произошла ошибка при обработке данных.")

    def clear_frame(self):
        for widget in self.frame.winfo_children():
            widget.destroy()
        # Удаляем ссылки на уничтоженные виджеты
        if hasattr(self, "entry_inn"):
            del self.entry_inn
        if hasattr(self, "entry_save"):
            del self.entry_save
