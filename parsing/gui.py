import tkinter as tk
from tkinter import filedialog, messagebox
from parser import Parser
from utils import validate_inn_file, validate_save_file, validate_column_index, validate_inn_in_column
import webbrowser

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Парсер данных по ИНН")
        self.root.geometry("800x600")
        self.root.configure(bg="lightblue")
        self.root.attributes("-alpha", 0.98)
        self.center_window()

        # Инициализация переменных состояния
        self.current_step = 0
        self.inn_file = ""
        self.save_file_path = ""
        self.source_option = ""
        self.selected_data = []
        self.checkbox_vars = {}
        self.inn_column = ""

        # Основной фрейм
        self.frame = tk.Frame(self.root, bg="lightblue")
        self.frame.pack(expand=True)

        # Инициализация парсера
        self.parser = Parser()

        # Данные для парсинга в зависимости от сайта
        self.data_options = {
            "https://www.list-org.com/": [
                "Полное юридическое наименование",
                "Руководитель",
                "Уставной капитал",
                "Численность персонала",
                "Статус",
                "Адрес",
                "Юридический адрес",
                "Телефон",
                "E-mail",
                "Сайт",
            ],
            "https://zachestnyibiznes.ru/": [
                "Полное юридическое наименование",
                "Руководитель",
                'ОГРН',
                'Дата регистрации',
                'Адрес регистрации',
                "Статус",
                "Уставной капитал",
                "Численность персонала",
                'Основное направление деятельности',
                'Доход',
                'Расход',
                "Телефон",
                "E-mail",
                "Сайт",
            ],
            "https://www.rusprofile.ru/": [
                "Полное юридическое наименование",
                "Руководитель",
                "Уставной капитал",
                "Статус",
                "Адрес",
                "Телефон",
                "E-mail",
                "Сайт",
            ],
        }

        # Запуск первого шага
        self.step0()

    def center_window(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (800 // 2)
        y = (screen_height // 2) - (600 // 2)
        self.root.geometry(f"800x600+{x}+{y}")

    def clear_frame(self):
        for widget in self.frame.winfo_children():
            widget.destroy()
        for attr in ["entry_inn", "entry_column", "entry_save"]:
            if hasattr(self, attr):
                delattr(self, attr)

    def step0(self):
        self.clear_frame()
        self.current_step = 0

        tk.Label(
            self.frame,
            text="Добро пожаловать в программу парсинга данных по ИНН!",
            bg="lightblue",
            font=("Arial", 16, "bold"),
        ).pack(pady=10)

        info_text = (
            "Эта программа позволяет автоматически собирать данные о компаниях по их ИНН.\n\n"
            "Рекомендации по использованию:\n\n"
            "1. Подготовьте файл Excel с ИНН компаний.\n\n"
            "2. Убедитесь, что файл содержит не более 50 ИНН для стабильной работы.\n\n"
            "3. Выберите данные, которые хотите получить.\n\n"
            "Если у вас есть вопросы, предложения или пожелания, пишите в Telegram:"
        )
        tk.Label(
            self.frame,
            text=info_text,
            bg="lightblue",
            font=("Arial", 14),
            justify=tk.LEFT,
            wraplength=600,
        ).pack(pady=5)

        telegram_link = tk.Label(
            self.frame,
            text="@gregormsh",
            bg="lightblue",
            font=("Arial", 14, "underline"),
            fg="blue",
            cursor="hand2",
        )
        telegram_link.pack()
        telegram_link.bind("<Button-1>", lambda e: self.open_telegram("https://t.me/gregormsh"))

        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=10)
        tk.Button(
            button_frame,
            text="Далее",
            command=self.step1,
            font=("Arial", 14),
            width=15,
        ).pack()

    def open_telegram(self, url):
        webbrowser.open(url)

    def step1(self):
        self.clear_frame()
        self.current_step = 1

        tk.Label(
            self.frame, text="Выберите файл с ИНН:", bg="lightblue", font=("Arial", 16)
        ).pack(pady=10)

        self.entry_inn = tk.Entry(self.frame, width=50, font=("Arial", 14))
        self.entry_inn.pack(pady=10)
        if self.inn_file:
            self.entry_inn.insert(0, self.inn_file)

        tk.Button(
            self.frame,
            text="Обзор",
            command=self.load_inn_file,
            font=("Arial", 14),
            width=15,
        ).pack(pady=10)

        tk.Label(
            self.frame, text="Введите номер столбца с ИНН:", bg="lightblue", font=("Arial", 16)
        ).pack(pady=10)

        self.entry_column = tk.Entry(self.frame, width=10, font=("Arial", 14))
        self.entry_column.pack(pady=10)
        if self.inn_column:
            self.entry_column.insert(0, self.inn_column)

        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=20)
        tk.Button(
            button_frame, text="Назад", command=self.step0, font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=40)
        tk.Button(
            button_frame, text="Далее", command=self.step2, font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=40)

    def load_inn_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.inn_file = file_path
            if hasattr(self, 'entry_inn'):
                self.entry_inn.delete(0, tk.END)
                self.entry_inn.insert(0, file_path)

    def step2(self, returning=False):
        if not returning:
            self.inn_column = self.entry_column.get().strip()
            if not self.inn_column or not self.inn_column.isdigit():
                messagebox.showerror("Ошибка", "Номер столбца должен быть числом.")
                return

            column_index = int(self.inn_column)
            if not validate_inn_file(self.inn_file):
                return
            if not validate_column_index(self.inn_file, column_index):
                return
            if not validate_inn_in_column(self.inn_file, column_index):
                return

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
        if self.save_file_path:
            self.entry_save.insert(0, self.save_file_path)

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
            self.save_file_path = file_path
            if hasattr(self, 'entry_save'):
                self.entry_save.delete(0, tk.END)
                self.entry_save.insert(0, file_path)

    def step3(self):
        if not validate_save_file(self.save_file_path):
            return

        self.clear_frame()
        self.current_step = 3

        tk.Label(
            self.frame,
            text="Выберите источник данных:",
            bg="lightblue",
            font=("Arial", 16),
        ).pack(pady=80)

        self.var_source = tk.StringVar(value=self.source_option)
        options = [
            ("List-org", "https://www.list-org.com/"),
            ("Rusprofile", "https://www.rusprofile.ru/"),
            ("Zachestnyibiznes", "https://zachestnyibiznes.ru/"),
        ]

        for text, value in options:
            rb = tk.Radiobutton(
                self.frame,
                text=text,
                variable=self.var_source,
                value=value,
                bg="lightblue",
                font=("Arial", 14),
                command=lambda v=value: self.update_source_option(v),  # Обновляем выбор при изменении
            )
            rb.pack(anchor=tk.W)

        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=80)

        tk.Button(
            button_frame, text="Назад", command=lambda: self.step2(returning=True), font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=40)
        tk.Button(
            button_frame,
            text="Далее",
            command=self.validate_source_and_proceed,
            font=("Arial", 14),
            width=15,
        ).pack(side=tk.RIGHT, padx=60)
    
    def update_source_option(self, value):
        """Обновляет выбранный сайт при изменении Radiobutton."""
        self.source_option = value
    
    def validate_source_and_proceed(self):
        source = self.var_source.get()
        self.source_option = source
        if source not in ["https://www.list-org.com/", "https://zachestnyibiznes.ru/"]:
            messagebox.showerror("Ошибка", "Источник пока не поддерживается.")
            return
        self.step4()  # Переходим к шагу 4 только если источник поддерживается

    def step4(self):
        self.clear_frame()
        self.current_step = 4

        tk.Label(
            self.frame,
            text="Выберите данные для парсинга:",
            bg="lightblue",
            font=("Arial", 16),
        ).pack(pady=20)

        checkbox_frame = tk.Frame(self.frame, bg="lightblue")
        checkbox_frame.pack(fill=tk.X, padx=20)

        checkboxes_frame = tk.Frame(checkbox_frame, bg="lightblue")
        checkboxes_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Получаем доступные данные для выбранного сайта
        available_data = self.data_options.get(self.source_option, [])

        for option in available_data:
            var = self.checkbox_vars.get(option, tk.BooleanVar(value=True))
            self.checkbox_vars[option] = var
            cb = tk.Checkbutton(
                checkboxes_frame,
                text=option,
                variable=var,
                bg="lightblue",
                font=("Arial", 13),
            )
            cb.pack(anchor=tk.W)

        buttons_frame = tk.Frame(checkbox_frame, bg="lightblue")
        buttons_frame.pack(side=tk.RIGHT, padx=20)

        tk.Button(
            buttons_frame,
            text="Выбрать все",
            command=self.select_all_checkboxes,
            font=("Arial", 12),
            width=15,
        ).pack(pady=20)

        tk.Button(
            buttons_frame,
            text="Очистить выбор",
            command=self.clear_all_checkboxes,
            font=("Arial", 12),
            width=15,
        ).pack(pady=20)

        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=40)

        tk.Button(
            button_frame, text="Назад", command=self.step3, font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=40)
        tk.Button(
            button_frame,
            text="Завершить",
            command=self.finish,
            font=("Arial", 14),
            width=15,
        ).pack(side=tk.RIGHT, padx=40)

    def select_all_checkboxes(self):
        for option, var in self.checkbox_vars.items():
            var.set(True)

    def clear_all_checkboxes(self):
        for var in self.checkbox_vars.values():
            var.set(False)

    def finish(self):
        inn_file = self.inn_file
        save_file = self.save_file_path
        source = self.source_option
        
        # Получаем порядок столбцов для выбранного сайта
        column_order = ["ИНН"] + self.data_options.get(source, [])  # Добавляем "ИНН" в начало
        
        # Обновляем выбранные данные на основе чекбоксов
        self.parser.selected_data = [
            option for option, var in self.checkbox_vars.items() if var.get()
        ]
        success = self.parser.process_data(inn_file, save_file, source, column_order)

        if success:
            messagebox.showinfo(
                "Готово", f"Данные успешно обработаны и сохранены в файл:\n{save_file}"
            )
            self.root.destroy()
        else:
            messagebox.showerror("Ошибка", "Произошла ошибка при обработке данных.")