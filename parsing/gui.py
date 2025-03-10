import tkinter as tk
from tkinter import filedialog, messagebox
from parser import Parser
from utils import validate_inn_file, validate_save_file, validate_column_index, validate_inn_in_column


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Парсер данных по ИНН")
        self.root.geometry("700x500")
        self.root.configure(bg="lightblue")
        self.root.attributes("-alpha", 0.98)
        self.center_window()

        # Инициализация переменных состояния
        self.current_step = 1
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

        # Запуск первого шага
        self.step1()

    def center_window(self):
        """Центрирует окно на экране."""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (700 // 2)
        y = (screen_height // 2) - (500 // 2)
        self.root.geometry(f"700x500+{x}+{y}")

    def clear_frame(self):
        """Очищает текущий фрейм и удаляет ссылки на виджеты."""
        for widget in self.frame.winfo_children():
            widget.destroy()
        # Удаляем ссылки на уничтоженные виджеты
        for attr in ["entry_inn", "entry_column", "entry_save"]:
            if hasattr(self, attr):
                delattr(self, attr)

    def step1(self):
        """Шаг 1: Выбор файла с ИНН и номера столбца."""
        self.clear_frame()
        self.current_step = 1

        # Заголовок
        tk.Label(
            self.frame, text="Выберите файл с ИНН:", bg="lightblue", font=("Arial", 16)
        ).pack(pady=10)

        # Поле для ввода пути к файлу
        self.entry_inn = tk.Entry(self.frame, width=50, font=("Arial", 14))
        self.entry_inn.pack(pady=10)
        if self.inn_file:  # Вставляем сохраненный путь, если он есть
            self.entry_inn.insert(0, self.inn_file)

        # Кнопка "Обзор"
        tk.Button(
            self.frame,
            text="Обзор",
            command=self.load_inn_file,
            font=("Arial", 14),
            width=15,
        ).pack(pady=10)

        # Поле для ввода номера столбца
        tk.Label(
            self.frame, text="Введите номер столбца с ИНН:", bg="lightblue", font=("Arial", 16)
        ).pack(pady=10)

        self.entry_column = tk.Entry(self.frame, width=10, font=("Arial", 14))
        self.entry_column.pack(pady=10)
        if self.inn_column:  # Вставляем сохраненное значение, если оно есть
            self.entry_column.insert(0, self.inn_column)

        # Кнопка "Далее"
        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=20)
        tk.Button(
            button_frame, text="Далее", command=self.step2, font=("Arial", 14), width=15
        ).pack()

    def load_inn_file(self):
        """Загружает файл с ИНН."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.inn_file = file_path
            if hasattr(self, 'entry_inn'):
                self.entry_inn.delete(0, tk.END)
                self.entry_inn.insert(0, file_path)

    def step2(self, returning=False):
        """Шаг 2: Выбор пути для сохранения результата."""
        if not returning:
            # Сохраняем номер столбца
            self.inn_column = self.entry_column.get().strip()

            # Проверяем, что номер столбца является числом
            if not self.inn_column or not self.inn_column.isdigit():
                messagebox.showerror("Ошибка", "Номер столбца должен быть числом.")
                return

            # Проверяем файл и номер столбца
            column_index = int(self.inn_column)
            if not validate_inn_file(self.inn_file):
                return
            if not validate_column_index(self.inn_file, column_index):
                return
            if not validate_inn_in_column(self.inn_file, column_index):
                return

        # Переходим к следующему шагу
        self.clear_frame()
        self.current_step = 2

        # Заголовок
        tk.Label(
            self.frame,
            text="Выберите путь для сохранения результата:",
            bg="lightblue",
            font=("Arial", 16),
        ).pack(pady=20)

        # Поле для ввода пути сохранения
        self.entry_save = tk.Entry(self.frame, width=50, font=("Arial", 14))
        self.entry_save.pack(pady=50)
        if self.save_file_path:  # Вставляем сохраненный путь, если он есть
            self.entry_save.insert(0, self.save_file_path)

        # Кнопка "Обзор"
        tk.Button(
            self.frame,
            text="Обзор",
            command=self.save_file,
            font=("Arial", 14),
            width=15,
        ).pack(pady=20)

        # Кнопки "Назад" и "Далее"
        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=40)
        tk.Button(
            button_frame, text="Назад", command=self.step1, font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=40)
        tk.Button(
            button_frame, text="Далее", command=self.step3, font=("Arial", 14), width=15
        ).pack(side=tk.RIGHT, padx=40)

    def save_file(self):
        """Выбирает путь для сохранения результата."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.save_file_path = file_path
            if hasattr(self, 'entry_save'):
                self.entry_save.delete(0, tk.END)
                self.entry_save.insert(0, file_path)

    def step3(self):
        """Шаг 3: Выбор данных для парсинга."""
        if not validate_save_file(self.save_file_path):
            return

        self.clear_frame()
        self.current_step = 3

        # Заголовок
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
            "Уставной капитал": True,
            "Численность персонала": True,
            "Статус": True,
            "Адрес": True,
            "Юридический адрес": True,
            "Телефон": True,
            "E-mail": True,
            "Сайт": True,
        }

        # Фрейм для чекбоксов и кнопок
        checkbox_frame = tk.Frame(self.frame, bg="lightblue")
        checkbox_frame.pack(fill=tk.X, padx=20)

        # Фрейм для чекбоксов
        checkboxes_frame = tk.Frame(checkbox_frame, bg="lightblue")
        checkboxes_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Создаем чекбоксы
        for option, enabled in self.data_options.items():
            var = self.checkbox_vars.get(option, tk.BooleanVar(value=enabled))
            self.checkbox_vars[option] = var
            cb = tk.Checkbutton(
                checkboxes_frame,
                text=option,
                variable=var,
                bg="lightblue",
                font=("Arial", 13),
                state=tk.NORMAL if enabled else tk.DISABLED,
            )
            cb.pack(anchor=tk.W)

        # Фрейм для кнопок "Выбрать все" и "Очистить выбор"
        buttons_frame = tk.Frame(checkbox_frame, bg="lightblue")
        buttons_frame.pack(side=tk.RIGHT, padx=20)

        # Кнопка "Выбрать все"
        tk.Button(
            buttons_frame,
            text="Выбрать все",
            command=self.select_all_checkboxes,
            font=("Arial", 12),
            width=15,
        ).pack(pady=20)

        # Кнопка "Очистить выбор"
        tk.Button(
            buttons_frame,
            text="Очистить выбор",
            command=self.clear_all_checkboxes,
            font=("Arial", 12),
            width=15,
        ).pack(pady=20)

        # Фрейм для кнопок "Назад" и "Далее"
        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=40)

        tk.Button(
            button_frame, text="Назад", command=lambda: self.step2(returning=True), font=("Arial", 14), width=15
        ).pack(side=tk.LEFT, padx=40)
        tk.Button(
            button_frame, text="Далее", command=self.step4, font=("Arial", 14), width=15
        ).pack(side=tk.RIGHT, padx=40)

    def select_all_checkboxes(self):
        """Выбирает все доступные чекбоксы."""
        for option, var in self.checkbox_vars.items():
            if self.data_options[option]:
                var.set(True)

    def clear_all_checkboxes(self):
        """Снимает выбор со всех чекбоксов."""
        for var in self.checkbox_vars.values():
            var.set(False)

    def step4(self):
        """Шаг 4: Выбор источника данных."""
        self.selected_data = [
            option for option, var in self.checkbox_vars.items() if var.get()
        ]
        print(f"Выбранные данные для парсинга: {self.selected_data}")

        self.clear_frame()
        self.current_step = 4

        # Заголовок
        tk.Label(
            self.frame,
            text="Выберите источник данных:",
            bg="lightblue",
            font=("Arial", 16),
        ).pack(pady=80)

        # Варианты источников данных
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

        # Фрейм для кнопок "Назад" и "Завершить"
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
        """Завершает работу программы."""
        inn_file = self.inn_file
        save_file = self.save_file_path
        source = self.var_source.get()

        if source not in ["https://www.list-org.com/"]:
            messagebox.showerror("Ошибка", "Источник пока не поддерживается.")
            return

        self.parser.selected_data = self.selected_data
        success = self.parser.process_data(inn_file, save_file, source)

        if success:
            messagebox.showinfo(
                "Готово", f"Данные успешно обработаны и сохранены в файл:\n{save_file}"
            )
            self.root.destroy()
        else:
            messagebox.showerror("Ошибка", "Произошла ошибка при обработке данных.")