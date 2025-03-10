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
        ).pack(pady=10)

        # Поле для ввода пути к файлу
        self.entry_inn = tk.Entry(self.frame, width=50, font=("Arial", 14))
        self.entry_inn.pack(pady=10)
        if hasattr(self, 'inn_file'):  # Вставляем сохраненный путь, если он есть
            self.entry_inn.insert(0, self.inn_file)  # Вставляем сохраненный путь
        
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
        if hasattr(self, 'inn_column'):  # Вставляем сохраненное значение, если оно есть
            self.entry_column.insert(0, self.inn_column)

        # Кнопка "Далее"
        button_frame = tk.Frame(self.frame, bg="lightblue")
        button_frame.pack(pady=20)
        tk.Button(
            button_frame, text="Далее", command=self.step2, font=("Arial", 14), width=15
        ).pack()

    def load_inn_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            if hasattr(self, 'entry_inn'):  # Проверяем, существует ли виджет
                self.entry_inn.delete(0, tk.END)
                self.entry_inn.insert(0, file_path)
            self.inn_file = file_path  # Сохраняем путь к файлу

    def step2(self, returning=False):
        if not returning:
            # Сохраняем номер столбца
            if hasattr(self, 'entry_column'):
                self.inn_column = self.entry_column.get().strip()
            else:
                self.inn_column = ""

            # Пропускаем валидацию номера столбца, если мы возвращаемся назад
            if not self.inn_column or not self.inn_column.isdigit():
                messagebox.showerror("Ошибка", "Номер столбца должен быть числом.")
                return

            # Преобразуем номер столбца в целое число
            column_index = int(self.inn_column)

            # Проверяем файл
            if not validate_inn_file(self.inn_file):
                return  # Прерываем выполнение, если файл не валиден
            
            # Проверяем номер столбца
            if not validate_column_index(self.inn_file, column_index):
                return  # Прерываем выполнение, если номер столбца не валиден
            
            # Проверяем, что в столбце есть хотя бы один корректный ИНН
            if not validate_inn_in_column(self.inn_file, column_index):
                return  # Прерываем выполнение, если в столбце нет ИНН

        # Переходим к следующему шагу
        self.clear_frame()
        self.current_step = 2

        tk.Label(
            self.frame,
            text="Выберите путь для сохранения результата:",
            bg="lightblue",
            font=("Arial", 16),
        ).pack(pady=20)

        # Поле для ввода пути сохранения
        self.entry_save = tk.Entry(self.frame, width=50, font=("Arial", 14))
        self.entry_save.pack(pady=50)
        if hasattr(self, 'save_file_path'):  # Вставляем сохраненный путь, если он есть
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
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            if hasattr(self, 'entry_save'):  # Проверяем, существует ли виджет
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
            "Уставной капитал": True,
            "Численность персонала": True,
            "Статус": True,
            "Адрес": True,
            "Юридический адрес": False,
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

        # Создаем чекбоксы для каждого варианта
        for option, enabled in self.data_options.items():
            # Если состояние чекбокса уже сохранено, используем его
            var = self.checkbox_vars.get(option, tk.BooleanVar(value=enabled))
            self.checkbox_vars[option] = var
            cb = tk.Checkbutton(
                checkboxes_frame,
                text=option,
                variable=var,
                bg="lightblue",
                font=("Arial", 13),
                state=(
                    tk.NORMAL if enabled else tk.DISABLED
                ),  # Отключаем неактивные чекбоксы
            )
            cb.pack(anchor=tk.W)

        # Фрейм для кнопок "Выбрать все" и "Очистить выбор"
        buttons_frame = tk.Frame(checkbox_frame, bg="lightblue")
        buttons_frame.pack(side=tk.RIGHT, padx=20)

        # Кнопка "Выбрать все"
        select_all_button = tk.Button(
            buttons_frame,
            text="Выбрать все",
            command=self.select_all_checkboxes,
            font=("Arial", 12),
            width=15,
        )
        select_all_button.pack(pady=20)

        # Кнопка "Очистить выбор"
        clear_all_button = tk.Button(
            buttons_frame,
            text="Очистить выбор",
            command=self.clear_all_checkboxes,
            font=("Arial", 12),
            width=15,
        )
        clear_all_button.pack(pady=20)

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
        """
        Выбирает все доступные чекбоксы.
        """
        for option, var in self.checkbox_vars.items():
            if self.data_options[option]:  # Проверяем, что чекбокс доступен
                var.set(True)

    def clear_all_checkboxes(self):
        """
        Снимает выбор со всех чекбоксов.
        """
        for var in self.checkbox_vars.values():
            var.set(False)

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
        if hasattr(self, "entry_column"):
            del self.entry_column
        if hasattr(self, "entry_save"):
            del self.entry_save

