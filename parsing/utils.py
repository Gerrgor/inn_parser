import os
from tkinter import messagebox
import pandas as pd

def validate_inn_file(inn_file):
    if not inn_file or not os.path.isfile(inn_file):
        messagebox.showerror("Ошибка", "Пожалуйста, выберите корректный файл с ИНН.")
        return False
    return True

def validate_save_file(save_file):
    if not save_file:
        messagebox.showerror("Ошибка", "Пожалуйста, укажите путь для сохранения результата.")
        return False
    return True

def validate_column_index(file_path, column_index):
    try:
        df = pd.read_excel(file_path)
        if column_index < 1 or column_index > len(df.columns):
            messagebox.showerror("Ошибка", "Столбец с указанным номером не существует.")
            return False
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл: {e}")
        return False

def validate_inn_in_column(file_path, column_index):
    try:
        import pandas as pd
        df = pd.read_excel(file_path)

        # Получаем данные из указанного столбца
        column_data = df.iloc[:, column_index - 1].astype(str).str.strip()

        # Проверяем, есть ли хотя бы одно значение, соответствующее ИНН
        has_valid_inn = any(value.isdigit() and len(value) == 10 for value in column_data)

        if not has_valid_inn:
            messagebox.showerror("Ошибка", "В указанном столбце отсутствуют корректные ИНН")
            return False
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось проверить столбец: {e}")
        return False