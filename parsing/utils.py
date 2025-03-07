import os
from tkinter import messagebox

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