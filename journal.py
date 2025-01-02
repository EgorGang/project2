"""
Модуль для работы с журналом студентов.
Содержит функциональность добавления, удаления студентов, ведения оценок и экспорта данных.
"""
import pandas as pd
import numpy as np
import tempfile
import os
from exceptions import JournalError

class Journal:
    """
    Класс для работы с журналом студентов.
    """
    def __init__(self):
        """
        Инициализация пустой таблицы журнала с колонкой для среднего балла.
        """
        self.table = pd.DataFrame(columns=["Средний балл"])

    def add_student(self, name):
        """
        Добавляет нового ученика в журнал.
        :param name: Имя ученика.
        :return: Сообщение об успехе или ошибке.
        :raises JournalError: Если имя пустое или уже существует.
        """
        if not name:
            raise JournalError("Имя не может быть пустым.")
        if name in self.table.index:
            return "Такой ученик уже есть!"
        self.table.loc[name] = np.nan
        self.table.loc[name, "Средний балл"] = 0.0
        return f"Ученик {name} добавлен."

    def remove_student(self, name):
        """
        Удаляет ученика из журнала.
        :param name: Имя ученика.
        :return: Сообщение об успехе или ошибке.
        :raises JournalError: Если ученика нет в журнале.
        """
        if name not in self.table.index:
            return "Такого ученика нет!"
        self.table.drop(name, inplace=True)
        return f"Ученик {name} удалён."

    def add_mark(self, name, day, mark):
        """
        Добавляет оценку ученику за указанный день.
        :param name: Имя ученика.
        :param day: Число месяца.
        :param mark: Оценка.
        :return: Сообщение об успехе или ошибке.
        :raises JournalError: Если имя ученика отсутствует в журнале.
        """
        if name not in self.table.index:
            raise JournalError("Такого ученика нет!")
        column = f"{day}"
        if column not in self.table.columns:
            self.table[column] = np.nan
        self.table.loc[name, column] = mark
        self._update_average(name)
        return f"Оценка {mark} добавлена для {name} за {day} число."
    
    def remove_mark(self, name, day):
        """
        Удаляет оценку ученику за указанный день.
        :param name: Имя ученика.
        :param day: Число месяца.
        :return: Сообщение об успехе или ошибке.
        :raises JournalError: Если имя ученика отсутствует в журнале.
        """
        if name in self.table.index:
            column = f"{day}"
            if column in self.table.columns:
                self.table.loc[name, column] = np.nan
                self._update_average(name)
                return f"Оценка удалена у ученика {name} за {day} число."
            else:
                return "За это число нет оценок!"
        else:
            return "Такого ученика нет в журнале!"

    def change_mark(self, name, day, new_mark):
        """
        Изменяет оценку ученику за указанный день.
        :param name: Имя ученика.
        :param day: Число месяца.
        :param new_mark: Новая оценка.
        :return: Сообщение об успехе или ошибке.
        :raises JournalError: Если имя ученика отсутствует в журнале.
        """
        if name in self.table.index:
            column = f"{day}"
            if column in self.table.columns:
                self.table.loc[name, column] = new_mark
                self._update_average(name)
                return f"Оценка исправлена у ученика {name} за {day} число."
            else:
                return "За это число нет оценок!"
        else:
            return "Такого ученика нет в журнале!"

    def _update_average(self, name):
        """
        Обновляет средний балл ученика.
        :param name: Имя ученика.
        """
        marks = self.table.loc[name, self.table.columns != "Средний балл"]
        mean = marks.mean(skipna=True)
        self.table.loc[name, "Средний балл"] = round(mean, 2) if not np.isnan(mean) else 0.0

    def student_info(self, name):
        """
        Возвращает информацию об ученике.
        :param name: Имя ученика.
        :return: Информация об ученике в виде строки.
        """
        if name not in self.table.index:
            return "Такого ученика нет!"
        info = f"Информация об ученике {name}:\n"
        student_data = self.table.loc[name]
        for column, mark in student_data.items():
            if not pd.isna(mark):
                info += f"{column}: {mark}\n"
        return info

    def export_to_excel(self, filename):
        """
        Экспортирует журнал в файл Excel.
        :param filename: Имя файла для экспорта.
        :return: Путь к созданному файлу.
        :raises JournalError: Ошибка при сохранении файла или некорректный формат.
        """
        try:
            # Создаем временную директорию
            temp_dir = tempfile.gettempdir()
            # Полный путь к файлу
            filepath = os.path.join(temp_dir, filename)
            # Сохранение таблицы в Excel
            self.table.to_excel(filepath)
            return filepath  # Возвращаем путь к файлу
        except Exception as e:
            raise JournalError(f"Ошибка при экспорте: {e}")