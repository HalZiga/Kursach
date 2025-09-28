# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import traceback
import xlsxwriter

# --- КОНФИГУРАЦИЯ ---
header_row_index_level1 = 9
# Заголовки для верхнего уровня (30 элементов, индексы 0-29)
level0_headers = [
    'Общие данные', 'Общие данные', 'Общие данные', 'Общие данные',
    'Общие данные', 'Общие данные',
    'количество', 'количество', 'количество', 'количество',
    'лекции', 'лекции',
    'практич. занятия', 'практич. занятия',
    'лаборат. занятия', 'лаборат. занятия',

    # Распределение учебной нагрузки (9 столбцов, включая КСР и Дополнительно)
    'Распределение учебной нагрузки (в часах)', 'Распределение учебной нагрузки (в часах)',
    'Распределение учебной нагрузки (в часах)', 'Распределение учебной нагрузки (в часах)',
    'Распределение учебной нагрузки (в часах)', 'Распределение учебной нагрузки (в часах)',
    'Распределение учебной нагрузки (в часах)', 'Распределение учебной нагрузки (в часах)',
    'Распределение учебной нагрузки (в часах)',  # Индекс 24 (Y) - Дополнительно

    # НОВЫЙ БЛОК: Всего + В том числе (1 + 4 = 5 столбцов)
    'Всего',  # Индекс 25 (Z)
    'В том числе', 'В том числе', 'В том числе', 'В том числе'  # Индексы 26-29 (AA-AD)
]

# Заголовки для нижнего уровня (строка 2, 30 элементов, индексы 0-29)
level1_headers = [
    'Цикл дисциплины по уч. плану', 'Наименование дисциплины',
    'Шифр направления/специальности', 'Наименование направления/специальности (профиль/специализация)',
    'Форма обучения', 'Курс',
    'студентов', 'потоков', 'групп', 'подгрупп',
    'всего', 'в дистанционном формате (ЭОР)',
    'всего', 'в дистанционном формате (ЭОР)',
    'всего', 'в дистанционном формате (ЭОР)',

    'КСР',
    'Консультации', 'Контр. работы', 'Зачеты', 'Экзамены', 'Практики',
    'Курсовые работы', 'Выпускные квалификационные работы',

    # Измененная структура для вертикальных слияний
    'Дополнительно',  # Индекс 24 (Y2) - Будет покрыт Y1:Y2 слиянием
    '',  # Индекс 25 (Z2) - Пустая ячейка, покрыта Z1:Z2 слиянием для 'Всего'
    'дн.',  # Индекс 26 (AA2)
    'веч.',  # Индекс 27 (AB2)
    'заоч.',  # Индекс 28 (AC2)
    'в дистанционном формате (ЭОР)'  # Индекс 29 (AD2)
]


def get_unique_names(series):
    """
    Извлекает все уникальные имена из серии pandas.
    Преобразует имена к нижнему регистру для нормализации.
    """
    unique_names = set()
    text = series.astype(str).str.replace(r'[\s\n\r\t]+', ',', regex=True).str.strip(',')
    for value in text.dropna():
        parts = re.split(r',', value)
        for part in parts:
            clean_name = re.sub(r'\(.*?\)|[0-9.,\-]', '', part).strip()
            if clean_name and clean_name.lower() != 'nan':
                unique_names.add(clean_name.lower())
    return sorted(list(unique_names))


def parse_groups_subgroups(s):
    """
    Парсит строку с группами и подгруппами.
    """
    if pd.isna(s): return 0, 0
    s = str(s).strip()
    match = re.match(r'(\d+)\((\d+)\)', s)
    if match:
        return int(match.group(1)), int(match.group(2))
    elif re.match(r'^\d+$', s):
        return int(s), int(s)
    return 0, 0


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработчик файлов F2")
        self.root.geometry("1000x600")

        self.df_f1 = None
        self.df_f2_processed = None
        self.editable_cols = [
            'Количество студентов', 'Количество потоков', 'Количество групп',
            'Количество подгрупп', 'Лекции: всего', 'Практические занятия: всего',
            'Лабораторные занятия: всего', 'Имена_Raw'
        ]

        self.f2_calc_cols = [
            'Лекции_По_Плану_F2', 'Практика_По_Плану_F2', 'Лабы_По_Плану_F2',
            'Экзамены_F2_Наличие', 'Практики_F2_Наличие', 'ВКР_F2_Наличие', 'Дополнительно_F2'
        ]

        self.setup_ui()

    def setup_ui(self):
        """
        Настраивает элементы пользовательского интерфейса.
        """
        # Фрейм для кнопок и статуса
        control_frame = ttk.Frame(self.root)
        control_frame.pack(pady=10, padx=10, fill="x")

        # Кнопка для выбора файла
        btn_select_file = ttk.Button(
            control_frame,
            text="Выбрать и обработать файл F2-2022.xlsx",
            command=self.process_file
        )
        btn_select_file.pack(side="left", padx=5)

        # Кнопка для сохранения отчёта
        btn_save_report = ttk.Button(
            control_frame,
            text="Сохранить отчёт F1",
            command=self.save_report,
            state="disabled"
        )
        btn_save_report.pack(side="left", padx=5)

        # Метка для отображения статуса
        self.status_label = ttk.Label(control_frame, text="Ожидание файла...", foreground="blue")
        self.status_label.pack(side="left", padx=15)

        # Фрейм для таблицы
        table_frame = ttk.Frame(self.root)
        table_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Настройка таблицы (Treeview)
        self.columns = (
            'id', 'Цикл', 'Дисциплина', 'Направление', 'Курс',
            'Студенты', 'Потоки', 'Группы', 'Подгруппы', 'Лекции', 'Практика',
            'Лабы', 'КСР', 'Доп', 'Всего', 'ДН', 'ЭОР', 'Имена'
        )
        self.tree = ttk.Treeview(
            table_frame,
            columns=self.columns,
            show='headings'
        )

        # Настройка заголовков столбцов
        self.tree.heading('id', text='ID')
        self.tree.heading('Цикл', text='Цикл')
        self.tree.heading('Дисциплина', text='Наименование дисциплины')
        self.tree.heading('Направление', text='Шифр')
        self.tree.heading('Курс', text='Курс')
        self.tree.heading('Студенты', text='Студенты')
        self.tree.heading('Потоки', text='Потоки')
        self.tree.heading('Группы', text='Группы')
        self.tree.heading('Подгруппы', text='Подгруппы')
        self.tree.heading('Лекции', text='Лекции')
        self.tree.heading('Практика', text='Практика')
        self.tree.heading('Лабы', text='Лабы')
        self.tree.heading('КСР', text='КСР')
        self.tree.heading('Доп', text='Доп.')  # Дополнительно
        self.tree.heading('Всего', text='Всего')
        self.tree.heading('ДН', text='дн.')
        self.tree.heading('ЭОР', text='ЭОР')
        self.tree.heading('Имена', text='Преподаватели')

        # Настройка ширины столбцов
        self.tree.column('id', width=30, anchor=tk.CENTER)
        self.tree.column('Цикл', width=90, anchor=tk.W)
        self.tree.column('Дисциплина', width=200, anchor=tk.W)
        self.tree.column('Направление', width=80, anchor=tk.CENTER)
        self.tree.column('Курс', width=50, anchor=tk.CENTER)
        self.tree.column('Студенты', width=70, anchor=tk.CENTER)
        self.tree.column('Потоки', width=60, anchor=tk.CENTER)
        self.tree.column('Группы', width=60, anchor=tk.CENTER)
        self.tree.column('Подгруппы', width=80, anchor=tk.CENTER)
        self.tree.column('Лекции', width=80, anchor=tk.CENTER)
        self.tree.column('Практика', width=80, anchor=tk.CENTER)
        self.tree.column('Лабы', width=80, anchor=tk.CENTER)
        self.tree.column('КСР', width=50, anchor=tk.CENTER)
        self.tree.column('Доп', width=50, anchor=tk.CENTER)
        self.tree.column('Всего', width=100, anchor=tk.CENTER)
        self.tree.column('ДН', width=50, anchor=tk.CENTER)
        self.tree.column('ЭОР', width=50, anchor=tk.CENTER)
        self.tree.column('Имена', width=150, anchor=tk.W)

        self.tree.pack(side="left", fill="both", expand=True)

        # Скроллбар для таблицы
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.btn_save_report = btn_save_report

        # Привязка двойного клика к функции редактирования
        self.tree.bind("<Double-1>", self.on_double_click)

    def on_double_click(self, event):
        """
        Обрабатывает двойной клик для редактирования ячейки.
        """
        try:
            item_id = self.tree.identify_row(event.y)
            column_id = self.tree.identify_column(event.x)

            if not item_id or not column_id:
                return

            col_index = int(column_id[1:]) - 1
            column_name = self.columns[col_index]

            df_col_name = None
            if column_name == 'Студенты':
                df_col_name = 'Количество студентов'
            elif column_name == 'Потоки':
                df_col_name = 'Количество потоков'
            elif column_name == 'Группы':
                df_col_name = 'Количество групп'
            elif column_name == 'Подгруппы':
                df_col_name = 'Количество подгрупп'
            elif column_name == 'Лекции':
                df_col_name = 'Лекции: всего'
            elif column_name == 'Практика':
                df_col_name = 'Практические занятия: всего'
            elif column_name == 'Лабы':
                df_col_name = 'Лабораторные занятия: всего'
            elif column_name == 'Имена':
                df_col_name = 'Имена_Raw'

            if df_col_name and df_col_name in self.editable_cols:
                self.edit_cell(item_id, column_name, df_col_name, event.x, event.y)
            else:
                messagebox.showinfo("Информация", "Эту ячейку нельзя редактировать.")
        except Exception as e:
            messagebox.showerror("Ошибка при двойном клике", f"Произошла ошибка: {e}")
            traceback.print_exc()

    def edit_cell(self, item_id, column_name, df_col_name, x, y):
        """
        Создает временный виджет для редактирования ячейки.
        """
        try:
            current_value = self.tree.item(item_id, 'values')[self.columns.index(column_name)]

            editor = ttk.Entry(self.tree)
            editor.insert(0, current_value)
            editor.bind("<FocusOut>", lambda e: self.update_cell(item_id, df_col_name, editor.get(), editor))
            editor.bind("<Return>", lambda e: self.update_cell(item_id, df_col_name, editor.get(), editor))

            self.tree.update_idletasks()
            bbox = self.tree.bbox(item_id, column=column_name)
            if bbox:
                x_pos = bbox[0]
                y_pos = bbox[1]
                width = bbox[2]
                height = bbox[3]
                editor.place(x=x_pos, y=y_pos, width=width, height=height, anchor='nw')
                editor.focus_set()
        except Exception as e:
            messagebox.showerror("Ошибка редактирования", f"Произошла ошибка: {e}")
            traceback.print_exc()

    def update_cell(self, item_id, df_col_name, new_value, editor_widget):
        """
        Обновляет данные в DataFrame и таблице после редактирования.
        """
        editor_widget.destroy()

        row_index = self.tree.index(item_id)
        if row_index is None:
            return

        try:
            if df_col_name in ['Имена_Raw']:
                self.df_f1.at[row_index, df_col_name] = new_value
            else:
                new_value_numeric = float(new_value)
                self.df_f1.at[row_index, df_col_name] = new_value_numeric

            self.recalculate_totals(row_index)

            current_values = list(self.tree.item(item_id, 'values'))
            current_values[self.columns.index('Студенты')] = self.df_f1.at[row_index, 'Количество студентов']
            current_values[self.columns.index('Потоки')] = self.df_f1.at[row_index, 'Количество потоков']
            current_values[self.columns.index('Группы')] = self.df_f1.at[row_index, 'Количество групп']
            current_values[self.columns.index('Подгруппы')] = self.df_f1.at[row_index, 'Количество подгрупп']
            current_values[self.columns.index('Лекции')] = self.df_f1.at[row_index, 'Лекции: всего']
            current_values[self.columns.index('Практика')] = self.df_f1.at[row_index, 'Практические занятия: всего']
            current_values[self.columns.index('Лабы')] = self.df_f1.at[row_index, 'Лабораторные занятия: всего']

            # Обновление итоговых столбцов
            current_values[self.columns.index('КСР')] = self.df_f1.at[row_index, 'КСР']
            current_values[self.columns.index('Доп')] = self.df_f1.at[row_index, 'Дополнительно']
            current_values[self.columns.index('Всего')] = self.df_f1.at[row_index, 'Всего']
            current_values[self.columns.index('ДН')] = self.df_f1.at[row_index, 'дн.']
            current_values[self.columns.index('ЭОР')] = self.df_f1.at[row_index, 'в дистанционном формате (ЭОР)']

            current_values[self.columns.index('Имена')] = self.df_f1.at[row_index, 'Имена_Raw']

            self.tree.item(item_id, values=current_values)

        except ValueError:
            messagebox.showerror("Ошибка ввода", "Пожалуйста, введите числовое значение.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить данные: {e}")
            traceback.print_exc()

    def recalculate_totals(self, row_index):
        """
        Пересчитывает итоговые часы после изменения исходных данных.
        """
        try:
            f2_row = self.df_f2_processed.iloc[row_index]

            total_lectures = self.df_f1.at[row_index, 'Лекции: всего']
            total_practice = self.df_f1.at[row_index, 'Практические занятия: всего']
            total_labs = self.df_f1.at[row_index, 'Лабораторные занятия: всего']

            if pd.isna(total_lectures): total_lectures = 0
            if pd.isna(total_practice): total_practice = 0
            if pd.isna(total_labs): total_labs = 0

            # Добавляем КСР в расчет, используя 'КСР_F2' из F2
            ksr_value = f2_row['КСР_F2']
            self.df_f1.at[row_index, 'КСР'] = ksr_value  # Записываем КСР в DF F1

            has_exam = f2_row['Экзамены_F2_Наличие'] > 0

            consultations_base = 0.05 * f2_row['Лекции_По_Плану_F2']
            consultations_exam_bonus = has_exam * 2

            # Расчет часов
            self.df_f1.at[row_index, 'Консультации'] = (consultations_base + consultations_exam_bonus) * self.df_f1.at[
                row_index, 'Количество групп']

            total_plan_hours_f2 = f2_row['Лекции_По_Плану_F2'] + f2_row['Практика_По_Плану_F2'] + f2_row[
                'Лабы_По_Плану_F2']

            self.df_f1.at[row_index, 'Контр. работы'] = 0.25 * self.df_f1.at[row_index, 'Количество студентов']
            if total_plan_hours_f2 >= 36:
                self.df_f1.at[row_index, 'Контр. работы'] *= 2

            self.df_f1.at[row_index, 'Зачеты'] = 0.25 * self.df_f1.at[row_index, 'Количество студентов']
            self.df_f1.at[row_index, 'Экзамены'] = 0.33 * self.df_f1.at[row_index, 'Количество студентов']
            self.df_f1.at[row_index, 'Практики'] = f2_row['Практики_F2_Наличие']
            self.df_f1.at[row_index, 'Курсовые работы'] = self.df_f1.at[row_index, 'Количество студентов'] * 2
            self.df_f1.at[row_index, 'Выпускные квалификационные работы'] = f2_row['ВКР_F2_Наличие']
            self.df_f1.at[row_index, 'Дополнительно'] = f2_row['Дополнительно_F2']

            # --- ЛОГИКА ЗАПОЛНЕНИЯ ИТОГОВЫХ СТОЛБЦОВ ---

            # Расчет часов личного присутствия (дн.)
            total_lich = (
                    total_lectures + total_practice + total_labs +
                    ksr_value +
                    self.df_f1.at[row_index, 'Консультации'] +
                    self.df_f1.at[row_index, 'Контр. работы'] +
                    self.df_f1.at[row_index, 'Зачеты'] +
                    self.df_f1.at[row_index, 'Экзамены'] +
                    self.df_f1.at[row_index, 'Практики'] +
                    self.df_f1.at[row_index, 'Курсовые работы'] +
                    self.df_f1.at[row_index, 'Выпускные квалификационные работы'] +
                    self.df_f1.at[row_index, 'Дополнительно']
            ).round(2)

            # Расчет часов в дистанционном формате (ЭОР)
            total_eor = (
                    self.df_f1.at[row_index, 'Лекции: в дистанционном формате (ЭОР)'] +
                    self.df_f1.at[row_index, 'Практические занятия: в дистанционном формате (ЭОР)'] +
                    self.df_f1.at[row_index, 'Лабораторные занятия: в дистанционном формате (ЭОР)']
            ).round(2)

            # Заполнение новых столбцов
            self.df_f1.at[row_index, 'дн.'] = total_lich
            self.df_f1.at[row_index, 'веч.'] = 0.00
            self.df_f1.at[row_index, 'заоч.'] = 0.00
            self.df_f1.at[row_index, 'в дистанционном формате (ЭОР)'] = total_eor
            self.df_f1.at[row_index, 'Всего'] = (total_lich + total_eor).round(2)

            # --- КОНЕЦ ЛОГИКИ ---

        except Exception as e:
            messagebox.showerror("Ошибка пересчета", f"Произошла ошибка при пересчете данных: {e}")
            traceback.print_exc()

    def process_file(self):
        """
        Основная логика обработки файла, вызываемая кнопкой.
        """
        file_path = filedialog.askopenfilename(
            title="Выберите файл F2-2022.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            self.status_label.config(text="Обработка отменена.")
            return

        self.status_label.config(text="Чтение и обработка файла...")
        self.root.update_idletasks()

        try:
            df_f2_raw = pd.read_excel(
                file_path,
                skiprows=header_row_index_level1 - 3,
                header=[0, 1]
            )
            df_f2_raw_columns = ['_'.join(col).strip() for col in df_f2_raw.columns.values]
            df_f2_raw.columns = df_f2_raw_columns

            column_mapping_f2_to_internal = {
                'Unnamed: 0_level_0_Unnamed: 0_level_1': 'Наименование_Дисциплины',
                'Unnamed: 1_level_0_Unnamed: 1_level_1': 'Шифр_Направления',
                'Unnamed: 2_level_0_Unnamed: 2_level_1': 'Наименование_Направления_Специализации',
                'Unnamed: 3_level_0_Unnamed: 3_level_1': 'Курс_F2',
                'Unnamed: 4_level_0_Unnamed: 4_level_1': 'Семестр_F2',
                'Unnamed: 5_level_0_студентов': 'Количество_Студентов_F2',
                'Unnamed: 6_level_0_потоков': 'Количество_Потоков_F2',
                'Unnamed: 7_level_0_групп': 'Группы_Подгруппы_F2_Raw',
                'Unnamed: 8_level_0_по плану': 'Лекции_По_Плану_F2',
                'Unnamed: 9_level_0_в дистанц. формате (ЭОР)': 'Лекции_Дистанц_F2',
                'Unnamed: 10_level_0_всего': 'Лекции_Всего_F2',
                'Unnamed: 11_level_0_по плану': 'Практика_По_Плану_F2',
                'Unnamed: 12_level_0_в дистанц. формате (ЭОР)': 'Практика_Дистанц_F2',
                'Unnamed: 13_level_0_всего': 'Практика_Всего_F2',
                'Unnamed: 14_level_0_по плану': 'Лабы_По_Плану_F2',
                'Unnamed: 15_level_0_в дистанц. формате (ЭОР)': 'Лабы_Дистанц_F2',
                'Unnamed: 16_level_0_всего': 'Лабы_Всего_F2',
                'Unnamed: 17_level_0_всего': 'КСР_F2',
                'Unnamed: 18_level_0_всего': 'Консультации_F2_Наличие',
                'Unnamed: 19_level_0_всего': 'Контр_Работы_F2_Наличие',
                'Unnamed: 20_level_0_всего': 'Зачеты_F2_Наличие',
                'Unnamed: 21_level_0_всего': 'Экзамены_F2_Наличие',
                'Unnamed: 22_level_0_всего': 'Практики_F2_Наличие',
                'Unnamed: 23_level_0_всего': 'Курсовые_F2_Наличие',
                'Unnamed: 24_level_0_всего': 'ВКР_F2_Наличие',
                'Unnamed: 25_level_0_всего': 'Дополнительно_F2',
                'Unnamed: 26_level_0_всего': 'ВСЕГО_F2_Часы',
                'Unnamed: 27_level_0_всего': 'Имена_Raw'
            }

            df_f2 = df_f2_raw.rename(columns=column_mapping_f2_to_internal, errors='ignore')
            df_f2 = df_f2.dropna(how='all')
            df_f2 = df_f2[
                ~df_f2['Наименование_Дисциплины'].astype(str).str.contains(r'^[IVX]+\.\s+', na=False, regex=True)]

            df_f2[['Количество_Групп_Число', 'Количество_Подгрупп_Число']] = df_f2['Группы_Подгруппы_F2_Raw'].apply(
                lambda x: pd.Series(parse_groups_subgroups(x)))

            numeric_cols = [
                'Курс_F2', 'Семестр_F2', 'Количество_Студентов_F2', 'Количество_Потоков_F2',
                'Лекции_По_Плану_F2', 'Лекции_Дистанц_F2', 'Практика_По_Плану_F2', 'Практика_Дистанц_F2',
                'Лабы_По_Плану_F2', 'Лабы_Дистанц_F2',
                'КСР_F2',
                'Консультации_F2_Наличие', 'Контр_Работы_F2_Наличие', 'Зачеты_F2_Наличие',
                'Экзамены_F2_Наличие', 'Практики_F2_Наличие', 'Курсовые_F2_Наличие', 'ВКР_F2_Наличие',
                'Дополнительно_F2', 'ВСЕГО_F2_Часы',
                'Количество_Групп_Число', 'Количество_Подгрупп_Число'
            ]
            for col in numeric_cols:
                if col in df_f2.columns:
                    df_f2[col] = pd.to_numeric(df_f2[col], errors='coerce').fillna(0)
            df_f2 = df_f2.dropna(subset=['Наименование_Дисциплины']).infer_objects(copy=False)

            speciality_map = {
                '010304.62': 'Прикладная математика и информатика',
                '09.03.01': 'Информатика и вычислительная техника',
                '010304.63': 'Прикладная математика, Прикладная математика и информатика',
            }
            df_f2['Наименование_Направления_Специализации_Mapped'] = df_f2['Шифр_Направления'].map(
                speciality_map).fillna(
                df_f2['Наименование_Направления_Специализации']).infer_objects(copy=False)

            # 30 столбцов (индексы 0-29)
            f1_columns_flat = [
                'Цикл дисциплины по уч. плану', 'Наименование дисциплины',
                'Шифр направления/специальности', 'Наименование направления/специальности (профиль/специализация)',
                'Форма обучения', 'Курс',
                'Количество студентов', 'Количество потоков', 'Количество групп', 'Количество подгрупп',
                'Лекции: всего', 'Лекции: в дистанционном формате (ЭОР)',
                'Практические занятия: всего', 'Практические занятия: в дистанционном формате (ЭОР)',
                'Лабораторные занятия: всего', 'Лабораторные занятия: в дистанционном формате (ЭОР)',
                'КСР',
                'Консультации', 'Контр. работы', 'Зачеты', 'Экзамены', 'Практики',
                'Курсовые работы', 'Выпускные квалификационные работы', 'Дополнительно',
                'Всего',
                'дн.', 'веч.', 'заоч.', 'в дистанционном формате (ЭОР)'
            ]
            self.df_f1 = pd.DataFrame(columns=f1_columns_flat)

            self.df_f1['Цикл дисциплины по уч. плану'] = df_f2['Наименование_Дисциплины'].astype(str).str.extract(
                r'^(Б\d+\.[О|В|Ф|П|Э|И].\d+(?:\.\d+)?)')
            self.df_f1['Наименование дисциплины'] = df_f2['Наименование_Дисциплины'].astype(str).str.replace(
                r'^(Б\d+\.[О|В|Ф|П|Э|И].\d+(?:\.\d+)?)\s*', '', regex=True)
            self.df_f1['Шифр направления/специальности'] = df_f2['Шифр_Направления']
            self.df_f1['Наименование направления/специальности (профиль/специализация)'] = df_f2[
                'Наименование_Направления_Специализации_Mapped']
            self.df_f1['Форма обучения'] = 'очная'
            self.df_f1['Курс'] = df_f2['Курс_F2']
            self.df_f1['Количество студентов'] = df_f2['Количество_Студентов_F2']
            self.df_f1['Количество потоков'] = df_f2['Количество_Потоков_F2']
            self.df_f1['Количество групп'] = df_f2['Количество_Групп_Число']
            self.df_f1['Количество подгрупп'] = df_f2['Количество_Подгрупп_Число']
            self.df_f1['Лекции: всего'] = df_f2['Лекции_По_Плану_F2'] * df_f2['Количество_Потоков_F2']
            self.df_f1['Лекции: в дистанционном формате (ЭОР)'] = df_f2['Лекции_Дистанц_F2']
            self.df_f1['Практические занятия: всего'] = df_f2['Практика_По_Плану_F2'] * self.df_f1['Количество групп']
            self.df_f1['Практические занятия: в дистанционном формате (ЭОР)'] = df_f2['Практика_Дистанц_F2']
            self.df_f1['Лабораторные занятия: всего'] = df_f2['Лабы_По_Плану_F2'] * self.df_f1['Количество подгрупп']
            self.df_f1['Лабораторные занятия: в дистанционном формате (ЭОР)'] = df_f2['Лабы_Дистанц_F2']

            self.df_f1['КСР'] = df_f2['КСР_F2']

            self.df_f2_processed = df_f2

            try:
                has_exam = self.df_f2_processed['Экзамены_F2_Наличие'] > 0
                consultations_base = (0.05 * self.df_f2_processed['Лекции_По_Плану_F2'])
                consultations_exam_bonus = has_exam.astype(int) * 2
                self.df_f1['Консультации'] = (consultations_base + consultations_exam_bonus) * self.df_f1[
                    'Количество групп']

                total_plan_hours_f2 = self.df_f2_processed['Лекции_По_Плану_F2'] + self.df_f2_processed[
                    'Практика_По_Плану_F2'] + self.df_f2_processed[
                                          'Лабы_По_Плану_F2']
                self.df_f1['Контр. работы'] = 0.25 * self.df_f1['Количество студентов']
                self.df_f1.loc[total_plan_hours_f2 >= 36, 'Контр. работы'] *= 2
                self.df_f1['Зачеты'] = 0.25 * self.df_f1['Количество студентов']
                self.df_f1['Экзамены'] = 0.33 * self.df_f1['Количество студентов']
                self.df_f1['Практики'] = self.df_f2_processed['Практики_F2_Наличие']
                self.df_f1['Курсовые работы'] = self.df_f1['Количество студентов'] * 2
                self.df_f1['Выпускные квалификационные работы'] = self.df_f2_processed['ВКР_F2_Наличие']
                self.df_f1['Дополнительно'] = self.df_f2_processed['Дополнительно_F2']

            except KeyError as ke:
                messagebox.showerror("Ошибка в данных", f"Отсутствует ключевой столбец в файле: {ke}")
                traceback.print_exc()
                return

            # --- БЛОК РАСЧЕТА ИТОГОВЫХ СТОЛБЦОВ ---

            total_lich = (
                    self.df_f1['Лекции: всего'] + self.df_f1['Практические занятия: всего'] +
                    self.df_f1['Лабораторные занятия: всего'] + self.df_f1['КСР'] +
                    self.df_f1['Консультации'] + self.df_f1['Контр. работы'] +
                    self.df_f1['Зачеты'] + self.df_f1['Экзамены'] + self.df_f1['Практики'] +
                    self.df_f1['Курсовые работы'] + self.df_f1['Выпускные квалификационные работы'] +
                    self.df_f1['Дополнительно']
            ).round(2)

            total_eor = (
                    self.df_f1['Лекции: в дистанционном формате (ЭОР)'] +
                    self.df_f1['Практические занятия: в дистанционном формате (ЭОР)'] +
                    self.df_f1['Лабораторные занятия: в дистанционном формате (ЭОР)']
            ).round(2)

            self.df_f1['дн.'] = total_lich
            self.df_f1['веч.'] = 0.00
            self.df_f1['заоч.'] = 0.00
            self.df_f1['в дистанционном формате (ЭОР)'] = total_eor
            self.df_f1['Всего'] = (total_lich + total_eor).round(2)

            # --- КОНЕЦ БЛОКА РАСЧЕТА ---

            self.df_f1 = self.df_f1.infer_objects(copy=False)
            self.df_f1['Имена_Raw'] = self.df_f2_processed['Имена_Raw']

            self.update_table(self.df_f1)
            self.status_label.config(
                text="Обработка завершена. Данные загружены. Дважды кликните по ячейке, чтобы изменить ее.")
            self.btn_save_report.config(state="normal")

        except FileNotFoundError:
            self.status_label.config(text="Ошибка: Файл не найден.")
            messagebox.showerror("Ошибка", "Файл F2-2022.xlsx не найден.")
        except Exception as e:
            self.status_label.config(text=f"Критическая ошибка: {e}", foreground="red")
            messagebox.showerror("Критическая ошибка", f"Произошла ошибка при обработке файла: {e}")
            traceback.print_exc()

    def update_table(self, df):
        """
        Заполняет таблицу данными из DataFrame.
        """
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, row in df.iterrows():
            values_to_show = (
                i,
                row['Цикл дисциплины по уч. плану'],
                row['Наименование дисциплины'],
                row['Шифр направления/специальности'],
                row['Курс'],
                row['Количество студентов'],
                row['Количество потоков'],
                row['Количество групп'],
                row['Количество подгрупп'],
                row['Лекции: всего'],
                row['Практические занятия: всего'],
                row['Лабораторные занятия: всего'],
                row['КСР'],
                row['Дополнительно'],  # Дополнительно
                row['Всего'],
                row['дн.'],
                row['в дистанционном формате (ЭОР)'],
                row['Имена_Raw']
            )
            self.tree.insert('', 'end', values=values_to_show)

    def save_report(self):
        """
        Сохраняет отчёт с отдельными листами для каждого имени.
        """
        if self.df_f1 is None:
            messagebox.showerror("Ошибка", "Сначала необходимо обработать файл.")
            return

        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="F1-2022_generated.xlsx"
        )
        if not output_file:
            return

        self.status_label.config(text="Сохранение отчёта...")
        self.root.update_idletasks()

        try:
            unique_names = get_unique_names(self.df_f1['Имена_Raw'].astype(str))
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book

                for name in unique_names:
                    pattern = re.compile(r'\b' + re.escape(name) + r'\b', re.IGNORECASE)
                    output_columns = [
                        'Цикл дисциплины по уч. плану', 'Наименование дисциплины',
                        'Шифр направления/специальности',
                        'Наименование направления/специальности (профиль/специализация)',
                        'Форма обучения', 'Курс',
                        'Количество студентов', 'Количество потоков', 'Количество групп', 'Количество подгрупп',
                        'Лекции: всего', 'Лекции: в дистанционном формате (ЭОР)',
                        'Практические занятия: всего', 'Практические занятия: в дистанционном формате (ЭОР)',
                        'Лабораторные занятия: всего', 'Лабораторные занятия: в дистанционном формате (ЭОР)',
                        'КСР', 'Консультации', 'Контр. работы', 'Зачеты', 'Экзамены', 'Практики',
                        'Курсовые работы', 'Выпускные квалификационные работы', 'Дополнительно',
                        'Всего',
                        'дн.', 'веч.', 'заоч.', 'в дистанционном формате (ЭОР)'
                    ]

                    filtered_df = self.df_f1[self.df_f1['Имена_Raw'].astype(str).str.contains(pattern, na=False)][
                        output_columns]

                    if not filtered_df.empty:
                        sheet_name = f'{name}'
                        if len(sheet_name) > 31:
                            sheet_name = sheet_name[:31]

                        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, header=False)

                        worksheet = writer.sheets[sheet_name]
                        self.format_worksheet(worksheet, workbook)

            self.status_label.config(text=f"Отчёт '{os.path.basename(output_file)}' успешно сохранён.",
                                     foreground="green")
            messagebox.showinfo("Готово", f"Отчёт '{os.path.basename(output_file)}' успешно сохранён.")

        except Exception as e:
            self.status_label.config(text=f"Критическая ошибка при сохранении: {e}", foreground="red")
            messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить отчёт: {e}")
            traceback.print_exc()

    def format_worksheet(self, worksheet, workbook):
        """
        Применяет форматирование заголовков и столбцов на листе Excel.
        """
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'border': 1
        })

        vertical_header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'rotation': 90, 'border': 1
        })

        merges = [
            ('A1:F1', 'Общие данные'),
            ('G1:J1', 'количество'),
            ('K1:L1', 'лекции'),
            ('M1:N1', 'практич. занятия'),
            ('O1:P1', 'лаборат. занятия'),
            # Q1:X1 - 8 столбцов. Y (24) - Дополнительно. Z (25) - Всего. AA-AD (26-29) - В том числе.
            ('Q1:X1', 'Распределение учебной нагрузки (в часах)'),
            ('Y1:Y2', 'Дополнительно'),  # FIX 1: Y1:Y2 для Дополнительно
            ('Z1:Z2', 'Всего'),  # FIX 2: Z1:Z2 для Всего
            ('AA1:AD1', 'В том числе')  # FIX 3: AA1:AD1 для В том числе
        ]
        for merge_range, title in merges:
            worksheet.merge_range(merge_range, title, header_format)

        vertical_columns_indices = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29]

        for i, header_text in enumerate(level1_headers):
            col_index = i

            if col_index in vertical_columns_indices:
                worksheet.write(1, col_index, header_text, vertical_header_format)
                worksheet.set_column(col_index, col_index, 5)
            else:
                worksheet.write(1, col_index, header_text, header_format)
                if col_index in [0, 1, 3]:
                    worksheet.set_column(col_index, col_index, 20)
                else:
                    worksheet.set_column(col_index, col_index, 10)

        # Устанавливаем высоту строк заголовков
        worksheet.set_row(0, 40)
        worksheet.set_row(1, 100)


if __name__ == "__main__":
    try:
        import pandas
        import xlsxwriter
    except ImportError as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Ошибка",
                             f"Необходимая библиотека не установлена: {e}\n\nПожалуйста, установите ее, выполнив в терминале:\npip install pandas xlsxwriter")
        root.destroy()
        exit()

    root = tk.Tk()
    app = App(root)
    root.mainloop()