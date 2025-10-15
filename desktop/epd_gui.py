#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ЕПД Парсер с GUI - Программа для обработки Единых Платежных Документов
Графический интерфейс для выбора полей для суммирования
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import PyPDF2
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple
import json


class EPDParser:
    """Класс для парсинга данных из ЕПД"""

    def __init__(self):
        self.data = {
            'период': None,
            'лицевой_счет': None,
            'адрес': None,
            'фио': None,
            'итого_к_оплате': None,
            'итого_к_оплате_без_страхования': None,
            'жилищные_услуги': [],
            'коммунальные_услуги': [],
            'страхование': None,
            'суммы_по_категориям': {}
        }

    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Извлекает текст из PDF файла"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ''
                for page in pdf_reader.pages:
                    text += page.extract_text()
                return text
        except Exception as e:
            raise Exception(f"Ошибка при чтении PDF: {e}")

    def parse_amount(self, text: str) -> float:
        """Преобразует строку с суммой в число (поддержка чисел с пробелами как разделителями тысяч)"""
        # Убираем все кроме цифр, запятой, точки и пробелов
        clean_text = re.sub(r'[^\d,.\s]', '', text)
        # Убираем пробелы (используются как разделители тысяч, например "1 927,72")
        clean_text = clean_text.replace(' ', '')
        # Заменяем запятую на точку
        clean_text = clean_text.replace(',', '.')
        try:
            return float(clean_text)
        except ValueError:
            return 0.0

    def parse_header_info(self, text: str):
        """Извлекает основную информацию из шапки документа"""
        period_match = re.search(r'ЗА\s+(\w+\s+\d{4})', text, re.IGNORECASE)
        if period_match:
            self.data['период'] = period_match.group(1)

        account_match = re.search(r'Лицевой счет:\s*(\d+[\s-]*\d+)', text, re.IGNORECASE)
        if account_match:
            self.data['лицевой_счет'] = account_match.group(1).strip()

        fio_match = re.search(r'ФИО:\s*([А-ЯЁ\s]+)', text)
        if fio_match:
            self.data['фио'] = fio_match.group(1).strip()

        address_match = re.search(r'Адрес:\s*(.+?)(?:\d+\s*руб|ИТОГО)', text, re.DOTALL | re.IGNORECASE)
        if address_match:
            addr = address_match.group(1).strip()
            addr = ' '.join(addr.split())
            # Убираем лишние символы
            addr = re.sub(r'\s+', ' ', addr)
            self.data['адрес'] = addr

        # Ищем итоговые суммы - улучшенные паттерны
        # Ищем строку с "6 201 руб. 04 коп." и "ИТОГО К ОПЛАТЕ ЗА ВСЕ УСЛУГИ СЧЕТА БЕЗ"
        total_no_ins_match = re.search(r'(\d[\s\d]*руб\.\s*\d+\s*коп\.)[\s\S]{0,100}?ИТОГО К ОПЛАТЕ ЗА ВСЕ УСЛУГИ.*?БЕЗ', text, re.IGNORECASE)
        if total_no_ins_match:
            self.data['итого_к_оплате_без_страхования'] = self.parse_amount(total_no_ins_match.group(1))

        total_with_ins_match = re.search(r'(\d[\s\d]*руб\.\s*\d+\s*коп\.)[\s\S]{0,100}?ИТОГО К ОПЛАТЕ ЗА ВСЕ УСЛУГИ.*?С\s+УЧЕТОМ', text, re.IGNORECASE)
        if total_with_ins_match:
            self.data['итого_к_оплате'] = self.parse_amount(total_with_ins_match.group(1))

        # Альтернативный способ - ищем строки "Итого к оплате"
        if not self.data['итого_к_оплате_без_страхования']:
            alt_match = re.search(r'Итого к оплате.*?без.*?(\d+[,\.]\d{2})', text, re.IGNORECASE)
            if alt_match:
                self.data['итого_к_оплате_без_страхования'] = self.parse_amount(alt_match.group(1))

    def parse_services(self, text: str):
        """Парсит услуги из таблицы расчетов"""
        # Улучшенный парсинг - ищем строки с конкретными паттернами
        # Паттерн для строк таблицы: Название Объем Ед.изм Тариф ... Итого

        current_section = None

        # Разбиваем текст на строки
        lines = text.split('\n')

        for i, line in enumerate(lines):
            line_stripped = line.strip()

            # Определяем секции
            if 'Начисления за жилищные услуги' in line:
                current_section = 'жилищные_услуги'
                continue
            elif 'Начисления за коммунальные услуги' in line:
                current_section = 'коммунальные_услуги'
                continue
            elif 'ДОБРОВОЛЬНОЕ СТРАХОВАНИЕ' in line:
                # Ищем все числа в строке
                numbers = re.findall(r'\d+[,\.]\d{2}', line)
                if numbers:
                    # Последнее число - итоговая сумма
                    self.data['страхование'] = self.parse_amount(numbers[-1])
                current_section = None
                continue
            elif 'Всего за' in line or 'Итого к оплате' in line.lower():
                current_section = None
                continue

            if not current_section:
                continue

            # Пропускаем пустые строки и заголовки
            if not line_stripped or any(keyword in line_stripped.lower() for keyword in
                                       ['виды услуг', 'объем услуг', 'начислено по тарифу']):
                continue

            # Ищем строки с услугами - должны содержать название и цифры
            # Используем более гибкий паттерн
            if re.search(r'[А-ЯЁ]', line) and re.search(r'\d+[,\.]\d{2}', line):
                try:
                    # Извлекаем все числа из строки
                    numbers = re.findall(r'\d+[,\.]\d+', line)
                    if not numbers:
                        continue

                    # Последнее число - итоговая сумма
                    total = self.parse_amount(numbers[-1])
                    if total <= 0:
                        continue

                    # Извлекаем название услуги (все до первого числа)
                    name_match = re.match(r'^([А-ЯЁа-яё\s\(\)/]+?)(?:\s+\d)', line_stripped)
                    if not name_match:
                        # Пробуем альтернативный вариант
                        name_match = re.match(r'^([А-ЯЁа-яё\s\(\)/]+)', line_stripped)

                    if name_match:
                        service_name = name_match.group(1).strip()
                    else:
                        # Берем первые слова как название
                        words = line_stripped.split()
                        service_name = ' '.join([w for w in words if not re.search(r'\d', w)])[:50]

                    # Пытаемся извлечь объем, единицы измерения и тариф
                    volume = 0.0
                    unit = ''
                    tariff = 0.0

                    # Ищем единицы измерения
                    unit_match = re.search(r'(кв\.м\.|куб\.\s*м\.|к[вВ]т[\./]?ч|Гкал)', line_stripped)
                    if unit_match:
                        unit = unit_match.group(1)

                    # Если есть несколько чисел, пытаемся определить объем и тариф
                    if len(numbers) >= 3:
                        volume = self.parse_amount(numbers[0])
                        tariff = self.parse_amount(numbers[1])
                    elif len(numbers) >= 2:
                        volume = self.parse_amount(numbers[0])

                    service_data = {
                        'название': service_name,
                        'объем': volume,
                        'ед_изм': unit,
                        'тариф': tariff,
                        'итого': total
                    }

                    self.data[current_section].append(service_data)

                except Exception as e:
                    print(f"Ошибка парсинга строки: {line_stripped[:50]}... - {e}")
                    continue

    def calculate_totals(self):
        """Вычисляет итоговые суммы по категориям"""
        housing_total = sum(item['итого'] for item in self.data['жилищные_услуги'])
        self.data['суммы_по_категориям']['Жилищные услуги'] = housing_total

        utility_total = sum(item['итого'] for item in self.data['коммунальные_услуги'])
        self.data['суммы_по_категориям']['Коммунальные услуги'] = utility_total

        if self.data['страхование']:
            self.data['суммы_по_категориям']['Добровольное страхование'] = self.data['страхование']

        total = housing_total + utility_total
        if self.data['страхование']:
            total += self.data['страхование']

        self.data['суммы_по_категориям']['ИТОГО'] = total

    def parse_pdf(self, pdf_path: str) -> Dict:
        """Основной метод парсинга PDF файла"""
        text = self.extract_text_from_pdf(pdf_path)
        if not text:
            raise Exception("Не удалось извлечь текст из PDF")

        self.parse_header_info(text)
        self.parse_services(text)
        self.calculate_totals()

        return self.data


class EPDGuiApp:
    """Графический интерфейс для парсера ЕПД"""

    def __init__(self, root):
        self.root = root
        self.root.title("ЕПД Парсер - Анализ платежных документов")
        self.root.geometry("1200x800")

        self.parser = EPDParser()
        self.loaded_files = []
        self.parsed_data = []
        self.include_insurance = tk.BooleanVar(value=False)  # По умолчанию страхование не включено

        self.setup_ui()

    def setup_ui(self):
        """Создает элементы интерфейса"""
        # Верхняя панель с кнопками
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)

        ttk.Button(top_frame, text="📂 Загрузить PDF файлы", command=self.load_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="🔄 Обработать все", command=self.process_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="💾 Сохранить в Excel", command=self.save_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="❌ Очистить", command=self.clear_all).pack(side=tk.LEFT, padx=5)

        # Основной контейнер с разделением
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Левая панель - список файлов
        left_frame = ttk.LabelFrame(main_paned, text="Загруженные файлы", padding="5")
        main_paned.add(left_frame, weight=1)

        # Список файлов
        self.files_listbox = tk.Listbox(left_frame, height=10)
        self.files_listbox.pack(fill=tk.BOTH, expand=True)
        self.files_listbox.bind('<<ListboxSelect>>', self.on_file_select)

        # Правая панель с вкладками
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=3)

        # Вкладки
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Вкладка 1: Информация о документе
        self.info_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.info_frame, text="📄 Информация")
        self.setup_info_tab()

        # Вкладка 2: Коммунальные услуги (объединены все услуги)
        self.services_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.services_frame, text="💧 Коммунальные услуги")
        self.setup_services_tab(self.services_frame)

        # Вкладка 3: Итоги
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="📊 Итоги")
        self.setup_summary_tab()

        # Статус бар
        self.status_label = ttk.Label(self.root, text="Готов к работе", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_info_tab(self):
        """Настройка вкладки с информацией"""
        info_text_frame = ttk.Frame(self.info_frame)
        info_text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.info_text = scrolledtext.ScrolledText(info_text_frame, wrap=tk.WORD, height=20)
        self.info_text.pack(fill=tk.BOTH, expand=True)

    def setup_services_tab(self, parent):
        """Настройка вкладки с услугами и чекбоксами"""
        # Фрейм с кнопками управления
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(control_frame, text="✓ Выбрать все",
                  command=lambda: self.select_all_services(True)).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="✗ Снять все",
                  command=lambda: self.select_all_services(False)).pack(side=tk.LEFT, padx=2)

        # Разделитель
        ttk.Separator(control_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)

        # Чекбокс для включения страхования
        insurance_check = ttk.Checkbutton(
            control_frame,
            text="🛡️ Включить добровольное страхование",
            variable=self.include_insurance,
            command=self.update_summary
        )
        insurance_check.pack(side=tk.LEFT, padx=5)

        # Фрейм с таблицей услуг
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Создаем Treeview для отображения услуг
        columns = ('Выбрать', 'Период', 'Категория', 'Название', 'Объем', 'Ед.изм.', 'Тариф', 'Сумма')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)

        # Настраиваем колонки
        tree.heading('Выбрать', text='☑')
        tree.heading('Период', text='Период')
        tree.heading('Категория', text='Категория')
        tree.heading('Название', text='Название услуги')
        tree.heading('Объем', text='Объем')
        tree.heading('Ед.изм.', text='Ед.изм.')
        tree.heading('Тариф', text='Тариф')
        tree.heading('Сумма', text='Сумма, руб.')

        tree.column('Выбрать', width=50, anchor=tk.CENTER)
        tree.column('Период', width=100)
        tree.column('Категория', width=120)
        tree.column('Название', width=220)
        tree.column('Объем', width=80, anchor=tk.E)
        tree.column('Ед.изм.', width=70)
        tree.column('Тариф', width=80, anchor=tk.E)
        tree.column('Сумма', width=100, anchor=tk.E)

        # Скроллбары
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(column=0, row=0, sticky='nsew')
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')

        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        # Обработка клика для переключения чекбокса
        tree.bind('<Button-1>', lambda e: self.toggle_checkbox(e, tree))

        # Сохраняем ссылку на дерево
        self.services_tree = tree
        self.services_checkboxes = {}

    def setup_summary_tab(self):
        """Настройка вкладки с итогами"""
        summary_text_frame = ttk.Frame(self.summary_frame)
        summary_text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.summary_text = scrolledtext.ScrolledText(summary_text_frame, wrap=tk.WORD, height=20, font=('Courier', 10))
        self.summary_text.pack(fill=tk.BOTH, expand=True)

    def load_files(self):
        """Загрузка PDF файлов"""
        files = filedialog.askopenfilenames(
            title="Выберите PDF файлы с ЕПД",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )

        if files:
            for file in files:
                if file not in self.loaded_files:
                    self.loaded_files.append(file)
                    filename = Path(file).name
                    self.files_listbox.insert(tk.END, filename)

            self.status_label.config(text=f"Загружено файлов: {len(self.loaded_files)}")

    def process_files(self):
        """Обработка всех загруженных файлов"""
        if not self.loaded_files:
            messagebox.showwarning("Предупреждение", "Сначала загрузите PDF файлы!")
            return

        self.parsed_data = []
        success_count = 0
        error_count = 0

        for file_path in self.loaded_files:
            try:
                self.status_label.config(text=f"Обработка: {Path(file_path).name}...")
                self.root.update()

                parser = EPDParser()
                data = parser.parse_pdf(file_path)
                data['file_path'] = file_path

                # Отладочная информация
                print(f"\n=== Обработан файл: {Path(file_path).name} ===")
                print(f"Период: {data.get('период')}")
                print(f"Лицевой счет: {data.get('лицевой_счет')}")
                print(f"Жилищных услуг: {len(data.get('жилищные_услуги', []))}")
                print(f"Коммунальных услуг: {len(data.get('коммунальные_услуги', []))}")
                print(f"Итого без страхования: {data.get('итого_к_оплате_без_страхования')}")

                self.parsed_data.append(data)
                success_count += 1

            except Exception as e:
                error_count += 1
                print(f"Ошибка при обработке {file_path}: {e}")
                import traceback
                traceback.print_exc()

        self.status_label.config(text=f"Обработано: {success_count} успешно, {error_count} с ошибками")

        # Обновляем отображение
        if self.parsed_data:
            self.display_all_data()

            # Выводим детальную информацию
            total_services = sum(len(d.get('жилищные_услуги', [])) + len(d.get('коммунальные_услуги', []))
                               for d in self.parsed_data)
            messagebox.showinfo("Успех",
                              f"Обработано документов: {success_count}\n"
                              f"Найдено услуг: {total_services}\n\n"
                              f"Проверьте консоль для подробной информации.")

    def display_all_data(self):
        """Отображение всех данных в таблицах"""
        # Очищаем таблицу
        for item in self.services_tree.get_children():
            self.services_tree.delete(item)

        self.services_checkboxes = {}

        # Заполняем таблицу данными из всех документов
        for data in self.parsed_data:
            period = data.get('период', 'Н/Д')

            # Жилищные услуги
            for service in data.get('жилищные_услуги', []):
                item_id = self.services_tree.insert('', 'end', values=(
                    '☑',  # По умолчанию выбрано
                    period,
                    'Жилищные',
                    service['название'],
                    f"{service['объем']:.2f}",
                    service['ед_изм'],
                    f"{service['тариф']:.2f}",
                    f"{service['итого']:.2f}"
                ))
                self.services_checkboxes[item_id] = {'checked': True, 'data': service, 'category': 'housing'}

            # Коммунальные услуги
            for service in data.get('коммунальные_услуги', []):
                item_id = self.services_tree.insert('', 'end', values=(
                    '☑',  # По умолчанию выбрано
                    period,
                    'Коммунальные',
                    service['название'],
                    f"{service['объем']:.2f}",
                    service['ед_изм'],
                    f"{service['тариф']:.2f}",
                    f"{service['итого']:.2f}"
                ))
                self.services_checkboxes[item_id] = {'checked': True, 'data': service, 'category': 'utility'}

        # Обновляем итоги
        self.update_summary()

        # Автоматически показываем информацию о первом файле
        if self.parsed_data:
            self.display_file_info(self.parsed_data[0])

    def toggle_checkbox(self, event, tree):
        """Переключение чекбокса при клике"""
        region = tree.identify("region", event.x, event.y)
        if region == "cell":
            column = tree.identify_column(event.x)
            if column == '#1':  # Колонка с чекбоксом
                item = tree.identify_row(event.y)
                if item and item in self.services_checkboxes:
                    self.services_checkboxes[item]['checked'] = not self.services_checkboxes[item]['checked']
                    symbol = '☑' if self.services_checkboxes[item]['checked'] else '☐'
                    values = list(tree.item(item, 'values'))
                    values[0] = symbol
                    tree.item(item, values=values)
                    self.update_summary()

    def select_all_services(self, select):
        """Выбрать/снять все услуги"""
        symbol = '☑' if select else '☐'

        for item in self.services_checkboxes:
            self.services_checkboxes[item]['checked'] = select
            values = list(self.services_tree.item(item, 'values'))
            values[0] = symbol
            self.services_tree.item(item, values=values)

        self.update_summary()

    def on_file_select(self, event):
        """Обработка выбора файла из списка"""
        selection = self.files_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        if index < len(self.parsed_data):
            data = self.parsed_data[index]
            self.display_file_info(data)

    def display_file_info(self, data):
        """Отображение информации о выбранном файле"""
        self.info_text.delete('1.0', tk.END)

        # Безопасное форматирование чисел
        total_no_insurance = data.get('итого_к_оплате_без_страхования') or 0.0
        total_with_insurance = data.get('итого_к_оплате') or 0.0

        info = f"""
╔══════════════════════════════════════════════════════════════╗
║              ИНФОРМАЦИЯ О ПЛАТЕЖНОМ ДОКУМЕНТЕ                 ║
╚══════════════════════════════════════════════════════════════╝

📅 Период: {data.get('период', 'Н/Д')}
🏠 Лицевой счет: {data.get('лицевой_счет', 'Н/Д')}
👤 ФИО: {data.get('фио', 'Н/Д')}
📍 Адрес: {data.get('адрес', 'Н/Д')}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💰 СУММЫ К ОПЛАТЕ:

   Без страхования: {total_no_insurance:.2f} руб.
   С учетом страхования: {total_with_insurance:.2f} руб.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🏠 ЖИЛИЩНЫЕ УСЛУГИ:
"""

        for service in data.get('жилищные_услуги', []):
            info += f"\n   • {service['название']}: {service['итого']:.2f} руб."

        info += "\n\n💧 КОММУНАЛЬНЫЕ УСЛУГИ:\n"

        for service in data.get('коммунальные_услуги', []):
            info += f"\n   • {service['название']}: {service['итого']:.2f} руб."

        if data.get('страхование'):
            info += f"\n\n🛡️ ДОБРОВОЛЬНОЕ СТРАХОВАНИЕ: {data.get('страхование'):.2f} руб."

        self.info_text.insert('1.0', info)

    def update_summary(self):
        """Обновление итоговых сумм"""
        housing_total = 0.0
        utility_total = 0.0
        insurance_total = 0.0

        # Суммируем выбранные услуги
        for item_id, item_data in self.services_checkboxes.items():
            if item_data['checked']:
                if item_data['category'] == 'housing':
                    housing_total += item_data['data']['итого']
                else:
                    utility_total += item_data['data']['итого']

        # Суммируем страхование из всех документов (если галочка включена)
        if self.include_insurance.get():
            for data in self.parsed_data:
                if data.get('страхование'):
                    insurance_total += data['страхование']

        grand_total = housing_total + utility_total + insurance_total

        # Формируем текст итогов
        summary = f"""
╔══════════════════════════════════════════════════════════════╗
║                    ИТОГОВАЯ СВОДКА                           ║
╚══════════════════════════════════════════════════════════════╝

📊 Обработано документов: {len(self.parsed_data)}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💰 СУММЫ ПО ВЫБРАННЫМ УСЛУГАМ:

   🏠 Жилищные услуги:        {housing_total:>12.2f} руб.
   💧 Коммунальные услуги:    {utility_total:>12.2f} руб.
   🛡️  Добровольное страхование: {insurance_total:>12.2f} руб.

   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   📌 ИТОГО:                  {grand_total:>12.2f} руб.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📈 ДЕТАЛИЗАЦИЯ ПО ПЕРИОДАМ:

"""

        # Детализация по периодам
        periods_summary = {}
        for data in self.parsed_data:
            period = data.get('период', 'Н/Д')
            if period not in periods_summary:
                periods_summary[period] = {'housing': 0.0, 'utility': 0.0, 'insurance': 0.0}

            # Жилищные
            for service in data.get('жилищные_услуги', []):
                periods_summary[period]['housing'] += service['итого']

            # Коммунальные
            for service in data.get('коммунальные_услуги', []):
                periods_summary[period]['utility'] += service['итого']

            # Страхование
            if data.get('страхование'):
                periods_summary[period]['insurance'] += data['страхование']

        for period, sums in periods_summary.items():
            total = sums['housing'] + sums['utility'] + sums['insurance']
            summary += f"\n   📅 {period}:\n"
            summary += f"      Жилищные: {sums['housing']:.2f} руб.\n"
            summary += f"      Коммунальные: {sums['utility']:.2f} руб.\n"
            summary += f"      Страхование: {sums['insurance']:.2f} руб.\n"
            summary += f"      ➜ Итого: {total:.2f} руб.\n"

        self.summary_text.delete('1.0', tk.END)
        self.summary_text.insert('1.0', summary)

    def save_to_excel(self):
        """Сохранение данных в Excel"""
        if not self.parsed_data:
            messagebox.showwarning("Предупреждение", "Нет данных для сохранения!")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"EPD_Анализ_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not file_path:
            return

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Лист 1: Сводная таблица по периодам
                summary_data = []
                for data in self.parsed_data:
                    row = {
                        'Период': data.get('период', 'Н/Д'),
                        'Лицевой счет': data.get('лицевой_счет', 'Н/Д'),
                        'Жилищные услуги': sum(s['итого'] for s in data.get('жилищные_услуги', [])),
                        'Коммунальные услуги': sum(s['итого'] for s in data.get('коммунальные_услуги', [])),
                        'Страхование': data.get('страхование', 0.0) or 0.0,
                        'Итого': data.get('итого_к_оплате', 0.0) or 0.0
                    }
                    summary_data.append(row)

                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Сводка по периодам', index=False)

                # Лист 2: Выбранные услуги
                services_data = []
                for item_id, item_data in self.services_checkboxes.items():
                    if item_data['checked']:
                        service = item_data['data']
                        values = self.services_tree.item(item_id, 'values')
                        category = 'Жилищные' if item_data['category'] == 'housing' else 'Коммунальные'
                        services_data.append({
                            'Период': values[1],
                            'Категория': category,
                            'Услуга': service['название'],
                            'Объем': service['объем'],
                            'Ед.изм.': service['ед_изм'],
                            'Тариф': service['тариф'],
                            'Сумма': service['итого']
                        })

                if services_data:
                    services_df = pd.DataFrame(services_data)
                    services_df.to_excel(writer, sheet_name='Выбранные услуги', index=False)

            messagebox.showinfo("Успех", f"Данные успешно сохранены в:\n{file_path}")
            self.status_label.config(text=f"Сохранено в: {Path(file_path).name}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")

    def clear_all(self):
        """Очистка всех данных"""
        if messagebox.askyesno("Подтверждение", "Очистить все загруженные данные?"):
            self.loaded_files = []
            self.parsed_data = []
            self.files_listbox.delete(0, tk.END)
            self.info_text.delete('1.0', tk.END)
            self.summary_text.delete('1.0', tk.END)

            for item in self.services_tree.get_children():
                self.services_tree.delete(item)

            self.services_checkboxes = {}
            self.include_insurance.set(False)

            self.status_label.config(text="Данные очищены. Готов к работе.")


def main():
    """Запуск приложения"""
    root = tk.Tk()
    app = EPDGuiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
