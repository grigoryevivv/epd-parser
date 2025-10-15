#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ЕПД Парсер Mobile - Мобильная версия для Android
Использует Kivy для создания кроссплатформенного интерфейса
"""

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.popup import Popup
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.checkbox import CheckBox
from kivy.uix.textinput import TextInput
from kivy.core.window import Window
from kivy.utils import platform

import re
import PyPDF2
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List
import os


class EPDParser:
    """Класс для парсинга данных из ЕПД (тот же что и в desktop версии)"""

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
        clean_text = re.sub(r'[^\d,.\s]', '', text)
        clean_text = clean_text.replace(' ', '')
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
            addr = re.sub(r'\s+', ' ', addr)
            self.data['адрес'] = addr

        total_no_ins_match = re.search(r'(\d[\s\d]*руб\.\s*\d+\s*коп\.)[\s\S]{0,100}?ИТОГО К ОПЛАТЕ ЗА ВСЕ УСЛУГИ.*?БЕЗ', text, re.IGNORECASE)
        if total_no_ins_match:
            self.data['итого_к_оплате_без_страхования'] = self.parse_amount(total_no_ins_match.group(1))

        total_with_ins_match = re.search(r'(\d[\s\d]*руб\.\s*\d+\s*коп\.)[\s\S]{0,100}?ИТОГО К ОПЛАТЕ ЗА ВСЕ УСЛУГИ.*?С\s+УЧЕТОМ', text, re.IGNORECASE)
        if total_with_ins_match:
            self.data['итого_к_оплате'] = self.parse_amount(total_with_ins_match.group(1))

        if not self.data['итого_к_оплате_без_страхования']:
            alt_match = re.search(r'Итого к оплате.*?без.*?(\d+[,\.]\d{2})', text, re.IGNORECASE)
            if alt_match:
                self.data['итого_к_оплате_без_страхования'] = self.parse_amount(alt_match.group(1))

    def parse_services(self, text: str):
        """Парсит услуги из таблицы расчетов"""
        current_section = None
        lines = text.split('\n')

        for i, line in enumerate(lines):
            line_stripped = line.strip()

            if 'Начисления за жилищные услуги' in line:
                current_section = 'жилищные_услуги'
                continue
            elif 'Начисления за коммунальные услуги' in line:
                current_section = 'коммунальные_услуги'
                continue
            elif 'ДОБРОВОЛЬНОЕ СТРАХОВАНИЕ' in line:
                numbers = re.findall(r'\d+[,\.]\d{2}', line)
                if numbers:
                    self.data['страхование'] = self.parse_amount(numbers[-1])
                current_section = None
                continue
            elif 'Всего за' in line or 'Итого к оплате' in line.lower():
                current_section = None
                continue

            if not current_section:
                continue

            if not line_stripped or any(keyword in line_stripped.lower() for keyword in
                                       ['виды услуг', 'объем услуг', 'начислено по тарифу']):
                continue

            if re.search(r'[А-ЯЁ]', line) and re.search(r'\d+[,\.]\d{2}', line):
                try:
                    numbers = re.findall(r'\d+[,\.]\d+', line)
                    if not numbers:
                        continue

                    total = self.parse_amount(numbers[-1])
                    if total <= 0:
                        continue

                    name_match = re.match(r'^([А-ЯЁа-яё\s\(\)/]+?)(?:\s+\d)', line_stripped)
                    if not name_match:
                        name_match = re.match(r'^([А-ЯЁа-яё\s\(\)/]+)', line_stripped)

                    if name_match:
                        service_name = name_match.group(1).strip()
                    else:
                        words = line_stripped.split()
                        service_name = ' '.join([w for w in words if not re.search(r'\d', w)])[:50]

                    volume = 0.0
                    unit = ''
                    tariff = 0.0

                    unit_match = re.search(r'(кв\.м\.|куб\.\s*м\.|к[вВ]т[\./]?ч|Гкал)', line_stripped)
                    if unit_match:
                        unit = unit_match.group(1)

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


class ServiceItem(BoxLayout):
    """Элемент списка услуг с чекбоксом"""
    def __init__(self, service_data, category, period, callback, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'horizontal'
        self.size_hint_y = None
        self.height = 40
        self.padding = 5
        self.spacing = 5

        self.service_data = service_data
        self.category = category
        self.callback = callback

        # Чекбокс
        self.checkbox = CheckBox(size_hint_x=0.1, active=True)
        self.checkbox.bind(active=self.on_checkbox)
        self.add_widget(self.checkbox)

        # Информация об услуге
        info_text = f"{service_data['название'][:30]}\n{service_data['итого']:.2f} руб."
        info_label = Label(
            text=info_text,
            size_hint_x=0.9,
            halign='left',
            valign='middle',
            text_size=(None, None)
        )
        info_label.bind(size=info_label.setter('text_size'))
        self.add_widget(info_label)

    def on_checkbox(self, instance, value):
        """Обработчик изменения чекбокса"""
        if self.callback:
            self.callback()


class EPDMobileApp(App):
    """Главное приложение"""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.parser = EPDParser()
        self.parsed_data = []
        self.service_items = []
        self.include_insurance = False

    def build(self):
        """Построение интерфейса"""
        Window.clearcolor = (0.95, 0.95, 0.95, 1)

        # Основной контейнер
        main_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        # Заголовок
        header = Label(
            text='ЕПД Парсер',
            size_hint_y=0.08,
            font_size='24sp',
            bold=True,
            color=(0.2, 0.4, 0.8, 1)
        )
        main_layout.add_widget(header)

        # Панель с кнопками
        button_panel = GridLayout(cols=2, size_hint_y=0.12, spacing=5)

        load_btn = Button(
            text='Загрузить PDF',
            on_press=self.load_file,
            background_color=(0.3, 0.6, 0.9, 1)
        )
        button_panel.add_widget(load_btn)

        export_btn = Button(
            text='Сохранить Excel',
            on_press=self.export_excel,
            background_color=(0.3, 0.8, 0.5, 1)
        )
        button_panel.add_widget(export_btn)

        main_layout.add_widget(button_panel)

        # Вкладки
        self.tabs = TabbedPanel(do_default_tab=False)

        # Вкладка: Информация
        info_tab = TabbedPanelItem(text='Информация')
        self.info_text = Label(
            text='Загрузите PDF файл для начала работы',
            halign='left',
            valign='top',
            padding=(10, 10)
        )
        info_scroll = ScrollView()
        info_scroll.add_widget(self.info_text)
        info_tab.add_widget(info_scroll)
        self.tabs.add_widget(info_tab)

        # Вкладка: Услуги
        services_tab = TabbedPanelItem(text='Услуги')
        services_layout = BoxLayout(orientation='vertical', spacing=5)

        # Чекбокс для страхования
        insurance_layout = BoxLayout(size_hint_y=0.1, spacing=10, padding=5)
        self.insurance_checkbox = CheckBox(size_hint_x=0.2)
        self.insurance_checkbox.bind(active=self.on_insurance_toggle)
        insurance_layout.add_widget(self.insurance_checkbox)
        insurance_layout.add_widget(Label(text='Включить страхование', size_hint_x=0.8))
        services_layout.add_widget(insurance_layout)

        # Список услуг
        self.services_container = GridLayout(cols=1, spacing=5, size_hint_y=None)
        self.services_container.bind(minimum_height=self.services_container.setter('height'))

        services_scroll = ScrollView()
        services_scroll.add_widget(self.services_container)
        services_layout.add_widget(services_scroll)

        services_tab.add_widget(services_layout)
        self.tabs.add_widget(services_tab)

        # Вкладка: Итоги
        summary_tab = TabbedPanelItem(text='Итоги')
        self.summary_text = Label(
            text='Итоги появятся после загрузки файла',
            halign='left',
            valign='top',
            padding=(10, 10)
        )
        summary_scroll = ScrollView()
        summary_scroll.add_widget(self.summary_text)
        summary_tab.add_widget(summary_scroll)
        self.tabs.add_widget(summary_tab)

        main_layout.add_widget(self.tabs)

        return main_layout

    def load_file(self, instance):
        """Загрузка PDF файла"""
        content = BoxLayout(orientation='vertical')

        # Определяем стартовую директорию в зависимости от платформы
        if platform == 'android':
            from android.storage import primary_external_storage_path
            start_path = primary_external_storage_path()
        else:
            start_path = str(Path.home())

        filechooser = FileChooserListView(
            path=start_path,
            filters=['*.pdf']
        )
        content.add_widget(filechooser)

        button_layout = BoxLayout(size_hint_y=0.15, spacing=10)

        select_btn = Button(text='Выбрать')
        cancel_btn = Button(text='Отмена')

        button_layout.add_widget(select_btn)
        button_layout.add_widget(cancel_btn)
        content.add_widget(button_layout)

        popup = Popup(title='Выберите PDF файл', content=content, size_hint=(0.9, 0.9))

        def on_select(btn):
            if filechooser.selection:
                self.process_file(filechooser.selection[0])
            popup.dismiss()

        select_btn.bind(on_press=on_select)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def process_file(self, file_path):
        """Обработка загруженного файла"""
        try:
            parser = EPDParser()
            data = parser.parse_pdf(file_path)
            data['file_path'] = file_path
            self.parsed_data = [data]

            self.display_info(data)
            self.display_services(data)
            self.update_summary()

            self.show_message('Успех', 'Файл успешно обработан!')

        except Exception as e:
            self.show_message('Ошибка', f'Не удалось обработать файл:\n{str(e)}')

    def display_info(self, data):
        """Отображение информации о документе"""
        total_no_insurance = data.get('итого_к_оплате_без_страхования') or 0.0
        total_with_insurance = data.get('итого_к_оплате') or 0.0

        info = f"""
ИНФОРМАЦИЯ О ДОКУМЕНТЕ

Период: {data.get('период', 'Н/Д')}
Лицевой счет: {data.get('лицевой_счет', 'Н/Д')}
ФИО: {data.get('фио', 'Н/Д')}

СУММЫ К ОПЛАТЕ:
Без страхования: {total_no_insurance:.2f} руб.
С учетом страхования: {total_with_insurance:.2f} руб.

ЖИЛИЩНЫЕ УСЛУГИ:
"""
        for service in data.get('жилищные_услуги', []):
            info += f"  • {service['название']}: {service['итого']:.2f} руб.\n"

        info += "\nКОММУНАЛЬНЫЕ УСЛУГИ:\n"
        for service in data.get('коммунальные_услуги', []):
            info += f"  • {service['название']}: {service['итого']:.2f} руб.\n"

        if data.get('страхование'):
            info += f"\nСТРАХОВАНИЕ: {data.get('страхование'):.2f} руб."

        self.info_text.text = info

    def display_services(self, data):
        """Отображение списка услуг"""
        self.services_container.clear_widgets()
        self.service_items = []

        period = data.get('период', 'Н/Д')

        # Добавляем жилищные услуги
        for service in data.get('жилищные_услуги', []):
            item = ServiceItem(service, 'housing', period, self.update_summary)
            self.services_container.add_widget(item)
            self.service_items.append(item)

        # Добавляем коммунальные услуги
        for service in data.get('коммунальные_услуги', []):
            item = ServiceItem(service, 'utility', period, self.update_summary)
            self.services_container.add_widget(item)
            self.service_items.append(item)

    def update_summary(self):
        """Обновление итогов"""
        housing_total = 0.0
        utility_total = 0.0
        insurance_total = 0.0

        # Суммируем выбранные услуги
        for item in self.service_items:
            if item.checkbox.active:
                if item.category == 'housing':
                    housing_total += item.service_data['итого']
                else:
                    utility_total += item.service_data['итого']

        # Добавляем страхование если включено
        if self.include_insurance:
            for data in self.parsed_data:
                if data.get('страхование'):
                    insurance_total += data['страхование']

        grand_total = housing_total + utility_total + insurance_total

        summary = f"""
ИТОГОВАЯ СВОДКА

Обработано документов: {len(self.parsed_data)}

СУММЫ ПО ВЫБРАННЫМ УСЛУГАМ:

Жилищные услуги:        {housing_total:>12.2f} руб.
Коммунальные услуги:    {utility_total:>12.2f} руб.
Добровольное страхование: {insurance_total:>12.2f} руб.

────────────────────────────────────
ИТОГО:                  {grand_total:>12.2f} руб.
"""
        self.summary_text.text = summary

    def on_insurance_toggle(self, checkbox, value):
        """Обработчик переключения страхования"""
        self.include_insurance = value
        self.update_summary()

    def export_excel(self, instance):
        """Экспорт в Excel"""
        if not self.parsed_data:
            self.show_message('Предупреждение', 'Сначала загрузите PDF файл!')
            return

        try:
            # Определяем путь для сохранения
            if platform == 'android':
                from android.storage import primary_external_storage_path
                output_dir = os.path.join(primary_external_storage_path(), 'Download')
            else:
                output_dir = str(Path.home() / 'Downloads')

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_dir, f"EPD_Анализ_{timestamp}.xlsx")

            # Создаем Excel файл
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Собираем данные по выбранным услугам
                services_data = []
                for item in self.service_items:
                    if item.checkbox.active:
                        service = item.service_data
                        category = 'Жилищные' if item.category == 'housing' else 'Коммунальные'
                        services_data.append({
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

            self.show_message('Успех', f'Файл сохранен:\n{output_file}')

        except Exception as e:
            self.show_message('Ошибка', f'Не удалось сохранить:\n{str(e)}')

    def show_message(self, title, message):
        """Показать всплывающее сообщение"""
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        content.add_widget(Label(text=message))

        btn = Button(text='OK', size_hint_y=0.3)
        content.add_widget(btn)

        popup = Popup(title=title, content=content, size_hint=(0.8, 0.4))
        btn.bind(on_press=popup.dismiss)
        popup.open()


if __name__ == '__main__':
    EPDMobileApp().run()
