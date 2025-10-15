#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ЕПД Парсер - Программа для обработки Единых Платежных Документов
Автор: Создано для обработки ЖКХ квитанций
"""

import re
import PyPDF2
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple


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
            print(f"Ошибка при чтении PDF: {e}")
            return ""

    def parse_amount(self, text: str) -> float:
        """Преобразует строку с суммой в число"""
        # Убираем все кроме цифр, запятой и точки
        clean_text = re.sub(r'[^\d,.]', '', text)
        # Заменяем запятую на точку
        clean_text = clean_text.replace(',', '.')
        try:
            return float(clean_text)
        except ValueError:
            return 0.0

    def parse_header_info(self, text: str):
        """Извлекает основную информацию из шапки документа"""
        # Период
        period_match = re.search(r'ЗА\s+(\w+\s+\d{4})', text)
        if period_match:
            self.data['период'] = period_match.group(1)

        # Лицевой счет
        account_match = re.search(r'Лицевой счет:\s*(\d+[\s-]*\d+)', text)
        if account_match:
            self.data['лицевой_счет'] = account_match.group(1).strip()

        # ФИО
        fio_match = re.search(r'ФИО:\s*([А-ЯЁ\s]+)', text)
        if fio_match:
            self.data['фио'] = fio_match.group(1).strip()

        # Адрес
        address_match = re.search(r'Адрес:\s*(.+?)(?:\n|ИТОГО)', text, re.DOTALL)
        if address_match:
            addr = address_match.group(1).strip()
            # Очищаем от лишних переносов
            addr = ' '.join(addr.split())
            self.data['адрес'] = addr

        # Итого к оплате
        total_match = re.search(r'ИТОГО К ОПЛАТЕ.*?(\d+\s*руб\.\s*\d+\s*коп\.)', text)
        if total_match:
            self.data['итого_к_оплате'] = total_match.group(1)

    def parse_services(self, text: str):
        """Парсит услуги из таблицы расчетов"""
        # Ищем секцию начислений за жилищные услуги
        housing_section = re.search(
            r'Начисления за жилищные услуги(.*?)Начисления за коммунальные услуги',
            text, re.DOTALL
        )

        if housing_section:
            self._parse_service_section(housing_section.group(1), 'жилищные_услуги')

        # Ищем секцию начислений за коммунальные услуги
        utility_section = re.search(
            r'Начисления за коммунальные услуги(.*?)Всего за',
            text, re.DOTALL
        )

        if utility_section:
            self._parse_service_section(utility_section.group(1), 'коммунальные_услуги')

        # Добровольное страхование
        insurance_match = re.search(
            r'ДОБРОВОЛЬНОЕ СТРАХОВАНИЕ.*?(\d+[,\.]\d{2}).*?(\d+[,\.]\d{2})\s*$',
            text, re.MULTILINE
        )
        if insurance_match:
            self.data['страхование'] = self.parse_amount(insurance_match.group(2))

    def _parse_service_section(self, section_text: str, category: str):
        """Парсит секцию услуг"""
        # Разбиваем на строки
        lines = section_text.strip().split('\n')

        for line in lines:
            # Пропускаем пустые строки
            if not line.strip():
                continue

            # Ищем строки с услугами (содержат название и итоговую сумму)
            # Паттерн: название услуги, объем, ед.изм, тариф, начислено, ..., итого
            service_match = re.search(
                r'^([А-ЯЁ\s/\(\)]+?)\s+(\d+[,.]?\d*)\s+([а-яё\.]+)\s+(\d+[,.]?\d*)\s+.*?(\d+[,\.]\d{2})\s*$',
                line
            )

            if service_match:
                service_name = service_match.group(1).strip()
                volume = self.parse_amount(service_match.group(2))
                unit = service_match.group(3).strip()
                tariff = self.parse_amount(service_match.group(4))
                total = self.parse_amount(service_match.group(5))

                service_data = {
                    'название': service_name,
                    'объем': volume,
                    'ед_изм': unit,
                    'тариф': tariff,
                    'итого': total
                }

                self.data[category].append(service_data)

    def calculate_totals(self):
        """Вычисляет итоговые суммы по категориям"""
        # Жилищные услуги
        housing_total = sum(item['итого'] for item in self.data['жилищные_услуги'])
        self.data['суммы_по_категориям']['Жилищные услуги'] = housing_total

        # Коммунальные услуги
        utility_total = sum(item['итого'] for item in self.data['коммунальные_услуги'])
        self.data['суммы_по_категориям']['Коммунальные услуги'] = utility_total

        # Страхование
        if self.data['страхование']:
            self.data['суммы_по_категориям']['Добровольное страхование'] = self.data['страхование']

        # Общий итог
        total = housing_total + utility_total
        if self.data['страхование']:
            total += self.data['страхование']

        self.data['суммы_по_категориям']['ИТОГО'] = total

    def parse_pdf(self, pdf_path: str) -> Dict:
        """Основной метод парсинга PDF файла"""
        print(f"Обработка файла: {pdf_path}")

        # Извлекаем текст
        text = self.extract_text_from_pdf(pdf_path)

        if not text:
            print("Не удалось извлечь текст из PDF")
            return None

        # Парсим данные
        self.parse_header_info(text)
        self.parse_services(text)
        self.calculate_totals()

        return self.data


class EPDAnalyzer:
    """Класс для анализа и суммирования данных из нескольких ЕПД"""

    def __init__(self):
        self.monthly_data = []

    def add_epd(self, epd_data: Dict):
        """Добавляет данные ЕПД в коллекцию"""
        if epd_data:
            self.monthly_data.append(epd_data)

    def create_summary_dataframe(self) -> pd.DataFrame:
        """Создает сводную таблицу по всем периодам"""
        summary_rows = []

        for epd in self.monthly_data:
            row = {
                'Период': epd.get('период', 'Н/Д'),
                'Лицевой счет': epd.get('лицевой_счет', 'Н/Д'),
                'ФИО': epd.get('фио', 'Н/Д'),
            }

            # Добавляем суммы по категориям
            for category, amount in epd.get('суммы_по_категориям', {}).items():
                row[category] = amount

            summary_rows.append(row)

        df = pd.DataFrame(summary_rows)
        return df

    def create_detailed_dataframe(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Создает детальные таблицы по жилищным и коммунальным услугам"""
        housing_rows = []
        utility_rows = []

        for epd in self.monthly_data:
            period = epd.get('период', 'Н/Д')

            # Жилищные услуги
            for service in epd.get('жилищные_услуги', []):
                housing_rows.append({
                    'Период': period,
                    'Услуга': service['название'],
                    'Объем': service['объем'],
                    'Ед. изм.': service['ед_изм'],
                    'Тариф': service['тариф'],
                    'Сумма': service['итого']
                })

            # Коммунальные услуги
            for service in epd.get('коммунальные_услуги', []):
                utility_rows.append({
                    'Период': period,
                    'Услуга': service['название'],
                    'Объем': service['объем'],
                    'Ед. изм.': service['ед_изм'],
                    'Тариф': service['тариф'],
                    'Сумма': service['итого']
                })

        housing_df = pd.DataFrame(housing_rows)
        utility_df = pd.DataFrame(utility_rows)

        return housing_df, utility_df

    def save_to_excel(self, output_file: str):
        """Сохраняет все данные в Excel файл"""
        print(f"\nСохранение результатов в файл: {output_file}")

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Сводная таблица
            summary_df = self.create_summary_dataframe()
            summary_df.to_excel(writer, sheet_name='Сводная таблица', index=False)

            # Детальные таблицы
            housing_df, utility_df = self.create_detailed_dataframe()

            if not housing_df.empty:
                housing_df.to_excel(writer, sheet_name='Жилищные услуги', index=False)

            if not utility_df.empty:
                utility_df.to_excel(writer, sheet_name='Коммунальные услуги', index=False)

            # Итоговая статистика
            if not summary_df.empty:
                stats_data = []
                numeric_columns = summary_df.select_dtypes(include=['float64', 'int64']).columns

                for col in numeric_columns:
                    if col in summary_df.columns:
                        stats_data.append({
                            'Категория': col,
                            'Сумма за все периоды': summary_df[col].sum(),
                            'Среднее за период': summary_df[col].mean(),
                            'Минимум': summary_df[col].min(),
                            'Максимум': summary_df[col].max()
                        })

                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='Статистика', index=False)

        print(f"✓ Файл успешно создан: {output_file}")


def main():
    """Основная функция программы"""
    print("=" * 60)
    print("    ЕПД ПАРСЕР - Обработка платежных документов ЖКХ")
    print("=" * 60)

    # Путь к папке с PDF файлами
    pdf_folder = Path(r"c:\Users\grigo\OneDrive\KU")

    # Создаем анализатор
    analyzer = EPDAnalyzer()

    # Ищем PDF файлы с ЕПД в названии
    pdf_files = list(pdf_folder.glob("ЕПД*.pdf"))

    if not pdf_files:
        print("\n⚠ Не найдено ни одного PDF файла с ЕПД в папке:")
        print(f"  {pdf_folder}")
        print("\nПоложите PDF файлы в эту папку и запустите программу снова.")
        return

    print(f"\nНайдено файлов: {len(pdf_files)}\n")

    # Обрабатываем каждый файл
    parser = EPDParser()

    for pdf_file in pdf_files:
        try:
            epd_data = parser.parse_pdf(str(pdf_file))
            if epd_data:
                analyzer.add_epd(epd_data)
                print(f"✓ Обработан: {pdf_file.name}")
                print(f"  Период: {epd_data.get('период', 'Н/Д')}")
                print(f"  Итого: {epd_data.get('суммы_по_категориям', {}).get('ИТОГО', 0):.2f} руб.\n")
        except Exception as e:
            print(f"✗ Ошибка при обработке {pdf_file.name}: {e}\n")

    # Сохраняем результаты
    if analyzer.monthly_data:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = pdf_folder / f"EPD_Анализ_{timestamp}.xlsx"
        analyzer.save_to_excel(str(output_file))

        print("\n" + "=" * 60)
        print("Обработка завершена!")
        print("=" * 60)
        print(f"\nВсего обработано документов: {len(analyzer.monthly_data)}")

        # Выводим итоговую статистику
        summary_df = analyzer.create_summary_dataframe()
        if not summary_df.empty:
            print("\nИТОГОВЫЕ СУММЫ:")
            numeric_columns = summary_df.select_dtypes(include=['float64', 'int64']).columns
            for col in numeric_columns:
                if col in summary_df.columns:
                    total = summary_df[col].sum()
                    print(f"  {col}: {total:,.2f} руб.")
    else:
        print("\n⚠ Не удалось обработать ни один документ")


if __name__ == "__main__":
    main()
