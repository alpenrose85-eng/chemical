import streamlit as st
import pandas as pd
from docx import Document
import json
import os
from datetime import datetime
import io
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

class ChemicalAnalyzer:
    def __init__(self):
        self.load_standards()
        
    def load_standards(self):
        """Загрузка стандартов из предустановленных файлов"""
        self.standards = {
            "12Х1МФ": {
                "C": (0.08, 0.15),
                "Si": (0.17, 0.37),
                "Mn": (0.40, 0.70),
                "Cr": (0.90, 1.20),
                "Mo": (0.25, 0.35),
                "V": (0.15, 0.30),
                "Cu": (None, 0.30),
                "S": (None, 0.025),
                "P": (None, 0.030),
                "Ni": (None, 0.30),
                "source": "ТУ 14-3Р-55-2001"
            },
            "12Х18Н12Т": {
                "C": (None, 0.12),
                "Si": (None, 0.80),
                "Mn": (1.00, 2.00),
                "Cr": (17.00, 19.00),
                "Ni": (11.00, 13.00),
                "Ti": (None, 0.70),
                "Cu": (None, 0.30),
                "S": (None, 0.020),
                "P": (None, 0.035),
                "source": "ГОСТ 5632-2014"
            },
            "сталь 20": {
                "C": (0.17, 0.24),
                "Si": (0.17, 0.37),
                "Mn": (0.35, 0.65),
                "Cr": (None, 0.25),
                "Ni": (None, 0.25),
                "Cu": (None, 0.30),
                "P": (None, 0.030),
                "S": (None, 0.025),
                "source": "ГОСТ 1050-2013"
            }
        }
        
        # Загрузка пользовательских стандартов если есть
        if os.path.exists("user_standards.json"):
            with open("user_standards.json", "r", encoding="utf-8") as f:
                user_std = json.load(f)
                self.standards.update(user_std)
    
    def save_user_standards(self):
        """Сохранение пользовательских стандартов"""
        with open("user_standards.json", "w", encoding="utf-8") as f:
            json.dump({k: v for k, v in self.standards.items() 
                      if k not in ["12Х1МФ", "12Х18Н12Т", "сталь 20"]}, f, ensure_ascii=False)
    
    def parse_protocol_file(self, file_content):
        """Парсинг файла протокола"""
        try:
            doc = Document(io.BytesIO(file_content))
            samples = []
            current_sample = None
            
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                
                # Поиск названия образца
                if "Наименование образца:" in text:
                    sample_name = text.split("Наименование образца:")[1].strip()
                    current_sample = {
                        "name": sample_name,
                        "steel_grade": None,
                        "composition": {}
                    }
                    samples.append(current_sample)
                
                # Поиск марки стали
                elif "Химический состав металла образца соответствует марке стали:" in text:
                    if current_sample:
                        # Извлечение марки стали (убираем лишние символы)
                        grade_text = text.split("марке стали:")[1].strip()
                        # Удаляем возможные ** вокруг марки стали
                        grade_text = re.sub(r'\*+', '', grade_text).strip()
                        current_sample["steel_grade"] = grade_text
            
            # Парсинг таблиц с химическим составом
            for i, table in enumerate(doc.tables):
                if i < len(samples):
                    composition = self.parse_composition_table(table)
                    samples[i]["composition"] = composition
            
            return samples
            
        except Exception as e:
            st.error(f"Ошибка при парсинге файла: {str(e)}")
            return []
    
    def parse_composition_table(self, table):
        """Парсинг таблицы с химическим составом"""
        composition = {}
        
        try:
            # Поиск строки со средними значениями
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                # Ищем строку, начинающуюся с "Среднее:"
                if cells and "Среднее:" in cells[0]:
                    # Сопоставляем элементы со значениями
                    elements = []
                    values = []
                    
                    # Собираем заголовки элементов из предыдущих строк
                    for prev_row in table.rows:
                        prev_cells = [cell.text.strip() for cell in prev_row.cells]
                        if prev_cells and prev_cells[0] in ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                                                           "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]:
                            elements = prev_cells
                            break
                    
                    if elements:
                        # Берем значения из текущей строки (пропускаем первый столбец "Среднее:")
                        values = cells[1:len(elements)]
                        
                        for elem, val in zip(elements, values):
                            if elem in ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                                      "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]:
                                try:
                                    # Заменяем точку на запятую и преобразуем в float
                                    num_val = float(val.replace(',', '.'))
                                    composition[elem] = num_val
                                except ValueError:
                                    continue
            
            return composition
            
        except Exception as e:
            st.error(f"Ошибка при парсинге таблицы: {str(e)}")
            return {}
    
    def check_compliance(self, sample):
        """Проверка соответствия нормативам"""
        if not sample["steel_grade"] or sample["steel_grade"] not in self.standards:
            return None
        
        standard = self.standards[sample["steel_grade"]]
        deviations = []
        borderlines = []
        
        for element, (min_val, max_val) in standard.items():
            if element == "source":
                continue
                
            if element in sample["composition"]:
                actual_val = sample["composition"][element]
                
                # Проверка соответствия
                if min_val is not None and actual_val < min_val:
                    deviations.append(f"{element}: {actual_val:.3f} < {min_val:.3f}")
                elif max_val is not None and actual_val > max_val:
                    deviations.append(f"{element}: {actual_val:.3f} > {max_val:.3f}")
                elif min_val is not None and actual_val <= min_val * 1.05:
                    borderlines.append(f"{element}: {actual_val:.3f} близко к мин. {min_val:.3f}")
                elif max_val is not None and actual_val >= max_val * 0.95:
                    borderlines.append(f"{element}: {actual_val:.3f} близко к макс. {max_val:.3f}")
        
        return {
            "deviations": deviations,
            "borderlines": borderlines,
            "is_compliant": len(deviations) == 0
        }
    
    def create_report_table(self, samples):
        """Создание сводной таблицы для отчета"""
        if not samples:
            return None
        
        # Собираем все уникальные марки стали
        steel_grades = list(set(sample["steel_grade"] for sample in samples if sample["steel_grade"]))
        
        tables = {}
        
        for grade in steel_grades:
            grade_samples = [s for s in samples if s["steel_grade"] == grade]
            
            if grade not in self.standards:
                continue
                
            standard = self.standards[grade]
            # Только нормируемые элементы
            norm_elements = [elem for elem in standard.keys() if elem != "source"]
            
            # Создаем DataFrame
            data = []
            for sample in grade_samples:
                row = {"Образец": sample["name"]}
                for elem in norm_elements:
                    if elem in sample["composition"]:
                        # Округление согласно требованиям
                        if elem in ["S", "P"]:
                            row[elem] = f"{sample['composition'][elem]:.3f}".replace('.', ',')
                        else:
                            row[elem] = f"{sample['composition'][elem]:.2f}".replace('.', ',')
                    else:
                        row[elem] = "-"
                data.append(row)
            
            # Добавляем строку с нормативами
            requirements_row = {"Образец": f"Требования {standard.get('source', '')} для стали марки {grade}"}
            for elem in norm_elements:
                min_val, max_val = standard[elem]
                if min_val is not None and max_val is not None:
                    requirements_row[elem] = f"{min_val:.2f}-{max_val:.2f}".replace('.', ',')
                elif min_val is not None:
                    requirements_row[elem] = f">={min_val:.2f}".replace('.', ',')
                elif max_val is not None:
                    requirements_row[elem] = f"<={max_val:.2f}".replace('.', ',')
                else:
                    requirements_row[elem] = "не нормируется"
            
            data.append(requirements_row)
            
            tables[grade] = pd.DataFrame(data)
        
        return tables

def main():
    st.set_page_config(page_title="Анализатор химсостава металла", layout="wide")
    st.title("🔬 Анализатор химического состава металла")
    
    analyzer = ChemicalAnalyzer()
    
    # Сайдбар для управления нормативами
    with st.sidebar:
        st.header("Управление нормативами")
        
        # Добавление нового норматива
        st.subheader("Добавить новую марку стали")
        new_grade = st.text_input("Марка стали")
        new_source = st.text_input("Нормативный документ")
        
        if new_grade:
            st.write("Укажите диапазоны для элементов (оставьте пустым если не нормируется):")
            col1, col2 = st.columns(2)
            elements_ranges = {}
            
            with col1:
                for elem in ["C", "Si", "Mn", "Cr", "Ni", "Mo"]:
                    min_val = st.number_input(f"{elem} мин", value=0.0, format="%.3f", key=f"min_{elem}")
                    max_val = st.number_input(f"{elem} макс", value=0.0, format="%.3f", key=f"max_{elem}")
                    if min_val > 0 or max_val > 0:
                        elements_ranges[elem] = (min_val if min_val > 0 else None, 
                                               max_val if max_val > 0 else None)
            
            with col2:
                for elem in ["V", "Cu", "S", "P", "Ti"]:
                    min_val = st.number_input(f"{elem} мин", value=0.0, format="%.3f", key=f"min_{elem}2")
                    max_val = st.number_input(f"{elem} макс", value=0.0, format="%.3f", key=f"max_{elem}2")
                    if min_val > 0 or max_val > 0:
                        elements_ranges[elem] = (min_val if min_val > 0 else None, 
                                               max_val if max_val > 0 else None)
            
            if st.button("Сохранить норматив") and new_grade:
                analyzer.standards[new_grade] = elements_ranges
                analyzer.standards[new_grade]["source"] = new_source
                analyzer.save_user_standards()
                st.success(f"Норматив для {new_grade} сохранен!")
    
    # Основная область для загрузки файлов
    st.header("Загрузка протоколов")
    
    uploaded_files = st.file_uploader(
        "Выберите файлы протоколов (.docx)", 
        type=["docx"], 
        accept_multiple_files=True
    )
    
    all_samples = []
    
    if uploaded_files:
        for uploaded_file in uploaded_files:
            st.write(f"**Обработка файла:** {uploaded_file.name}")
            
            samples = analyzer.parse_protocol_file(uploaded_file.getvalue())
            all_samples.extend(samples)
            
            for sample in samples:
                st.write(f"- Образец: {sample['name']}, Марка стали: {sample['steel_grade']}")
        
        # Анализ и отображение результатов
        if all_samples:
            st.header("Результаты анализа")
            
            # Создание таблиц для отчета
            report_tables = analyzer.create_report_table(all_samples)
            
            for grade, table in report_tables.items():
                st.subheader(f"Марка стали: {grade}")
                
                # Отображение таблицы в Streamlit
                st.dataframe(table)
                
                # Детальный анализ
                st.write("**Детальный анализ:**")
                grade_samples = [s for s in all_samples if s["steel_grade"] == grade]
                
                for sample in grade_samples:
                    compliance = analyzer.check_compliance(sample)
                    if compliance:
                        if compliance["is_compliant"]:
                            st.success(f"✅ {sample['name']} - Соответствует нормам")
                        else:
                            st.error(f"❌ {sample['name']} - Не соответствует нормам")
                            
                        if compliance["deviations"]:
                            st.write("Отклонения:")
                            for dev in compliance["deviations"]:
                                st.write(f"  - {dev}")
                        
                        if compliance["borderlines"]:
                            st.warning("Пограничные значения:")
                            for border in compliance["borderlines"]:
                                st.write(f"  - {border}")
            
            # Экспорт в Word
            if st.button("📄 Экспорт в Word"):
                create_word_report(report_tables, all_samples, analyzer)
                st.success("Отчет готов к скачиванию!")

def create_word_report(tables, samples, analyzer):
    """Создание Word отчета"""
    doc = Document()
    
    # Титульная страница
    title = doc.add_heading('Протокол анализа химического состава', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    doc.add_paragraph(f"Проанализировано образцов: {len(samples)}")
    doc.add_paragraph("")
    
    # Добавляем таблицы для каждой марки стали
    for grade, table_df in tables.items():
        doc.add_heading(f'Марка стали: {grade}', level=1)
        
        # Создаем таблицу в Word
        word_table = doc.add_table(rows=len(table_df)+1, cols=len(table_df.columns))
        word_table.style = 'Table Grid'
        
        # Заголовки
        for j, col in enumerate(table_df.columns):
            word_table.cell(0, j).text = str(col)
        
        # Данные
        for i, row in table_df.iterrows():
            for j, col in enumerate(table_df.columns):
                word_table.cell(i+1, j).text = str(row[col])
        
        doc.add_paragraph()
    
    # Сохраняем документ
    doc.save("химический_анализ_отчет.docx")
    st.success("Отчет сохранен как 'химический_анализ_отчет.docx'")

if __name__ == "__main__":
    main()
