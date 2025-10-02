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
            },
            "Ди82": {
                "C": (0.08, 0.12),
                "Si": (None, 0.5),
                "Mn": (0.30, 0.60),
                "Cr": (8.60, 10.00),
                "Ni": (None, 0.70),
                "Mo": (0.60, 0.80),
                "V": (0.10, 0.20),
                "Nb": (0.10, 0.20),
                "Cu": (None, 0.30),
                "S": (None, 0.015),
                "P": (None, 0.03),
                "source": "Спецификация"
            },
            "Ди59": {
                "C": (0.06, 0.10),
                "Si": (1.8, 2.2),
                "Mn": (12.00, 13.50),
                "Cr": (11.50, 13.00),
                "Ni": (1.8, 2.5),
                "Nb": (0.60, 1.00),
                "Cu": (2.00, 2.50),
                "S": (None, 0.02),
                "P": (None, 0.03),
                "source": "Спецификация"
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
            # Сохраняем только пользовательские стандарты (не предустановленные)
            predefined = ["12Х1МФ", "12Х18Н12Т", "сталь 20", "Ди82", "Ди59"]
            user_standards = {k: v for k, v in self.standards.items() if k not in predefined}
            json.dump(user_standards, f, ensure_ascii=False, indent=2)
    
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
                
                # Поиск марки стали - улучшенная версия
                elif "Химический состав металла образца соответствует марке стали:" in text:
                    if current_sample:
                        # Извлечение марки стали (убираем лишние символы и комментарии)
                        grade_text = text.split("марке стали:")[1].strip()
                        # Удаляем возможные ** вокруг марки стали и все что после запятой
                        grade_text = re.sub(r'\*+', '', grade_text).strip()
                        # Берем только первую часть до запятой (основную марку стали)
                        grade_text = grade_text.split(',')[0].strip()
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
        """Парсинг таблицы с химическим составом - по фиксированным строкам"""
        composition = {}
        
        try:
            # Преобразуем таблицу в список строк
            table_data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)
            
            # Проверяем, что таблица имеет достаточно строк
            if len(table_data) < 13:
                st.warning(f"Таблица имеет только {len(table_data)} строк, ожидалось минимум 13")
                return composition
            
            # Строка 0: заголовки первой группы элементов
            headers_row1 = table_data[0]
            # Строка 5: средние значения первой группы
            values_row1 = table_data[5]
            
            # Строка 7: заголовки второй группы элементов  
            headers_row2 = table_data[7]
            # Строка 12: средние значения второй группы
            values_row2 = table_data[12]
            
            # Все возможные элементы для поиска
            all_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                           "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]
            
            # Обрабатываем первую группу элементов
            for i, header in enumerate(headers_row1):
                if header in all_elements and i < len(values_row1):
                    value_str = values_row1[i]
                    try:
                        # Очищаем и преобразуем значение
                        value_str = value_str.replace(',', '.').replace(' ', '')
                        if '±' in value_str:
                            value_str = value_str.split('±')[0]
                        value = float(value_str)
                        composition[header] = value
                    except (ValueError, IndexError):
                        continue
            
            # Обрабатываем вторую группу элементов
            for i, header in enumerate(headers_row2):
                if header in all_elements and i < len(values_row2):
                    value_str = values_row2[i]
                    try:
                        # Очищаем и преобразуем значение
                        value_str = value_str.replace(',', '.').replace(' ', '')
                        if '±' in value_str:
                            value_str = value_str.split('±')[0]
                        value = float(value_str)
                        composition[header] = value
                    except (ValueError, IndexError):
                        continue
            
            return composition
            
        except Exception as e:
            st.error(f"Ошибка при парсинге таблицы: {str(e)}")
            return {}
    
    def check_element_compliance(self, element, value, standard):
        """Проверка соответствия элемента нормативам"""
        if element not in standard or element == "source":
            return "normal"
        
        min_val, max_val = standard[element]
        
        # Проверка соответствия
        if min_val is not None and value < min_val:
            return "deviation"
        elif max_val is not None and value > max_val:
            return "deviation"
        elif min_val is not None and value <= min_val * 1.05:
            return "borderline"
        elif max_val is not None and value >= max_val * 0.95:
            return "borderline"
        else:
            return "normal"
    
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
                st.warning(f"Нет нормативов для марки стали: {grade}")
                continue
                
            standard = self.standards[grade]
            # Только нормируемые элементы (исключаем 'source')
            norm_elements = [elem for elem in standard.keys() if elem != "source"]
            
            # Создаем DataFrame
            data = []
            compliance_data = []  # Для хранения информации о соответствии
            
            for sample in grade_samples:
                row = {"Образец": sample["name"]}
                compliance_row = {"Образец": "normal"}  # Статус для названия образца
                
                for elem in norm_elements:
                    if elem in sample["composition"]:
                        value = sample["composition"][elem]
                        # Округление согласно требованиям
                        if elem in ["S", "P"]:
                            row[elem] = f"{value:.3f}".replace('.', ',')
                        else:
                            row[elem] = f"{value:.2f}".replace('.', ',')
                        
                        # Проверяем соответствие
                        status = self.check_element_compliance(elem, value, standard)
                        compliance_row[elem] = status
                    else:
                        row[elem] = "-"
                        compliance_row[elem] = "normal"
                
                data.append(row)
                compliance_data.append(compliance_row)
            
            # Добавляем строку с нормативами
            requirements_row = {"Образец": f"Требования {standard.get('source', '')} для стали марки {grade}"}
            requirements_compliance = {"Образец": "requirements"}
            
            for elem in norm_elements:
                min_val, max_val = standard[elem]
                if min_val is not None and max_val is not None:
                    if elem in ["S", "P"]:
                        requirements_row[elem] = f"{min_val:.3f}-{max_val:.3f}".replace('.', ',')
                    else:
                        requirements_row[elem] = f"{min_val:.2f}-{max_val:.2f}".replace('.', ',')
                elif min_val is not None:
                    if elem in ["S", "P"]:
                        requirements_row[elem] = f">={min_val:.3f}".replace('.', ',')
                    else:
                        requirements_row[elem] = f">={min_val:.2f}".replace('.', ',')
                elif max_val is not None:
                    if elem in ["S", "P"]:
                        requirements_row[elem] = f"<={max_val:.3f}".replace('.', ',')
                    else:
                        requirements_row[elem] = f"<={max_val:.2f}".replace('.', ',')
                else:
                    requirements_row[elem] = "не нормируется"
                
                requirements_compliance[elem] = "requirements"
            
            data.append(requirements_row)
            compliance_data.append(requirements_compliance)
            
            tables[grade] = {
                "data": pd.DataFrame(data),
                "compliance": compliance_data
            }
        
        return tables

def apply_styling(df, compliance_data):
    """Применяет стили к DataFrame на основе данных о соответствии"""
    styled_df = df.copy()
    
    # CSS стили для разных статусов
    styles = []
    for i, row in df.iterrows():
        if i < len(compliance_data):
            compliance_row = compliance_data[i]
            for col in df.columns:
                if col in compliance_row:
                    status = compliance_row[col]
                    if status == "deviation":
                        styles.append(f"background-color: #ffcccc; color: #cc0000; font-weight: bold;")  # Красный
                    elif status == "borderline":
                        styles.append(f"background-color: #fffacd; color: #b8860b;")  # Желтый
                    elif status == "requirements":
                        styles.append(f"background-color: #f0f0f0; font-style: italic;")  # Серый для требований
                    else:
                        styles.append("")  # Нормальный стиль
                else:
                    styles.append("")
    
    # Применяем стили
    styled = df.style
    for i in range(len(df)):
        for j, col in enumerate(df.columns):
            idx = i * len(df.columns) + j
            if idx < len(styles) and styles[idx]:
                styled = styled.set_properties(subset=(i, col), **{'css': styles[idx]})
    
    return styled

def main():
    st.set_page_config(page_title="Анализатор химсостава металла", layout="wide")
    st.title("🔬 Анализатор химического состава металла")
    
    analyzer = ChemicalAnalyzer()
    
    # Сайдбар для управления нормативами
    with st.sidebar:
        st.header("Управление нормативами")
        
        # Просмотр существующих нормативов
        st.subheader("Существующие марки стали")
        selected_standard = st.selectbox(
            "Выберите марку для просмотра",
            options=list(analyzer.standards.keys())
        )
        
        if selected_standard:
            st.write(f"**Норматив для {selected_standard}:**")
            standard = analyzer.standards[selected_standard]
            for elem, value_range in standard.items():
                # Пропускаем поле 'source'
                if elem == "source":
                    continue
                
                # Проверяем, что это действительно диапазон значений
                if isinstance(value_range, tuple) and len(value_range) == 2:
                    min_val, max_val = value_range
                    if min_val is not None and max_val is not None:
                        st.write(f"- {elem}: {min_val:.3f} - {max_val:.3f}")
                    elif min_val is not None:
                        st.write(f"- {elem}: ≥ {min_val:.3f}")
                    elif max_val is not None:
                        st.write(f"- {elem}: ≤ {max_val:.3f}")
            st.write(f"Источник: {standard.get('source', 'не указан')}")
        
        st.divider()
        
        # Добавление новых нормативов
        st.subheader("Добавить новую марку стали")
        
        new_grade = st.text_input("Марка стали")
        new_source = st.text_input("Нормативный документ")
        
        if new_grade:
            st.write("**Добавление элементов:**")
            
            # Инициализация session_state для элементов
            if 'elements' not in st.session_state:
                st.session_state.elements = []
            
            # Поля для добавления нового элемента
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1:
                new_element = st.text_input("Элемент (например: Nb, W, B)", key="new_element")
            with col2:
                new_min = st.number_input("Мин. значение", value=0.0, format="%.3f", key="new_min")
            with col3:
                new_max = st.number_input("Макс. значение", value=0.0, format="%.3f", key="new_max")
            
            if st.button("Добавить элемент") and new_element:
                st.session_state.elements.append({
                    "element": new_element.strip().upper(),
                    "min": new_min if new_min > 0 else None,
                    "max": new_max if new_max > 0 else None
                })
            
            # Отображение добавленных элементов
            if st.session_state.elements:
                st.write("Добавленные элементы:")
                elements_to_remove = []
                
                for i, elem_data in enumerate(st.session_state.elements):
                    col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
                    with col1:
                        st.write(f"**{elem_data['element']}**")
                    with col2:
                        min_val = elem_data['min']
                        st.write(f"Мин: {min_val:.3f}" if min_val else "Мин: не норм.")
                    with col3:
                        max_val = elem_data['max']
                        st.write(f"Макс: {max_val:.3f}" if max_val else "Макс: не норм.")
                    with col4:
                        if st.button("❌", key=f"del_{i}"):
                            elements_to_remove.append(i)
                
                # Удаляем отмеченные элементы
                for i in sorted(elements_to_remove, reverse=True):
                    st.session_state.elements.pop(i)
            
            # Кнопка сохранения
            if st.button("💾 Сохранить норматив"):
                if not st.session_state.elements:
                    st.error("Добавьте хотя бы один элемент!")
                elif new_grade in analyzer.standards:
                    st.error(f"Марка стали {new_grade} уже существует!")
                else:
                    # Создаем словарь с элементами
                    elements_ranges = {}
                    for elem_data in st.session_state.elements:
                        elements_ranges[elem_data["element"]] = (
                            elem_data["min"], 
                            elem_data["max"]
                        )
                    
                    elements_ranges["source"] = new_source
                    analyzer.standards[new_grade] = elements_ranges
                    analyzer.save_user_standards()
                    
                    # Очищаем session state
                    st.session_state.elements = []
                    
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
            
            # Легенда цветов
            st.markdown("""
            **Легенда:**
            - <span style='background-color: #ffcccc; padding: 2px 5px; border-radius: 3px;'>🔴 Красный</span> - отклонение от норм
            - <span style='background-color: #fffacd; padding: 2px 5px; border-radius: 3px;'>🟡 Желтый</span> - пограничное значение
            - <span style='background-color: #f0f0f0; padding: 2px 5px; border-radius: 3px;'>⚪ Серый</span> - нормативные требования
            """, unsafe_allow_html=True)
            
            for grade, table_data in report_tables.items():
                st.subheader(f"Марка стали: {grade}")
                
                # Редактирование таблицы
                st.write("**Редактирование таблицы:**")
                edited_df = st.data_editor(
                    table_data["data"],
                    key=f"editor_{grade}",
                    num_rows="fixed",
                    use_container_width=True,
                    column_config={
                        "Образец": st.column_config.TextColumn(
                            "Образец",
                            help="Название образца",
                            required=True
                        )
                    }
                )
                
                # Применяем стили к отредактированной таблице
                styled_table = apply_styling(edited_df, table_data["compliance"])
                st.write("**Таблица с визуализацией отклонений:**")
                st.dataframe(styled_table, use_container_width=True)
            
            # Экспорт в Word
            if st.button("📄 Экспорт в Word"):
                # Используем отредактированные данные для экспорта
                edited_tables = {}
                for grade in report_tables.keys():
                    if f"editor_{grade}" in st.session_state:
                        edited_tables[grade] = st.session_state[f"editor_{grade}"]
                    else:
                        edited_tables[grade] = report_tables[grade]["data"]
                
                create_word_report(edited_tables, all_samples, analyzer)
                st.success("Отчет готов к скачиванию!")

def create_word_report(tables, samples, analyzer):
    """Создание Word отчета"""
    try:
        doc = Document()
        
        # Титульная страница
        title = doc.add_heading('Протокол анализа химического состава', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        doc.add_paragraph(f"Проанализировано образцов: {len(samples)}")
        doc.add_paragraph("")
        
        # Легенда
        doc.add_heading('Легенда', level=1)
        legend_table = doc.add_table(rows=4, cols=2)
        legend_table.style = 'Table Grid'
        
        legend_table.cell(0, 0).text = "Цвет"
        legend_table.cell(0, 1).text = "Значение"
        
        legend_table.cell(1, 0).text = "🔴"
        legend_table.cell(1, 1).text = "Отклонение от норм"
        
        legend_table.cell(2, 0).text = "🟡" 
        legend_table.cell(2, 1).text = "Пограничное значение"
        
        legend_table.cell(3, 0).text = "⚪"
        legend_table.cell(3, 1).text = "Нормативные требования"
        
        doc.add_paragraph()
        
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
        
        # Предоставляем ссылку для скачивания
        with open("химический_анализ_отчет.docx", "rb") as file:
            btn = st.download_button(
                label="📥 Скачать отчет",
                data=file,
                file_name="химический_анализ_отчет.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
    except Exception as e:
        st.error(f"Ошибка при создании Word отчета: {str(e)}")

if __name__ == "__main__":
    main()
