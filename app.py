import streamlit as st
import pandas as pd
from docx import Document
import json
import os
from datetime import datetime
import io
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
import re

class SampleNameMatcher:
    def __init__(self):
        self.surface_types = {
            'ЭПК': ['ЭПК'],
            'ШПП': ['ШПП'],
            'ПС КШ': ['ПС КШ', 'труба_ПТКМ', 'труба ПТКМ', 'ПТКМ', 'труба'],
            'КПП ВД': ['КПП ВД', 'ВД'],
            'КПП НД-1': ['КПП НД-1', 'НД-1'],
            'КПП НД-2': ['КПП НД-2', 'НД-2']
        }
        self.letters = ['А', 'Б', 'В', 'Г']
    
    def parse_correct_names(self, file_content):
        """Парсинг файла с правильными названиями образцов из таблицы"""
        try:
            doc = Document(io.BytesIO(file_content))
            correct_names = []
            
            # Парсим таблицы в документе
            for table in doc.tables:
                for row in table.rows:
                    if len(row.cells) >= 2:  # Как минимум 2 столбца: номер и название
                        number_cell = row.cells[0].text.strip()
                        name_cell = row.cells[1].text.strip()
                        
                        # Пропускаем пустые строки и заголовки
                        if number_cell and name_cell and number_cell.isdigit():
                            correct_names.append({
                                'number': int(number_cell),
                                'original': name_cell,
                                'surface_type': self.extract_surface_type(name_cell),
                                'tube_number': self.extract_tube_number(name_cell),
                                'letter': self.extract_letter(name_cell)
                            })
            
            # Если таблиц нет, пробуем парсить как обычный текст
            if not correct_names:
                for paragraph in doc.paragraphs:
                    text = paragraph.text.strip()
                    # Ищем строки с форматом "число   название"
                    match = re.match(r'^\s*(\d+)\s+([^\s].*)$', text)
                    if match:
                        number = match.group(1)
                        name = match.group(2).strip()
                        if number.isdigit():
                            correct_names.append({
                                'number': int(number),
                                'original': name,
                                'surface_type': self.extract_surface_type(name),
                                'tube_number': self.extract_tube_number(name),
                                'letter': self.extract_letter(name)
                            })
            
            return correct_names
        except Exception as e:
            st.error(f"Ошибка при парсинге файла с правильными названиями: {str(e)}")
            return []
    
    def extract_surface_type(self, name):
        """Извлечение типа поверхности нагрева из названия с учетом опечаток"""
        for surface_type, patterns in self.surface_types.items():
            for pattern in patterns:
                if pattern in name:
                    return surface_type
        return None
    
    def extract_tube_number(self, name):
        """Извлечение номера трубы из названия"""
        # Ищем числа в скобках или после названия типа
        matches = re.findall(r'\((\d+)[,-]', name)
        if matches:
            return matches[0]
        
        # Альтернативные паттерны
        matches = re.findall(r'(\d+)[,]\s*[А-Г]\)', name)
        if matches:
            return matches[0]
        
        # Для формата типа "ШПП (4-1,А)" - берем первое число
        matches = re.findall(r'(\d+)-\d+', name)
        if matches:
            return matches[0]
            
        return None
    
    def extract_letter(self, name):
        """Извлечение буквы (А, Б, В, Г) из названия"""
        for letter in self.letters:
            if f',{letter}' in name or f', {letter}' in name or f'({letter})' in name or f',{letter})' in name:
                return letter
        return None
    
    def parse_protocol_sample_name(self, sample_name):
        """Парсинг названия образца из протокола химического анализа"""
        # Определяем букву из префикса (НА, НБ, НВ, НГ)
        letter_map = {'НА': 'А', 'НБ': 'Б', 'НВ': 'В', 'НГ': 'Г', 'Н-Г': 'Г'}
        letter = None
        for prefix, mapped_letter in letter_map.items():
            if sample_name.startswith(prefix):
                letter = mapped_letter
                break
        
        # Определяем тип поверхности
        surface_type = None
        for stype, patterns in self.surface_types.items():
            for pattern in patterns:
                if pattern in sample_name:
                    surface_type = stype
                    break
            if surface_type:
                break
        
        # Извлекаем номер трубы
        tube_number = None
        # Ищем числа в названии
        numbers = re.findall(r'\d+', sample_name)
        if numbers:
            # Для ПС КШ берем первое число как номер трубы
            if surface_type == 'ПС КШ':
                tube_number = numbers[0]
            # Для других типов пытаемся найти номер после типа
            else:
                # Ищем паттерн "тип (число"
                pattern_match = re.search(r'(\d+)[_ ]', sample_name)
                if pattern_match:
                    tube_number = pattern_match.group(1)
                else:
                    # Берем первое найденное число как номер трубы
                    tube_number = numbers[0]
        
        return {
            'original': sample_name,
            'surface_type': surface_type,
            'tube_number': tube_number,
            'letter': letter
        }
    
    def find_best_match(self, protocol_sample, correct_samples):
        """Нахождение наилучшего соответствия для образца из протокола"""
        best_match = None
        best_score = 0
        
        for correct_sample in correct_samples:
            score = self.calculate_match_score(protocol_sample, correct_sample)
            if score > best_score:
                best_score = score
                best_match = correct_sample
        
        # Возвращаем совпадение только если score достаточно высок
        return best_match if best_score >= 2 else None
    
    def calculate_match_score(self, protocol_sample, correct_sample):
        """Вычисление оценки соответствия между образцами с улучшенной логикой"""
        score = 0
        
        # Совпадение типа поверхности (2 балла)
        if (protocol_sample['surface_type'] and 
            correct_sample['surface_type'] and 
            protocol_sample['surface_type'] == correct_sample['surface_type']):
            score += 2
        # Частичное совпадение типа (1 балл) - если один из типов None, но есть другие признаки
        elif (protocol_sample['surface_type'] is None or 
              correct_sample['surface_type'] is None):
            # Если тип не определен с одной стороны, но есть сильные другие признаки
            score += 0  # не даем баллов за неопределенность
        
        # Совпадение номера трубы (2 балла)
        if (protocol_sample['tube_number'] and 
            correct_sample['tube_number'] and 
            protocol_sample['tube_number'] == correct_sample['tube_number']):
            score += 2
        
        # Совпадение буквы (1 балл)
        if (protocol_sample['letter'] and 
            correct_sample['letter'] and 
            protocol_sample['letter'] == correct_sample['letter']):
            score += 1
        
        # ДОПОЛНИТЕЛЬНО: если номер труба и буква совпадают, но тип поверхности разный,
        # даем шанс на сопоставление (особенно для ПС КШ / труба_ПТКМ)
        if (protocol_sample['tube_number'] and correct_sample['tube_number'] and
            protocol_sample['letter'] and correct_sample['letter'] and
            protocol_sample['tube_number'] == correct_sample['tube_number'] and
            protocol_sample['letter'] == correct_sample['letter']):
            score += 1  # дополнительный балл за полное совпадение номера и буквы
        
        return score

    def _filter_correct_names(self, options, filter_text, correct_samples):
        """Фильтрация вариантов названий по номеру или букве"""
        if not filter_text:
            return options
        
        filter_text = filter_text.upper().strip()
        filtered_options = ["Не сопоставлен"]
        
        # Ищем совпадения
        for cs in correct_samples:
            # Поиск по номеру трубы
            if cs.get('tube_number') and filter_text in cs['tube_number']:
                filtered_options.append(cs['original'])
                continue
                
            # Поиск по номеру в списке
            if cs.get('number') and filter_text in str(cs['number']):
                filtered_options.append(cs['original'])
                continue
                
            # Поиск по букве
            if cs.get('letter') and filter_text == cs['letter']:
                filtered_options.append(cs['original'])
                continue
                
            # Поиск по названию (частичное совпадение)
            if filter_text in cs['original'].upper():
                filtered_options.append(cs['original'])
                continue
        
        # Удаляем дубликаты и сохраняем порядок
        seen = set()
        unique_options = []
        for option in filtered_options:
            if option not in seen:
                seen.add(option)
                unique_options.append(option)
        
        return unique_options if unique_options else ["Не сопоставлен"]

class ChemicalAnalyzer:
    def __init__(self):
        self.load_standards()
        self.name_matcher = SampleNameMatcher()
        self.debug_mode = False  # Флаг отладки
        
    def load_standards(self):
        """Загрузка стандартов из предустановленных файлов"""
        self.standards = {
            "12Х1МФ": {
                "C": (0.10, 0.15),
                "Si": (0.17, 0.37),
                "Mn": (0.40, 0.70),
                "Cr": (0.90, 1.20),
                "Mo": (0.25, 0.35),
                "V": (0.15, 0.30),
                "Ni": (None, 0.25),
                "Cu": (None, 0.20),
                "S": (None, 0.025),
                "P": (None, 0.025),
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
                "source": "ТУ 14-3Р-55-2001"
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
                "source": "ТУ 14-3Р-55-2001"
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
                "source": "ТУ 14-3Р-55-2001"
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
                "source": "ТУ 14-3Р-55-2001"
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
        """Парсинг файла протокола - УЛУЧШЕННАЯ ВЕРСИЯ"""
        try:
            doc = Document(io.BytesIO(file_content))
            samples = []
            current_sample = None
            
            # Собираем все параграфы
            paragraphs = [p for p in doc.paragraphs if p.text.strip()]
            
            i = 0
            while i < len(paragraphs):
                text = paragraphs[i].text.strip()
                
                # Поиск названия образца
                if "Наименование образца:" in text:
                    # Если есть текущий образец, сохраняем его
                    if current_sample:
                        samples.append(current_sample)
                    
                    sample_name = text.split("Наименование образца:")[1].strip()
                    current_sample = {
                        "name": sample_name,
                        "original_name": sample_name,
                        "steel_grade": None,
                        "composition": {}
                    }
                    i += 1
                    continue
                
                # Поиск марки стали - улучшенная версия
                if current_sample and not current_sample["steel_grade"]:
                    if "Химический состав металла образца соответствует марке стали:" in text:
                        grade_text = text.split("марке стали:")[1].strip()
                        # Очистка текста марки стали
                        grade_text = re.sub(r'\*+', '', grade_text).strip()
                        # Удаляем комментарии о допустимых отклонениях
                        if "," in grade_text:
                            grade_text = grade_text.split(",")[0].strip()
                        if "с учетом" in grade_text:
                            grade_text = grade_text.split("с учетом")[0].strip()
                        current_sample["steel_grade"] = grade_text
                    
                    elif "Химический состав металла образца близок к марке стали:" in text:
                        grade_text = text.split("марке стали:")[1].strip()
                        grade_text = re.sub(r'\*+', '', grade_text).strip()
                        if "," in grade_text:
                            grade_text = grade_text.split(",")[0].strip()
                        current_sample["steel_grade"] = grade_text
                
                i += 1
            
            # Добавляем последний образец
            if current_sample:
                samples.append(current_sample)
            
            # Парсинг таблиц с химическим составом
            for i, table in enumerate(doc.tables):
                if i < len(samples):
                    composition = self.parse_composition_table_corrected(table, sample_index=i)
                    samples[i]["composition"] = composition
            
            # Отладочная информация
            st.success(f"✅ Найдено образцов: {len(samples)}")
            
            # Показываем статистику по маркам стали
            grade_stats = {}
            for sample in samples:
                grade = sample.get("steel_grade", "Не распознана")
                grade_stats[grade] = grade_stats.get(grade, 0) + 1
            
            st.info("📊 Статистика по маркам стали:")
            for grade, count in grade_stats.items():
                st.write(f"  - {grade}: {count} образцов")
            
            return samples
            
        except Exception as e:
            st.error(f"Ошибка при парсинге файла: {str(e)}")
            import traceback
            st.error(f"Детали ошибки: {traceback.format_exc()}")
            return []
    
    def parse_composition_table_corrected(self, table, sample_index=0):
        """Правильный парсинг таблицы с химическим составом - УЛУЧШЕННАЯ ВЕРСИЯ С ОТЛАДКОЙ"""
        composition = {}
        
        try:
            # Собираем все данные из таблицы
            all_data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                # Фильтруем пустые строки
                if any(cell.strip() for cell in row_data):
                    all_data.append(row_data)
            
            # Если таблица пустая, возвращаем пустой состав
            if not all_data:
                return composition
            
            # РАСШИРЕННАЯ ОТЛАДКА: показываем полную структуру таблицы
            with st.expander(f"🔍 РАСШИРЕННАЯ ОТЛАДКА ТАБЛИЦЫ (образец {sample_index+1})", expanded=False):
                st.write("**Полная структура таблицы:**")
                
                # Создаем DataFrame для наглядного отображения
                debug_df_data = []
                for i, row in enumerate(all_data):
                    row_data = {"Строка": i}
                    for j, cell in enumerate(row):
                        row_data[f"Столбец {j}"] = cell
                    debug_df_data.append(row_data)
                
                if debug_df_data:
                    debug_df = pd.DataFrame(debug_df_data).fillna("")
                    st.dataframe(debug_df, use_container_width=True)
                
                # Анализ ячеек с числами
                st.write("**Анализ числовых значений:**")
                numeric_cells = []
                for i, row in enumerate(all_data):
                    for j, cell in enumerate(row):
                        if self._is_numeric_value(cell):
                            try:
                                value = self._parse_numeric_value(cell)
                                numeric_cells.append({
                                    "Строка": i,
                                    "Столбец": j,
                                    "Значение": cell,
                                    "Число": value
                                })
                            except:
                                pass
                
                if numeric_cells:
                    st.table(pd.DataFrame(numeric_cells))
                else:
                    st.write("Числовые значения не найдены")
                
                # Поиск заголовков элементов
                st.write("**Поиск химических элементов в заголовках:**")
                elements_found = []
                chemical_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                                   "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]
                
                for i, row in enumerate(all_data):
                    for j, cell in enumerate(row):
                        for element in chemical_elements:
                            if element in cell.upper():
                                elements_found.append({
                                    "Элемент": element,
                                    "Строка": i,
                                    "Столбец": j,
                                    "Текст": cell
                                })
                                break
                
                if elements_found:
                    st.table(pd.DataFrame(elements_found))
                else:
                    st.write("Химические элементы в заголовках не найдены")
            
            # Определяем структуру таблицы автоматически
            composition = self._auto_detect_table_structure(all_data, sample_index)
            
            # Если автоматическое определение не сработало, используем интерактивный режим
            if not composition and self.debug_mode:
                composition = self._interactive_table_parsing(all_data, sample_index)
            
            return composition
            
        except Exception as e:
            st.error(f"Ошибка при парсинге таблицы: {str(e)}")
            import traceback
            st.error(f"Детали ошибки: {traceback.format_exc()}")
            return {}

    def _auto_detect_table_structure(self, all_data, sample_index):
        """Автоматическое определение структуры таблицы"""
        composition = {}
        
        # СЛУЧАЙ 1: Стандартная структура с двумя группами элементов
        composition = self._parse_standard_two_group_structure(all_data)
        if composition:
            st.success(f"✅ Образец {sample_index+1}: Использована стандартная структура")
            return composition
        
        # СЛУЧАЙ 2: Горизонтальная структура (элементы в строках, значения в столбцах)
        composition = self._parse_horizontal_structure(all_data)
        if composition:
            st.success(f"✅ Образец {sample_index+1}: Использована горизонтальная структура")
            return composition
        
        # СЛУЧАЙ 3: Вертикальная структура (элементы в столбцах, значения в строках)
        composition = self._parse_vertical_structure(all_data)
        if composition:
            st.success(f"✅ Образец {sample_index+1}: Использована вертикальная структура")
            return composition
        
        # СЛУЧАЙ 4: Резервный метод - поиск по шаблонам
        composition = self._parse_fallback_method(all_data)
        if composition:
            st.success(f"✅ Образец {sample_index+1}: Использован резервный метод")
            return composition
        
        st.warning(f"⚠️ Образец {sample_index+1}: Не удалось определить структуру таблицы")
        return {}

    def _parse_standard_two_group_structure(self, all_data):
        """Парсинг стандартной структуры с двумя группами элементов"""
        composition = {}
        
        try:
            # ПЕРВАЯ ГРУППА ЭЛЕМЕНТОВ (обычно строки 0-6)
            first_group_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni"]
            
            # Ищем строку с заголовками первой группы
            header_row_1 = None
            for i in range(min(5, len(all_data))):  # Ищем в первых 5 строках
                row = all_data[i]
                found_elements = [elem for elem in first_group_elements if any(elem in cell for cell in row)]
                if len(found_elements) >= 3:
                    header_row_1 = i
                    break
            
            if header_row_1 is not None:
                # Ищем строку со значениями для первой группы
                for value_row_idx in range(header_row_1 + 1, min(header_row_1 + 4, len(all_data))):
                    values_row = all_data[value_row_idx]
                    
                    # Сопоставляем заголовки со значениями
                    headers = all_data[header_row_1]
                    values = values_row
                    
                    for i, header in enumerate(headers):
                        if i < len(values):
                            for element in first_group_elements:
                                if element in header and self._is_numeric_value(values[i]):
                                    try:
                                        value = self._parse_numeric_value(values[i])
                                        composition[element] = value
                                        break
                                    except:
                                        continue
            
            # ВТОРАЯ ГРУППА ЭЛЕМЕНТОВ (обычно строки 7-13)
            second_group_elements = ["Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]
            
            # Ищем строку с заголовками второй группы
            header_row_2 = None
            for i in range(5, min(10, len(all_data))):  # Ищем в строках 5-9
                row = all_data[i]
                found_elements = [elem for elem in second_group_elements if any(elem in cell for cell in row)]
                if len(found_elements) >= 2:
                    header_row_2 = i
                    break
            
            if header_row_2 is not None:
                # Ищем строку со значениями для второй группы
                for value_row_idx in range(header_row_2 + 1, min(header_row_2 + 4, len(all_data))):
                    values_row = all_data[value_row_idx]
                    
                    # Сопоставляем заголовки со значениями
                    headers = all_data[header_row_2]
                    values = values_row
                    
                    for i, header in enumerate(headers):
                        if i < len(values):
                            for element in second_group_elements:
                                if element in header and self._is_numeric_value(values[i]):
                                    try:
                                        value = self._parse_numeric_value(values[i])
                                        composition[element] = value
                                        break
                                    except:
                                        continue
            
            return composition
            
        except Exception as e:
            return {}

    def _parse_horizontal_structure(self, all_data):
        """Парсинг горизонтальной структуры (элементы в строках)"""
        composition = {}
        
        try:
            chemical_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                               "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]
            
            for row in all_data:
                for i, cell in enumerate(row):
                    # Проверяем, содержит ли ячейка название элемента
                    for element in chemical_elements:
                        if element in cell.upper():
                            # Ищем числовое значение в соседних ячейках
                            for j in range(max(0, i-2), min(len(row), i+3)):
                                if j != i and self._is_numeric_value(row[j]):
                                    try:
                                        value = self._parse_numeric_value(row[j])
                                        composition[element] = value
                                        break
                                    except:
                                        continue
                            break
            
            return composition
            
        except Exception as e:
            return {}

    def _parse_vertical_structure(self, all_data):
        """Парсинг вертикальной структуры (элементы в столбцах)"""
        composition = {}
        
        try:
            chemical_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                               "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]
            
            # Транспонируем данные для анализа по столбцам
            if len(all_data) > 0:
                max_cols = max(len(row) for row in all_data)
                transposed_data = [[] for _ in range(max_cols)]
                
                for row in all_data:
                    for j, cell in enumerate(row):
                        if j < max_cols:
                            transposed_data[j].append(cell)
                
                # Анализируем каждый столбец
                for col_idx, column in enumerate(transposed_data):
                    element_found = None
                    for cell in column:
                        for element in chemical_elements:
                            if element in cell.upper():
                                element_found = element
                                break
                        if element_found:
                            break
                    
                    if element_found:
                        # Ищем числовое значение в этом столбце
                        for cell in column:
                            if self._is_numeric_value(cell):
                                try:
                                    value = self._parse_numeric_value(cell)
                                    composition[element_found] = value
                                    break
                                except:
                                    continue
            
            return composition
            
        except Exception as e:
            return {}

    def _parse_fallback_method(self, all_data):
        """Резервный метод парсинга - поиск по шаблонам"""
        composition = {}
        
        try:
            # Объединяем все данные в один текст для поиска по шаблонам
            full_text = " ".join([" ".join(row) for row in all_data])
            
            # Шаблоны для поиска элементов со значениями
            patterns = {
                "C": r"C[^A-Za-z0-9]*([0-9]+[,.][0-9]+)",
                "Si": r"Si[^A-Za-z0-9]*([0-9]+[,.][0-9]+)", 
                "Mn": r"Mn[^A-Za-z0-9]*([0-9]+[,.][0-9]+)",
                "P": r"P[^A-Za-z0-9]*([0-9]+[,.][0-9]+)",
                "S": r"S[^A-Za-z0-9]*([0-9]+[,.][0-9]+)",
                "Cr": r"Cr[^A-Za-z0-9]*([0-9]+[,.][0-9]+)",
                "Mo": r"Mo[^A-Za-z0-9]*([0-9]+[,.][0-9]+)",
                "Ni": r"Ni[^A-Za-z0-9]*([0-9]+[,.][0-9]+)",
                "Cu": r"Cu[^A-Za-z0-9]*([0-9]+[,.][0-9]+)",
                "V": r"V[^A-Za-z0-9]*([0-9]+[,.][0-9]+)"
            }
            
            for element, pattern in patterns.items():
                matches = re.findall(pattern, full_text, re.IGNORECASE)
                if matches:
                    try:
                        # Берем первое найденное значение
                        value_str = matches[0].replace(',', '.')
                        value = float(value_str)
                        composition[element] = value
                    except:
                        continue
            
            return composition
            
        except Exception as e:
            return {}

    def _interactive_table_parsing(self, all_data, sample_index):
        """Интерактивный режим парсинга таблицы для отладки"""
        composition = {}
        
        st.warning(f"🔧 РУЧНОЙ РЕЖИМ ДЛЯ ОБРАЗЦА {sample_index+1}")
        
        with st.expander("🎯 ИНТЕРАКТИВНЫЙ ПАРСИНГ", expanded=True):
            st.write("**Выберите соответствия элементов и значений:**")
            
            # Показываем таблицу с номерами строк и столбцов
            st.write("**Структура таблицы:**")
            debug_data = []
            for i, row in enumerate(all_data):
                row_data = {"Строка": i}
                for j, cell in enumerate(row):
                    row_data[f"Столбец {j}"] = f'"{cell}"'
                debug_data.append(row_data)
            
            debug_df = pd.DataFrame(debug_data).fillna("")
            st.dataframe(debug_df, use_container_width=True)
            
            # Позволяем пользователю вручную сопоставить элементы и значения
            chemical_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                               "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]
            
            for element in chemical_elements:
                col1, col2, col3 = st.columns([1, 2, 1])
                
                with col1:
                    st.write(f"**{element}**")
                
                with col2:
                    # Выбор строки и столбца для элемента
                    row_options = [f"Строка {i}" for i in range(len(all_data))]
                    col_options = [f"Столбец {j}" for j in range(len(all_data[0]) if all_data else 0)]
                    
                    selected_row = st.selectbox(
                        f"Строка для {element}",
                        options=row_options,
                        key=f"manual_{sample_index}_{element}_row"
                    )
                    
                    selected_col = st.selectbox(
                        f"Столбец для {element}",
                        options=col_options,
                        key=f"manual_{sample_index}_{element}_col"
                    )
                
                with col3:
                    # Извлекаем значение
                    if selected_row and selected_col:
                        row_idx = int(selected_row.split(" ")[1])
                        col_idx = int(selected_col.split(" ")[1])
                        
                        if (row_idx < len(all_data) and 
                            col_idx < len(all_data[row_idx]) and
                            self._is_numeric_value(all_data[row_idx][col_idx])):
                            
                            value = self._parse_numeric_value(all_data[row_idx][col_idx])
                            composition[element] = value
                            st.success(f"{value}")
                        else:
                            st.warning("Не число")
            
            # Кнопка применения ручных настроек
            if st.button(f"✅ Применить ручные настройки для образца {sample_index+1}"):
                st.success(f"Ручные настройки применены для {len(composition)} элементов")
        
        return composition

    def _is_numeric_value(self, text):
        """Проверяет, является ли текст числовым значением"""
        if not text or text.strip() == "":
            return False
        
        # Очищаем текст от лишних символов
        clean_text = text.replace(',', '.').replace('±', ' ').replace(' ', '').split()[0]
        
        # Проверяем на число
        try:
            float(clean_text)
            return True
        except:
            return False

    def _parse_numeric_value(self, text):
        """Извлекает числовое значение из текста"""
        if not text:
            return 0.0
        
        # Очищаем текст
        clean_text = text.replace(',', '.').replace('±', ' ').split()[0]
        
        try:
            return float(clean_text)
        except:
            raise ValueError(f"Не могу преобразовать '{text}' в число")
    
    def match_sample_names(self, samples, correct_names_file):
        """Сопоставление названий образцов с правильными названиями"""
        if not correct_names_file:
            return samples, []
        
        correct_samples = self.name_matcher.parse_correct_names(correct_names_file.getvalue())
        
        if not correct_samples:
            st.warning("Не удалось загрузить правильные названия образцов")
            return samples, []
        
        matched_samples = []
        unmatched_samples = []
        
        for sample in samples:
            # ВАЖНО: сохраняем оригинальное название до любых изменений
            original_protocol_name = sample['name']
            
            protocol_sample_info = self.name_matcher.parse_protocol_sample_name(original_protocol_name)
            best_match = self.name_matcher.find_best_match(protocol_sample_info, correct_samples)
            
            if best_match:
                # Создаем копию образца с исправленным названием и номером
                corrected_sample = sample.copy()
                corrected_sample['original_name'] = original_protocol_name  # Сохраняем оригинальное название
                corrected_sample['name'] = best_match['original']           # Заменяем на правильное
                corrected_sample['correct_number'] = best_match['number']   # Сохраняем номер для сортировки
                corrected_sample['automatically_matched'] = True
                matched_samples.append(corrected_sample)
            else:
                # Если совпадение не найдено, оставляем оригинальное название
                sample['original_name'] = original_protocol_name  # Сохраняем для информации
                sample['correct_number'] = None                   # Нет номера для сортировки
                sample['automatically_matched'] = False
                unmatched_samples.append(sample)
        
        # Выводим информацию о сопоставлении
        if matched_samples:
            st.success(f"Успешно сопоставлено {len(matched_samples)} образцов")
            
            with st.expander("📋 Просмотр сопоставленных образцов"):
                match_data = []
                for sample in matched_samples:
                    match_data.append({
                        'Номер': sample['correct_number'],
                        'Исходное название': sample['original_name'],
                        'Правильное название': sample['name']
                    })
                # Сортируем по номеру
                match_data.sort(key=lambda x: x['Номер'])
                st.table(pd.DataFrame(match_data))
        
        if unmatched_samples:
            st.warning(f"Не удалось сопоставить {len(unmatched_samples)} образцов")
            
            with st.expander("⚠️ Просмотр несопоставленных образцов"):
                unmatched_data = []
                for sample in unmatched_samples:
                    unmatched_data.append({
                        'Образец': sample['original_name'],
                        'Марка стали': sample['steel_grade']
                    })
                st.table(pd.DataFrame(unmatched_data))
        
        # Сортируем сопоставленные образцы по номеру, несопоставленные оставляем в конце
        matched_samples.sort(key=lambda x: x['correct_number'] if x['correct_number'] is not None else float('inf'))
        return matched_samples + unmatched_samples, correct_samples
    
    def check_element_compliance(self, element, value, standard):
        """Проверка соответствия элемента нормативам"""
        if element not in standard or element == "source":
            return "normal"
        
        min_val, max_val = standard[element]
        
        # Проверка на отклонение
        if min_val is not None and value < min_val:
            return "deviation"
        elif max_val is not None and value > max_val:
            return "deviation"
        else:
            return "normal"
    
    def create_report_table_with_original_names(self, samples):
        """Создание сводной таблицы для отчета с колонкой исходных названий"""
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
            
            # Для стали 12Х1МФ устанавливаем особый порядок столбцов
            if grade == "12Х1МФ":
                # Порядок: основные элементы, затем вредные примеси
                main_elements = ["C", "Si", "Mn", "Cr", "Mo", "V", "Ni"]
                harmful_elements = ["Cu", "S", "P"]
                # Добавляем остальные элементы, если есть
                other_elements = [elem for elem in norm_elements if elem not in main_elements + harmful_elements]
                norm_elements = main_elements + other_elements + harmful_elements
            
            # Сортируем образцы
            grade_samples_sorted = sorted(
                grade_samples, 
                key=lambda x: x.get('correct_number', float('inf')) if x.get('correct_number') is not None else float('inf')
            )
            
            # Создаем DataFrame с колонкой исходных названий
            data = []
            compliance_data = []  # Для хранения информации о соответствии
            
            # Добавляем образцы - нумерация начинается с 1 для каждой таблицы
            for idx, sample in enumerate(grade_samples_sorted, 1):
                # Используем порядковый номер в таблице (начинается с 1)
                display_number = idx
                
                # Добавляем колонку с исходным названием
                original_name = sample.get('original_name', '')
                row = {
                    "№": display_number, 
                    "Исходное название": original_name,
                    "Образец": sample["name"]
                }
                compliance_row = {"№": "normal", "Исходное название": "normal", "Образец": "normal"}
                
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
            requirements_row = {"№": "", "Исходное название": "", "Образец": f"Требования ТУ 14-3Р-55-2001 для стали марки {grade}"}
            requirements_compliance = {"№": "requirements", "Исходное название": "requirements", "Образец": "requirements"}
            
            for elem in norm_elements:
                min_val, max_val = standard[elem]
                if min_val is not None and max_val is not None:
                    if elem in ["S", "P"]:
                        requirements_row[elem] = f"{min_val:.3f}-{max_val:.3f}".replace('.', ',')
                    else:
                        requirements_row[elem] = f"{min_val:.2f}-{max_val:.2f}".replace('.', ',')
                elif min_val is not None:
                    if elem in ["S", "P"]:
                        requirements_row[elem] = f"≥{min_val:.3f}".replace('.', ',')
                    else:
                        requirements_row[elem] = f"≥{min_val:.2f}".replace('.', ',')
                elif max_val is not None:
                    if elem in ["S", "P"]:
                        requirements_row[elem] = f"≤{max_val:.3f}".replace('.', ',')
                    else:
                        requirements_row[elem] = f"≤{max_val:.2f}".replace('.', ',')
                else:
                    requirements_row[elem] = "не нормируется"
                
                requirements_compliance[elem] = "requirements"
            
            data.append(requirements_row)
            compliance_data.append(requirements_compliance)
            
            tables[grade] = {
                "data": pd.DataFrame(data),
                "compliance": compliance_data,
                "columns_order": ["№", "Исходное название", "Образец"] + norm_elements
            }
        
        return tables

# Остальные функции (add_manual_matching_interface, add_manual_steel_grade_correction, 
# add_manual_composition_correction, apply_styling, set_font_times_new_roman, 
# create_word_report) остаются без изменений...

def main():
    st.set_page_config(page_title="Анализатор химсостава металла", layout="wide")
    st.title("🔬 Анализатор химического состава металла")
    
    # Инициализация session_state для хранения образцов
    if 'final_samples' not in st.session_state:
        st.session_state.final_samples = None
    if 'manually_matched_samples' not in st.session_state:
        st.session_state.manually_matched_samples = None
    
    try:
        analyzer = ChemicalAnalyzer()
        
        # Переключатель режима отладки
        analyzer.debug_mode = st.sidebar.checkbox("🔧 Включить режим отладки", value=False)
        
        if analyzer.debug_mode:
            st.sidebar.info("Режим отладки включен. Будут показаны детальные отладочные информации.")
        
        # Остальной код остается без изменений...
        # [Здесь должен быть остальной код из предыдущей версии]

    except Exception as e:
        st.error(f"Произошла ошибка при запуске приложения: {str(e)}")
        import traceback
        st.error(f"Детали ошибки: {traceback.format_exc()}")

if __name__ == "__main__":
    main()
