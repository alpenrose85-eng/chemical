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
from difflib import SequenceMatcher

class SampleNameMatcher:
    def __init__(self):
        self.surface_types = {
            'ЭПК': ['ЭПК'],
            'ШПП': ['ШПП'],
            'ПС КШ': ['ПС КШ', 'ПТ КШ', 'ПТКМ', 'ПТ КМ', 'ПТКМ'],
            'КПП ВД': ['КПП ВД', 'ВД'],
            'КПП НД-1': ['КПП НД-1', 'КПП НД-I', 'НД-1', 'НД-I'],
            'КПП НД-2': ['КПП НД-2', 'КПП НД-II', 'НД-2', 'НД-II']
        }
        self.letters = ['А', 'Б', 'В', 'Г']
        self.roman_to_arabic = {'I': '1', 'II': '2', 'III': '3', 'IV': '4'}
    
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
        # Сначала проверяем точные совпадения
        for surface_type, patterns in self.surface_types.items():
            for pattern in patterns:
                if pattern in name:
                    return surface_type
        
        # Если точных совпадений нет, ищем похожие названия
        for surface_type, patterns in self.surface_types.items():
            for pattern in patterns:
                if self.similar(pattern, name) > 0.7:  # 70% схожести
                    return surface_type
        
        return None
    
    def similar(self, a, b):
        """Вычисление схожести строк"""
        return SequenceMatcher(None, a, b).ratio()
    
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
        
        # Ищем просто число в скобках
        matches = re.findall(r'\((\d+)\)', name)
        if matches:
            return matches[0]
        
        # Ищем любое число в названии
        matches = re.findall(r'\b(\d+)\b', name)
        if matches:
            return matches[0]
            
        return None
    
    def extract_letter(self, name):
        """Извлечение буквы (А, Б, В, Г) из названия"""
        # Ищем букву в скобках в конце: (30,Г) или (А)
        matches = re.findall(r'\([^)]*([А-Г])\)', name)
        if matches:
            return matches[0]
        
        # Ищем букву после запятой в скобках
        matches = re.findall(r',\s*([А-Г])\)', name)
        if matches:
            return matches[0]
        
        # Ищем букву в скобках отдельно
        matches = re.findall(r'\(([А-Г])\)', name)
        if matches:
            return matches[0]
            
        return None
    
    def parse_protocol_sample_name(self, sample_name):
        """Парсинг названия образца из протокола химического анализа"""
        # Определяем букву из префикса (НА, НБ, НВ, НГ) - это нитка!
        letter_map = {'НА': 'А', 'НБ': 'Б', 'НВ': 'В', 'НГ': 'Г', 'Н-Г': 'Г'}
        letter = None
        
        # Сначала ищем букву нитки из префикса (самый надежный способ)
        for prefix, mapped_letter in letter_map.items():
            if sample_name.startswith(prefix):
                letter = mapped_letter
                break
        
        # Если не нашли по префиксу, ищем в тексте
        if not letter:
            for l in self.letters:
                if f'Н{l}' in sample_name or f'Н-{l}' in sample_name:
                    letter = l
                    break
        
        # Определяем тип поверхности (с учетом римских цифр)
        surface_type = None
        
        # Заменяем римские цифры на арабские для унификации
        normalized_name = sample_name
        for roman, arabic in self.roman_to_arabic.items():
            normalized_name = normalized_name.replace(roman, arabic)
        
        for stype, patterns in self.surface_types.items():
            for pattern in patterns:
                # Нормализуем паттерн тоже
                normalized_pattern = pattern
                for roman, arabic in self.roman_to_arabic.items():
                    normalized_pattern = normalized_pattern.replace(roman, arabic)
                
                if normalized_pattern in normalized_name:
                    surface_type = stype
                    break
            if surface_type:
                break
        
        # Если точного совпадения нет, ищем похожие
        if not surface_type:
            for stype, patterns in self.surface_types.items():
                for pattern in patterns:
                    normalized_pattern = pattern
                    for roman, arabic in self.roman_to_arabic.items():
                        normalized_pattern = normalized_pattern.replace(roman, arabic)
                    
                    if self.similar(normalized_pattern, normalized_name) > 0.7:
                        surface_type = stype
                        break
                if surface_type:
                    break
        
        # Извлекаем номер трубы
        tube_number = None
        numbers = re.findall(r'\d+', sample_name)
        if numbers:
            # Ищем число после "НГ" или подобных конструкций
            pattern_match = re.search(r'Н[А-Г-]\s*(\d+)', sample_name)
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
    
    def match_samples(self, protocol_samples, correct_samples):
        """Многоэтапное сопоставление образцов"""
        matched_samples = []
        unmatched_protocol = protocol_samples.copy()
        used_correct = set()
        
        # Этап 1: Точное совпадение типа поверхности + номера + буквы
        matches_stage1 = self._match_stage1(unmatched_protocol, correct_samples, used_correct)
        matched_samples.extend(matches_stage1)
        unmatched_protocol = [s for s in unmatched_protocol if s not in [m[0] for m in matches_stage1]]
        
        # Этап 2: Совпадение типа поверхности + номера
        matches_stage2 = self._match_stage2(unmatched_protocol, correct_samples, used_correct)
        matched_samples.extend(matches_stage2)
        unmatched_protocol = [s for s in unmatched_protocol if s not in [m[0] for m in matches_stage2]]
        
        # Этап 3: Совпадение типа поверхности + буквы
        matches_stage3 = self._match_stage3(unmatched_protocol, correct_samples, used_correct)
        matched_samples.extend(matches_stage3)
        unmatched_protocol = [s for s in unmatched_protocol if s not in [m[0] for m in matches_stage3]]
        
        # Этап 4: Частичное совпадение типа поверхности
        matches_stage4 = self._match_stage4(unmatched_protocol, correct_samples, used_correct)
        matched_samples.extend(matches_stage4)
        unmatched_protocol = [s for s in unmatched_protocol if s not in [m[0] for m in matches_stage4]]
        
        return matched_samples, unmatched_protocol
    
    def _match_stage1(self, protocol_samples, correct_samples, used_correct):
        """Этап 1: Точное совпадение типа поверхности + номера + буквы"""
        matches = []
        for protocol in protocol_samples:
            protocol_info = self.parse_protocol_sample_name(protocol['name'])
            best_match = None
            
            for correct in correct_samples:
                if correct['original'] in used_correct:
                    continue
                
                if (protocol_info['surface_type'] == correct['surface_type'] and
                    protocol_info['tube_number'] == correct['tube_number'] and
                    protocol_info['letter'] == correct['letter']):
                    best_match = correct
                    break
            
            if best_match:
                matches.append((protocol, best_match, "Этап 1: полное совпадение"))
                used_correct.add(best_match['original'])
        
        return matches
    
    def _match_stage2(self, protocol_samples, correct_samples, used_correct):
        """Этап 2: Совпадение типа поверхности + номера"""
        matches = []
        for protocol in protocol_samples:
            protocol_info = self.parse_protocol_sample_name(protocol['name'])
            best_match = None
            
            for correct in correct_samples:
                if correct['original'] in used_correct:
                    continue
                
                if (protocol_info['surface_type'] == correct['surface_type'] and
                    protocol_info['tube_number'] == correct['tube_number']):
                    best_match = correct
                    break
            
            if best_match:
                matches.append((protocol, best_match, "Этап 2: тип+номер"))
                used_correct.add(best_match['original'])
        
        return matches
    
    def _match_stage3(self, protocol_samples, correct_samples, used_correct):
        """Этап 3: Совпадение типа поверхности + буквы"""
        matches = []
        for protocol in protocol_samples:
            protocol_info = self.parse_protocol_sample_name(protocol['name'])
            best_match = None
            
            for correct in correct_samples:
                if correct['original'] in used_correct:
                    continue
                
                if (protocol_info['surface_type'] == correct['surface_type'] and
                    protocol_info['letter'] == correct['letter']):
                    best_match = correct
                    break
            
            if best_match:
                matches.append((protocol, best_match, "Этап 3: тип+буква"))
                used_correct.add(best_match['original'])
        
        return matches
    
    def _match_stage4(self, protocol_samples, correct_samples, used_correct):
        """Этап 4: Частичное совпадение типа поверхности"""
        matches = []
        for protocol in protocol_samples:
            protocol_info = self.parse_protocol_sample_name(protocol['name'])
            best_match = None
            best_similarity = 0
            
            for correct in correct_samples:
                if correct['original'] in used_correct:
                    continue
                
                # Проверяем схожесть названий поверхностей
                if protocol_info['surface_type'] and correct['surface_type']:
                    similarity = self.similar(protocol_info['surface_type'], correct['surface_type'])
                    if similarity > 0.7 and similarity > best_similarity:
                        best_similarity = similarity
                        best_match = correct
            
            if best_match:
                matches.append((protocol, best_match, f"Этап 4: похожий тип ({best_similarity:.2f})"))
                used_correct.add(best_match['original'])
        
        return matches

class ChemicalAnalyzer:
    def __init__(self):
        self.load_standards()
        self.name_matcher = SampleNameMatcher()
        
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
    
    def match_sample_names(self, samples, correct_names_file):
        """Сопоставление названий образцов с правильными названиями"""
        if not correct_names_file:
            return samples, []
        
        correct_samples = self.name_matcher.parse_correct_names(correct_names_file.getvalue())
        
        if not correct_samples:
            st.warning("Не удалось загрузить правильные названия образцов")
            return samples, []
        
        # Многоэтапное сопоставление
        matched_pairs, unmatched_protocol = self.name_matcher.match_samples(samples, correct_samples)
        
        # Формируем итоговые списки
        matched_samples = []
        for protocol_sample, correct_sample, match_stage in matched_pairs:
            corrected_sample = protocol_sample.copy()
            corrected_sample['original_name'] = protocol_sample['name']
            corrected_sample['name'] = correct_sample['original']
            corrected_sample['correct_number'] = correct_sample['number']
            corrected_sample['automatically_matched'] = True
            corrected_sample['match_stage'] = match_stage
            matched_samples.append(corrected_sample)
        
        unmatched_samples = []
        for sample in unmatched_protocol:
            sample['original_name'] = sample['name']
            sample['correct_number'] = None
            sample['automatically_matched'] = False
            unmatched_samples.append(sample)
        
        # Выводим информацию о сопоставлении
        if matched_samples:
            st.success(f"Успешно сопоставлено {len(matched_samples)} образцов")
            
            with st.expander("📋 Детали автоматического сопоставления"):
                match_data = []
                for sample in matched_samples:
                    protocol_info = self.name_matcher.parse_protocol_sample_name(sample['original_name'])
                    match_data.append({
                        'Номер': sample['correct_number'],
                        'Исходное название': sample['original_name'],
                        'Правильное название': sample['name'],
                        'Этап': sample.get('match_stage', 'н/д'),
                        'Тип': protocol_info['surface_type'] or 'н/д',
                        'Труба': protocol_info['tube_number'] or 'н/д',
                        'Нитка': protocol_info['letter'] or 'н/д'
                    })
                # Сортируем по номеру
                match_data.sort(key=lambda x: x['Номер'])
                st.table(pd.DataFrame(match_data))
        
        if unmatched_samples:
            st.warning(f"Не удалось сопоставить {len(unmatched_samples)} образцов")
            
            with st.expander("⚠️ Просмотр несопоставленных образцов"):
                unmatched_data = []
                for sample in unmatched_samples:
                    protocol_info = self.name_matcher.parse_protocol_sample_name(sample['name'])
                    unmatched_data.append({
                        'Образец': sample['original_name'],
                        'Марка стали': sample['steel_grade'],
                        'Тип': protocol_info['surface_type'] or 'н/д',
                        'Труба': protocol_info['tube_number'] or 'н/д',
                        'Нитка': protocol_info['letter'] or 'н/д'
                    })
                st.table(pd.DataFrame(unmatched_data))
        
        # Сортируем сопоставленные образцы по номеру, несопоставленные оставляем в конце
        matched_samples.sort(key=lambda x: x['correct_number'])
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
            
            # Для стали 12Х1МФ устанавливаем особый порядок столбцов
            if grade == "12Х1МФ":
                main_elements = ["C", "Si", "Mn", "Cr", "Mo", "V", "Ni"]
                harmful_elements = ["Cu", "S", "P"]
                other_elements = [elem for elem in norm_elements if elem not in main_elements + harmful_elements]
                norm_elements = main_elements + other_elements + harmful_elements
            
            # Сортируем образцы
            grade_samples_sorted = sorted(
                grade_samples,
                key=lambda x: (x.get('correct_number') is None, x.get('correct_number', float('inf')))
            )
            
            # Создаем DataFrame
            data = []
            compliance_data = []
            
            # Добавляем образцы с последовательной нумерацией
            for idx, sample in enumerate(grade_samples_sorted, 1):
                display_number = idx
                
                row = {
                    "№": display_number, 
                    "Образец": sample["name"]
                }
                compliance_row = {"№": "normal", "Образец": "normal"}
                
                for elem in norm_elements:
                    if elem in sample["composition"]:
                        value = sample["composition"][elem]
                        if elem in ["S", "P"]:
                            row[elem] = f"{value:.3f}".replace('.', ',')
                        else:
                            row[elem] = f"{value:.2f}".replace('.', ',')
                        
                        status = self.check_element_compliance(elem, value, standard)
                        compliance_row[elem] = status
                    else:
                        row[elem] = "-"
                        compliance_row[elem] = "normal"
                
                data.append(row)
                compliance_data.append(compliance_row)
            
            # Добавляем строку с нормативами
            requirements_row = {"№": "", "Образец": f"Требования ТУ 14-3Р-55-2001 для стали марки {grade}"}
            requirements_compliance = {"№": "requirements", "Образец": "requirements"}
            
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
                "columns_order": ["№", "Образец"] + norm_elements
            }
        
        return tables

# Остальные функции (add_manual_matching_interface, apply_styling, set_font_times_new_roman, main, create_word_report) 
# остаются без изменений, как в предыдущем коде

def add_manual_matching_interface(samples, correct_samples, analyzer):
    """Интерфейс для ручного сопоставления образцов с подсветкой проблем"""
    st.header("🔧 Ручное сопоставление образцов")
    
    editable_samples = samples.copy()
    correct_names_dict = {cs['original']: cs for cs in correct_samples}
    correct_names_list = [cs['original'] for cs in correct_samples]
    
    # Анализируем потенциальные конфликты
    conflict_candidates = {}
    for correct in correct_samples:
        potential_matches = []
        for sample in editable_samples:
            protocol_info = analyzer.name_matcher.parse_protocol_sample_name(sample['name'])
            if (protocol_info['surface_type'] == correct['surface_type'] and
                protocol_info['tube_number'] == correct['tube_number']):
                potential_matches.append(sample)
        
        if len(potential_matches) > 1:
            conflict_candidates[correct['original']] = potential_matches
    
    options = ["Не сопоставлен"] + correct_names_list
    manual_matches = {}
    
    st.write("**Сопоставьте образцы вручную:**")
    
    for i, sample in enumerate(editable_samples):
        col1, col2 = st.columns([2, 3])
        
        with col1:
            # Подсвечиваем проблемные образцы
            is_conflict = any(sample in matches for matches in conflict_candidates.values())
            conflict_color = "🔴" if is_conflict else "⚪"
            
            st.write(f"{conflict_color} **{sample.get('original_name', sample['name'])}**")
            if sample.get('steel_grade'):
                st.write(f"*Марка: {sample['steel_grade']}*")
            
            protocol_info = analyzer.name_matcher.parse_protocol_sample_name(sample['name'])
            st.write(f"*Тип: {protocol_info['surface_type'] or 'н/д'}*")
            st.write(f"*Труба: {protocol_info['tube_number'] or 'н/д'}*")
            st.write(f"*Нитка: {protocol_info['letter'] or 'н/д'}*")
            
            if is_conflict:
                st.warning("⚠️ Возможный конфликт сопоставления")
        
        with col2:
            current_match = sample['name'] if sample['name'] in correct_names_list else "Не сопоставлен"
            
            selected = st.selectbox(
                f"Выберите правильное название для образца {i+1}",
                options=options,
                index=options.index(current_match) if current_match in options else 0,
                key=f"manual_match_{i}"
            )
            
            if selected != "Не сопоставлен":
                manual_matches[sample['name']] = selected
    
    if st.button("✅ Применить ручное сопоставление"):
        updated_samples = []
        
        for sample in editable_samples:
            if sample['name'] in manual_matches:
                correct_name = manual_matches[sample['name']]
                correct_sample = correct_names_dict[correct_name]
                
                updated_sample = sample.copy()
                updated_sample['original_name'] = sample['name']
                updated_sample['name'] = correct_name
                updated_sample['correct_number'] = correct_sample['number']
                updated_sample['manually_matched'] = True
                
                updated_samples.append(updated_sample)
            else:
                sample['manually_matched'] = False
                updated_samples.append(sample)
        
        st.success(f"Ручное сопоставление применено! Обновлено {len(manual_matches)} образцов.")
        return updated_samples
    
    return editable_samples

def apply_styling(df, compliance_data):
    """Применяет стили к DataFrame на основе данных о соответствии"""
    styled_df = df.copy()
    
    styles = []
    for i, row in df.iterrows():
        if i < len(compliance_data):
            compliance_row = compliance_data[i]
            for col in df.columns:
                if col in compliance_row:
                    status = compliance_row[col]
                    if status == "deviation":
                        styles.append(f"background-color: #ffcccc; color: #cc0000; font-weight: bold;")
                    elif status == "requirements":
                        styles.append(f"background-color: #f0f0f0; font-style: italic;")
                    else:
                        styles.append("")
                else:
                    styles.append("")
    
    styled = df.style
    for i in range(len(df)):
        for j, col in enumerate(df.columns):
            idx = i * len(df.columns) + j
            if idx < len(styles) and styles[idx]:
                styled = styled.set_properties(subset=(i, col), **{'css': styles[idx]})
    
    return styled

def set_font_times_new_roman(doc):
    """Устанавливает шрифт Times New Roman для всего документа"""
    styles = doc.styles
    for style in styles:
        if hasattr(style, 'font'):
            style.font.name = 'Times New Roman'
    
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'

def main():
    st.set_page_config(page_title="Анализатор химсостава металла", layout="wide")
    st.title("🔬 Анализатор химического состава металла")
    
    analyzer = ChemicalAnalyzer()
    
    # Сайдбар для управления нормативами
    with st.sidebar:
        st.header("Управление нормативами")
        
        st.subheader("Существующие марки стали")
        selected_standard = st.selectbox(
            "Выберите марку для просмотра",
            options=list(analyzer.standards.keys())
        )
        
        if selected_standard:
            st.write(f"**Норматив для {selected_standard}:**")
            standard = analyzer.standards[selected_standard]
            for elem, value_range in standard.items():
                if elem == "source":
                    continue
                
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
        
        st.subheader("Добавить новую марку стали")
        
        new_grade = st.text_input("Марка стали")
        new_source = st.text_input("Нормативный документ", value="ТУ 14-3Р-55-2001")
        
        if new_grade:
            st.write("**Добавление элементов:**")
            
            if 'elements' not in st.session_state:
                st.session_state.elements = []
            
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
                
                for i in sorted(elements_to_remove, reverse=True):
                    st.session_state.elements.pop(i)
            
            if st.button("💾 Сохранить норматив"):
                if not st.session_state.elements:
                    st.error("Добавьте хотя бы один элемент!")
                elif new_grade in analyzer.standards:
                    st.error(f"Марка стали {new_grade} уже существует!")
                else:
                    elements_ranges = {}
                    for elem_data in st.session_state.elements:
                        elements_ranges[elem_data["element"]] = (
                            elem_data["min"], 
                            elem_data["max"]
                        )
                    
                    elements_ranges["source"] = new_source
                    analyzer.standards[new_grade] = elements_ranges
                    analyzer.save_user_standards()
                    
                    st.session_state.elements = []
                    st.success(f"Норматив для {new_grade} сохранен!")
    
    # Основная область для загрузки файлов
    st.header("Загрузка протоколов")
    
    st.subheader("1. Загрузите файл с правильными названиями образцов")
    correct_names_file = st.file_uploader(
        "Файл с правильными названиями (.docx)",
        type=["docx"],
        key="correct_names"
    )
    
    correct_samples = []
    if correct_names_file:
        correct_samples = analyzer.name_matcher.parse_correct_names(correct_names_file.getvalue())
        if correct_samples:
            st.success(f"Загружено {len(correct_samples)} правильных названий образцов")
            with st.expander("📋 Просмотр загруженных названий"):
                preview_data = []
                for sample in correct_samples:
                    preview_data.append({
                        'Номер': sample['number'],
                        'Название': sample['original'],
                        'Тип': sample['surface_type'] or 'н/д',
                        'Труба': sample['tube_number'] or 'н/д', 
                        'Нитка': sample['letter'] or 'н/д'
                    })
                st.table(pd.DataFrame(preview_data))
    
    st.subheader("2. Загрузите файлы протоколов химического анализа")
    uploaded_files = st.file_uploader(
        "Файлы протоколов (.docx)", 
        type=["docx"], 
        accept_multiple_files=True,
        key="protocol_files"
    )
    
    all_samples = []
    
    if uploaded_files:
        for uploaded_file in uploaded_files:
            samples = analyzer.parse_protocol_file(uploaded_file.getvalue())
            all_samples.extend(samples)
        
        if correct_names_file and correct_samples:
            st.subheader("🔍 Автоматическое сопоставление названий образцов")
            all_samples, correct_samples_loaded = analyzer.match_sample_names(all_samples, correct_names_file)
            
            all_samples = add_manual_matching_interface(all_samples, correct_samples_loaded, analyzer)
        
        if all_samples:
            st.header("Результаты анализа")
            
            st.markdown("""
            **Легенда:**
            - <span style='background-color: #ffcccc; padding: 2px 5px; border-radius: 3px;'>🔴 Красный</span> - отклонение от норм
            - <span style='background-color: #f0f0f0; padding: 2px 5px; border-radius: 3px;'>⚪ Серый</span> - нормативные требования
            """, unsafe_allow_html=True)
            
            report_tables = analyzer.create_report_table_with_original_names(all_samples)
            
            export_tables = {}
            
            if report_tables:
                for grade, table_data in report_tables.items():
                    st.subheader(f"Марка стали: {grade}")
                    
                    styled_table = apply_styling(table_data["data"], table_data["compliance"])
                    st.dataframe(styled_table, use_container_width=True, hide_index=True)
                    
                    export_tables[grade] = table_data["data"]
                
                if st.button("📄 Экспорт в Word"):
                    create_word_report(export_tables, all_samples, analyzer)
                    st.success("Отчет готов к скачиванию!")
            
            st.header("Обработанные образцы")
            for sample in all_samples:
                with st.expander(f"📋 {sample['name']} - {sample['steel_grade']}"):
                    if 'original_name' in sample:
                        st.write(f"**Исходное название:** {sample['original_name']}")
                    if 'correct_number' in sample:
                        st.write(f"**Номер в списке:** {sample['correct_number']}")
                    st.write(f"**Марка стали:** {sample['steel_grade']}")
                    st.write("**Химический состав:**")
                    for element, value in sample['composition'].items():
                        st.write(f"- {element}: {value}")

def create_word_report(tables, samples, analyzer):
    """Создание Word отчета"""
    try:
        doc = Document()
        
        set_font_times_new_roman(doc)
        
        title = doc.add_heading('Протокол анализа химического состава', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        doc.add_paragraph(f"Проанализировано образцов: {len(samples)}")
        doc.add_paragraph("")
        
        doc.add_heading('Легенда', level=1)
        legend_table = doc.add_table(rows=3, cols=2)
        legend_table.style = 'Table Grid'
        
        legend_table.cell(0, 0).text = "Цвет"
        legend_table.cell(0, 1).text = "Значение"
        
        legend_table.cell(1, 0).text = "🔴"
        legend_table.cell(1, 1).text = "Отклонение от норм"
        
        legend_table.cell(2, 0).text = "⚪"
        legend_table.cell(2, 1).text = "Нормативные требования"
        
        doc.add_paragraph()
        
        for grade, table_df in tables.items():
            doc.add_heading(f'Марка стали: {grade}', level=1)
            
            word_table = doc.add_table(rows=len(table_df)+1, cols=len(table_df.columns))
            word_table.style = 'Table Grid'
            
            for j, col in enumerate(table_df.columns):
                word_table.cell(0, j).text = str(col)
            
            for i, row in table_df.iterrows():
                for j, col in enumerate(table_df.columns):
                    word_table.cell(i+1, j).text = str(row[col])
            
            doc.add_paragraph()
        
        doc.save("химический_анализ_отчет.docx")
        
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
