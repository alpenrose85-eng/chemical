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
                    composition = self.parse_composition_table_improved(table)
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
    
    def parse_composition_table_improved(self, table):
        """Улучшенный парсинг таблицы с химическим составом"""
        composition = {}
        
        try:
            # Собираем все данные из таблицы
            all_data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_data:  # Только непустые строки
                    all_data.append(row_data)
            
            if not all_data:
                return composition
            
            # Все возможные элементы для поиска
            all_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                           "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]
            
            # Ищем строки с элементами и значениями
            elements_found = []
            values_found = []
            
            # Проходим по всем строкам и ищем химические элементы
            for row in all_data:
                for cell in row:
                    # Ищем обозначения элементов
                    if cell in all_elements:
                        elements_found.append(cell)
            
            # Ищем строку со средними значениями
            for i, row in enumerate(all_data):
                # Ищем строку, где большинство ячеек - числа
                numeric_count = 0
                for cell in row:
                    try:
                        # Пробуем преобразовать в число (учитываем формат с запятой)
                        cell_clean = cell.replace(',', '.').replace('±', ' ').split()[0]
                        float(cell_clean)
                        numeric_count += 1
                    except:
                        pass
                
                if numeric_count >= 5:  # Если хотя бы 5 чисел в строке
                    values_found = row
                    break
            
            # Альтернативный подход: берем предпоследнюю строку (часто там средние значения)
            if not values_found and len(all_data) >= 2:
                values_found = all_data[-2]  # Предпоследняя строка
            
            # Если не нашли значения, попробуем найти строку со словом "Среднее"
            if not values_found:
                for i, row in enumerate(all_data):
                    if any('Среднее' in cell for cell in row):
                        if i+1 < len(all_data):
                            values_found = all_data[i+1]
                        break
            
            # Сопоставляем элементы со значениями по позициям
            if elements_found and values_found:
                # Простой подход: предполагаем стандартный порядок элементов
                standard_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni", 
                                   "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]
                
                for idx, elem in enumerate(standard_elements):
                    if idx < len(values_found):
                        value_str = values_found[idx]
                        try:
                            value_str = value_str.replace(',', '.').replace('±', ' ').split()[0]
                            value = float(value_str)
                            composition[elem] = value
                        except:
                            continue
            
            return composition
            
        except Exception as e:
            st.error(f"Ошибка при парсинге таблицы: {str(e)}")
            return {}
    
    def match_sample_names(self, samples, correct_names_file):
        """Сопоставление названий образцов с правильными названиями - УЛУЧШЕННАЯ ВЕРСИЯ"""
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
        # ИСПРАВЛЕНИЕ: избегаем сравнения None с int
        matched_samples.sort(key=lambda x: x['correct_number'] if x['correct_number'] is not None else float('inf'))
        return matched_samples + unmatched_samples, correct_samples
    
    def check_element_compliance(self, element, value, standard):
        """Проверка соответствия элемента нормативам - УПРОЩЕННАЯ ВЕРСИЯ"""
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
            
            # ИСПРАВЛЕНИЕ: сортируем образцы, избегая сравнения None с int
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

def add_manual_matching_interface(samples, correct_samples, analyzer):
    """Интерфейс для ручного сопоставления образцов с фильтрацией - ИСПРАВЛЕННАЯ ВЕРСИЯ"""
    st.header("🔧 Ручное сопоставление образцов")
    
    # Используем session_state для хранения результатов ручного сопоставления
    if 'manually_matched_samples' not in st.session_state:
        st.session_state.manually_matched_samples = samples.copy()
    
    # Создаем копию samples для редактирования
    editable_samples = st.session_state.manually_matched_samples.copy()
    
    # Создаем словарь для быстрого доступа к правильным названиям
    correct_names_dict = {cs['original']: cs for cs in correct_samples}
    correct_names_list = [cs['original'] for cs in correct_samples]
    
    # Добавляем опцию "Не сопоставлен"
    options = ["Не сопоставлен"] + correct_names_list
    
    st.write("**Сопоставьте образцы вручную:**")
    
    manual_matches = {}
    
    for i, sample in enumerate(editable_samples):
        col1, col2, col3 = st.columns([2, 2, 3])
        
        with col1:
            # ВАЖНОЕ ИСПРАВЛЕНИЕ: показываем исходное название из протокола, а не текущее сопоставленное
            original_name = sample.get('original_name', sample['name'])
            st.write(f"**{original_name}**")
            if sample.get('steel_grade'):
                st.write(f"*Марка: {sample['steel_grade']}*")
        
        with col2:
            # Поле для фильтрации
            filter_text = st.text_input(
                "🔍 Фильтр (номер/буква)",
                placeholder="Например: 4, А, 12",
                key=f"filter_{i}"
            )
        
        with col3:
            # Определяем текущее сопоставление
            # ВАЖНОЕ ИСПРАВЛЕНИЕ: используем правильное название только если оно есть в списке корректных
            current_match = sample['name'] if sample['name'] in correct_names_list else "Не сопоставлен"
            
            # Фильтруем варианты на основе введенного текста
            # ИСПРАВЛЕНИЕ: используем name_matcher из analyzer
            filtered_options = analyzer.name_matcher._filter_correct_names(options, filter_text, correct_samples)
            
            # Выпадающий список с отфильтрованными вариантами
            selected = st.selectbox(
                f"Выберите правильное название для образца {i+1}",
                options=filtered_options,
                index=filtered_options.index(current_match) if current_match in filtered_options else 0,
                key=f"manual_match_{i}"
            )
            
            if selected != "Не сопоставлен":
                # Сохраняем сопоставление по исходному имени
                original_name = sample.get('original_name', sample['name'])
                manual_matches[original_name] = selected
    
    # Кнопка применения изменений
    if st.button("✅ Применить ручное сопоставление"):
        updated_samples = []
        
        for sample in editable_samples:
            original_name = sample.get('original_name', sample['name'])
            
            if original_name in manual_matches:
                correct_name = manual_matches[original_name]
                correct_sample = correct_names_dict[correct_name]
                
                # Обновляем sample - ВАЖНОЕ ИСПРАВЛЕНИЕ: сохраняем исходное название
                updated_sample = sample.copy()
                updated_sample['original_name'] = original_name  # Сохраняем оригинал из протокола
                updated_sample['name'] = correct_name            # Устанавливаем правильное название
                updated_sample['correct_number'] = correct_sample['number']
                updated_sample['manually_matched'] = True
                
                updated_samples.append(updated_sample)
            else:
                # Оставляем без изменений, но гарантируем что original_name сохранен
                sample['original_name'] = original_name
                sample['manually_matched'] = False
                updated_samples.append(sample)
        
        # Сохраняем результаты в session_state
        st.session_state.manually_matched_samples = updated_samples
        st.session_state.final_samples = updated_samples
        
        st.success(f"Ручное сопоставление применено! Обновлено {len(manual_matches)} образцов.")
        st.info("💡 Обновите страницу для отображения обновленных данных")
    
    return st.session_state.manually_matched_samples

def add_manual_steel_grade_correction(samples):
    """Интерфейс для ручного указания марки стали для нераспознанных образцов"""
    st.header("🔧 Ручное указание марки стали")
    
    unrecognized_samples = [s for s in samples if not s.get("steel_grade") or s["steel_grade"] == "Не распознана"]
    
    if not unrecognized_samples:
        st.success("✅ Все марки стали распознаны автоматически!")
        return samples
    
    st.warning(f"Не распознаны марки стали для {len(unrecognized_samples)} образцов:")
    
    # Создаем копию для редактирования
    updated_samples = samples.copy()
    
    for i, sample in enumerate(unrecognized_samples):
        # Находим индекс образца в основном списке
        sample_idx = updated_samples.index(sample)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.write(f"**{sample['name']}**")
            if sample.get('composition'):
                comp_text = ", ".join([f"{k}: {v}" for k, v in list(sample['composition'].items())[:5]])
                st.write(f"Состав: {comp_text}")
        
        with col2:
            # Предлагаем выбрать из известных марок или ввести вручную
            steel_options = ["12Х1МФ", "12Х18Н12Т", "сталь 20", "Ди82", "Ди59", "Другая..."]
            selected_grade = st.selectbox(
                f"Выберите марку стали для образца {i+1}",
                options=steel_options,
                key=f"steel_grade_{i}"
            )
            
            if selected_grade == "Другая...":
                custom_grade = st.text_input(
                    "Введите марку стали вручную:",
                    key=f"custom_steel_{i}"
                )
                if custom_grade:
                    updated_samples[sample_idx]["steel_grade"] = custom_grade
            else:
                updated_samples[sample_idx]["steel_grade"] = selected_grade
    
    if st.button("✅ Применить ручные исправления марок стали"):
        st.success("Марки стали обновлены!")
        st.info("💡 Обновите страницу для отображения обновленных данных")
        return updated_samples
    
    return samples

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

def set_font_times_new_roman(doc):
    """Устанавливает шрифт Times New Roman для всего документа"""
    # Устанавливаем шрифт для стилей
    styles = doc.styles
    for style in styles:
        if hasattr(style, 'font'):
            style.font.name = 'Times New Roman'
    
    # Устанавливаем шрифт для всех параграфов
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
    
    # Устанавливаем шрифт для всех таблиц
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'

def create_word_report(tables, samples, analyzer):
    """Создание Word отчета"""
    try:
        doc = Document()
        
        # Устанавливаем шрифт Times New Roman для всего документа
        set_font_times_new_roman(doc)
        
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
            
            # Добавление новых нормативов
            st.subheader("Добавить новую марку стали")
            
            new_grade = st.text_input("Марка стали")
            new_source = st.text_input("Нормативный документ", value="ТУ 14-3Р-55-2001")
            
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
        
        # Загрузка файла с правильными названиями
        st.subheader("1. Загрузите файл с правильными названиями образцов")
        correct_names_file = st.file_uploader(
            "Файл с правильными названиями (.docx)",
            type=["docx"],
            key="correct_names"
        )
        
        correct_samples = []
        if correct_names_file:
            # Показываем preview правильных названий
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
                            'Буква': sample['letter'] or 'н/д'
                        })
                    st.table(pd.DataFrame(preview_data))
        
        # Загрузка файлов протоколов
        st.subheader("2. Загрузите файлы протоколов химического анализа")
        uploaded_files = st.file_uploader(
            "Файлы протоколов (.docx)", 
            type=["docx"], 
            accept_multiple_files=True,
            key="protocol_files"
        )
        
        all_samples = []
        
        if uploaded_files:
            # Парсим все образцы из загруженных файлов
            for uploaded_file in uploaded_files:
                samples = analyzer.parse_protocol_file(uploaded_file.getvalue())
                all_samples.extend(samples)
            
            # Добавляем ручную коррекцию марок стали
            all_samples = add_manual_steel_grade_correction(all_samples)
            
            # Сопоставляем названия, если загружен файл с правильными названиями
            if correct_names_file and correct_samples:
                st.subheader("🔍 Автоматическое сопоставление названий образцов")
                all_samples, correct_samples_loaded = analyzer.match_sample_names(all_samples, correct_names_file)
                
                # Сохраняем автоматически сопоставленные образцы в session_state
                if st.session_state.final_samples is None:
                    st.session_state.final_samples = all_samples
                
                # Показываем интерфейс ручного сопоставления
                if st.session_state.manually_matched_samples is None:
                    st.session_state.manually_matched_samples = all_samples
                
                st.session_state.final_samples = add_manual_matching_interface(
                    st.session_state.manually_matched_samples, correct_samples_loaded, analyzer
                )
            else:
                # Если нет файла с правильными названиями, используем исходные образцы
                if st.session_state.final_samples is None:
                    st.session_state.final_samples = all_samples
            
            # Анализ и отображение результатов
            if st.session_state.final_samples:
                st.header("Результаты анализа")
                
                # Легенда цветов
                st.markdown("""
                **Легенда:**
                - <span style='background-color: #ffcccc; padding: 2px 5px; border-radius: 3px;'>🔴 Красный</span> - отклонение от норм
                - <span style='background-color: #f0f0f0; padding: 2px 5px; border-radius: 3px;'>⚪ Серый</span> - нормативные требования
                """, unsafe_allow_html=True)
                
                # Создание таблиц для отчета
                report_tables = analyzer.create_report_table_with_original_names(st.session_state.final_samples)
                
                if report_tables:
                    # Подготовка данных для экспорта
                    export_tables = {}
                    
                    for grade, table_data in report_tables.items():
                        st.subheader(f"Марка стали: {grade}")
                        
                        # Применяем стили к таблице
                        styled_table = apply_styling(table_data["data"], table_data["compliance"])
                        st.dataframe(styled_table, use_container_width=True, hide_index=True)
                        
                        # Сохраняем для экспорта
                        export_tables[grade] = table_data["data"]
                    
                    # Экспорт в Word
                    if st.button("📄 Экспорт в Word"):
                        create_word_report(export_tables, st.session_state.final_samples, analyzer)
                        st.success("Отчет готов к скачиванию!")
                else:
                    st.warning("Не удалось создать таблицы отчета. Проверьте данные образцов.")
                
                # Раздел с обработанными образцами (в самом конце)
                st.header("Обработанные образцы")
                for sample in st.session_state.final_samples:
                    with st.expander(f"📋 {sample['name']} - {sample['steel_grade']}"):
                        if 'original_name' in sample:
                            st.write(f"**Исходное название:** {sample['original_name']}")
                        if 'correct_number' in sample:
                            st.write(f"**Номер в списке:** {sample['correct_number']}")
                        st.write(f"**Марка стали:** {sample['steel_grade']}")
                        st.write("**Химический состав:**")
                        for element, value in sample['composition'].items():
                            st.write(f"- {element}: {value}")
                
    except Exception as e:
        st.error(f"Произошла ошибка при запуске приложения: {str(e)}")
        import traceback
        st.error(f"Детали ошибки: {traceback.format_exc()}")
        st.info("Попробуйте обновить страницу и загрузить файлы заново")

if __name__ == "__main__":
    main()
