import streamlit as st
import pandas as pd
from docx import Document
import json
import os
from datetime import datetime
import io
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
from difflib import SequenceMatcher

class SampleNameMatcher:
    def __init__(self):
        self.surface_types = {
            'ЭПК': ['ЭПК'],
            'ШПП': ['ШПП'],
            'ПС КШ': ['ПС КШ', 'ПТ КШ', 'труба_ПТКМ', 'труба ПТКМ', 'ПТКМ', 'труба'],
            'КПП ВД': ['КПП ВД', 'ВД'],
            'КПП НД-1': ['КПП НД-1', 'КПП НД-I', 'НД-1', 'НД-I'],
            'КПП НД-2': ['КПП НД-2', 'КПП НД-II', 'НД-2', 'НД-II', 'КПП НД-IIст', 'НД-IIст']
        }

    def parse_correct_names(self, file_content):
        try:
            doc = Document(io.BytesIO(file_content))
            correct_names = []
            for table in doc.tables:
                for row in table.rows:
                    if len(row.cells) >= 2:
                        number_cell = row.cells[0].text.strip()
                        name_cell = row.cells[1].text.strip()
                        if number_cell and name_cell and number_cell.isdigit():
                            correct_names.append({
                                'number': int(number_cell),
                                'original': name_cell,
                                'surface_type': self.extract_surface_type(name_cell),
                                'tube_number': self.extract_tube_number(name_cell),
                                'letter': self.extract_letter(name_cell)
                            })
            if not correct_names:
                for paragraph in doc.paragraphs:
                    text = paragraph.text.strip()
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
            correct_names.sort(key=lambda x: x['number'])
            return correct_names
        except Exception as e:
            st.error(f"Ошибка при парсинге правильных названий: {str(e)}")
            return []

    def extract_tube_number(self, text):
        matches = re.findall(r'\((\d+)\)', text)
        if matches:
            return matches[0]
        matches = re.findall(r'\b(\d+)\b', text)
        if matches:
            return matches[-1]
        return None

    def extract_surface_type(self, name):
        for surface_type, patterns in self.surface_types.items():
            for pattern in patterns:
                if pattern in name:
                    return surface_type
        return None

    def extract_letter(self, name):
        matches = re.findall(r'\([^)]*([А-Г])\)', name)
        if matches:
            return matches[0]
        return None

    def extract_tube_number_from_protocol(self, sample_name):
        patterns = [
            r'тр\.\s*№?\s*(\d+)',
            r'тр\s*(\d+)',
            r'труба\s*(\d+)',
            r'тр\.\s*(\d+)',
            r'\((\d+)\)',
        ]
        for pattern in patterns:
            match = re.search(pattern, sample_name)
            if match:
                return match.group(1)
        numbers = re.findall(r'\b\d+\b', sample_name)
        if numbers:
            return max(numbers, key=lambda x: int(x))
        return None

    def parse_protocol_sample_name(self, sample_name):
        letter = None
        letter_map = {'НА': 'А', 'НБ': 'Б', 'НВ': 'В', 'НГ': 'Г', 'Н-Г': 'Г'}
        for prefix, mapped_letter in letter_map.items():
            if prefix in sample_name:
                letter = mapped_letter
                break
        if not letter:
            patterns = [r'Н[_\s\-]?([А-Г])', r'Н([А-Г])[_\s]', r'[_\s]Н([А-Г])']
            for pattern in patterns:
                matches = re.findall(pattern, sample_name)
                if matches:
                    letter = matches[0]
                    break
        tube_number = self.extract_tube_number_from_protocol(sample_name)
        surface_type = self.extract_surface_type(sample_name)
        return {
            'original': sample_name,
            'surface_type': surface_type,
            'tube_number': tube_number,
            'letter': letter
        }

    def match_samples(self, protocol_samples, correct_samples):
        matched = []
        unmatched = []
        used_correct = set()
        for protocol in protocol_samples:
            protocol_info = self.parse_protocol_sample_name(protocol['name'])
            found = False
            for correct in correct_samples:
                if correct['original'] in used_correct:
                    continue
                if (protocol_info['tube_number'] and correct['tube_number'] and
                    protocol_info['tube_number'] == correct['tube_number']):
                    matched.append((protocol, correct, "по номеру трубы"))
                    used_correct.add(correct['original'])
                    found = True
                    break
            if not found:
                unmatched.append(protocol)
        return matched, unmatched


class ChemicalAnalyzer:
    def __init__(self):
        self.load_standards()
        self.name_matcher = SampleNameMatcher()
        self.all_elements = ["C", "Si", "Mn", "P", "S", "Cr", "Mo", "Ni",
                             "Cu", "Al", "Co", "Nb", "Ti", "V", "W", "Fe"]

    def load_standards(self):
        self.standards = {
            "12Х1МФ": {
                "C": (0.10, 0.15), "Si": (0.17, 0.37), "Mn": (0.40, 0.70),
                "Cr": (0.90, 1.20), "Mo": (0.25, 0.35), "V": (0.15, 0.30),
                "Ni": (None, 0.25), "Cu": (None, 0.20), "S": (None, 0.025),
                "P": (None, 0.025), "source": "ТУ 14-3Р-55-2001"
            },
            "12Х18Н12Т": {
                "C": (None, 0.12), "Si": (None, 0.80), "Mn": (1.00, 2.00),
                "Cr": (17.00, 19.00), "Ni": (11.00, 13.00), "Ti": (None, 0.70),
                "Cu": (None, 0.30), "S": (None, 0.020), "P": (None, 0.035),
                "source": "ТУ 14-3Р-55-2001"
            },
            "20": {
                "C": (0.17, 0.24), "Si": (0.17, 0.37), "Mn": (0.35, 0.65),
                "Cr": (None, 0.25), "Ni": (None, 0.25), "Cu": (None, 0.30),
                "P": (None, 0.030), "S": (None, 0.025), "source": "ТУ 14-3Р-55-2001"
            },
            "Ди82": {
                "C": (0.08, 0.12), "Si": (None, 0.5), "Mn": (0.30, 0.60),
                "Cr": (8.60, 10.00), "Ni": (None, 0.70), "Mo": (0.60, 0.80),
                "V": (0.10, 0.20), "Nb": (0.10, 0.20), "Cu": (None, 0.30),
                "S": (None, 0.015), "P": (None, 0.03), "source": "ТУ 14-3Р-55-2001"
            },
            "Ди59": {
                "C": (0.06, 0.10), "Si": (1.8, 2.2), "Mn": (12.00, 13.50),
                "Cr": (11.50, 13.00), "Ni": (1.8, 2.5), "Nb": (0.60, 1.00),
                "Cu": (2.00, 2.50), "S": (None, 0.02), "P": (None, 0.03),
                "source": "ТУ 14-3Р-55-2001"
            }
        }

    def parse_protocol_file(self, file_content):
        try:
            doc = Document(io.BytesIO(file_content))
            samples = []
            current_sample = None
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if "Наименование образца:" in text:
                    sample_name = text.split("Наименование образца:")[1].strip()
                    current_sample = {
                        "name": sample_name,
                        "steel_grade": None,
                        "composition": {},
                        "original_name": sample_name
                    }
                    samples.append(current_sample)
                elif "соответствует марке стали:" in text:
                    if current_sample:
                        parts = text.split("марке стали:")
                        if len(parts) > 1:
                            grade_text = parts[1].strip()
                            grade_text = re.sub(r'[\*\,].*', '', grade_text).strip()
                            if grade_text:
                                current_sample["steel_grade"] = grade_text
            # парсинг таблиц
            table_index = 0
            for table in doc.tables:
                if table_index < len(samples):
                    composition = self.parse_composition_table(table)
                    samples[table_index]["composition"] = composition
                    table_index += 1
            return samples
        except Exception as e:
            st.error(f"Ошибка парсинга протокола: {str(e)}")
            return []

    def parse_composition_table(self, table):
        composition = {}
        try:
            table_data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)
            if len(table_data) < 13:
                return composition
            headers_row1 = table_data[0]
            values_row1 = table_data[5]
            headers_row2 = table_data[7]
            values_row2 = table_data[12]
            for i, header in enumerate(headers_row1):
                if header in self.all_elements and i < len(values_row1):
                    val_str = values_row1[i].replace(',', '.').replace(' ', '')
                    if '±' in val_str:
                        val_str = val_str.split('±')[0]
                    try:
                        composition[header] = float(val_str)
                    except:
                        pass
            for i, header in enumerate(headers_row2):
                if header in self.all_elements and i < len(values_row2):
                    val_str = values_row2[i].replace(',', '.').replace(' ', '')
                    if '±' in val_str:
                        val_str = val_str.split('±')[0]
                    try:
                        composition[header] = float(val_str)
                    except:
                        pass
            return composition
        except:
            return composition

    def apply_manual_matches(self, samples, correct_dict, manual_matches):
        updated = []
        for s in samples:
            new_s = s.copy()
            if s['original_name'] in manual_matches:
                correct_name = manual_matches[s['original_name']]
                if correct_name in correct_dict:
                    new_s['name'] = correct_name
                    new_s['correct_number'] = correct_dict[correct_name]['number']
                    new_s['manually_matched'] = True
                    new_s['automatically_matched'] = False
                else:
                    new_s['correct_number'] = None
                    new_s['manually_matched'] = False
            else:
                if not new_s.get('automatically_matched'):
                    new_s['correct_number'] = None
            updated.append(new_s)
        return updated

    def create_report_tables(self, samples):
        if not samples:
            return None
        matched_samples = [s for s in samples if s.get('correct_number') is not None]
        if not matched_samples:
            return None
        steel_grades = list(set(s["steel_grade"] for s in matched_samples if s["steel_grade"]))
        tables = {}
        for grade in steel_grades:
            if grade not in self.standards:
                continue
            grade_samples = [s for s in matched_samples if s["steel_grade"] == grade]
            standard = self.standards[grade]
            norm_elements = [e for e in standard.keys() if e != "source"]
            sorted_samples = sorted(grade_samples, key=lambda x: x.get('correct_number', float('inf')))
            data = []
            for idx, sample in enumerate(sorted_samples, 1):
                row = {"№": idx, "Образец": sample["name"]}
                for elem in norm_elements:
                    val = sample["composition"].get(elem)
                    if val is not None:
                        if elem in ["S","P"]:
                            row[elem] = f"{val:.3f}".replace('.',',')
                        else:
                            row[elem] = f"{val:.2f}".replace('.',',')
                    else:
                        row[elem] = "-"
                data.append(row)
            # строка требований
            req_row = {"№": "", "Образец": f"Требования ТУ для {grade}"}
            for elem in norm_elements:
                if elem in standard:
                    mn, mx = standard[elem]
                    if mn is not None and mx is not None:
                        if elem in ["S","P"]:
                            req_row[elem] = f"{mn:.3f}-{mx:.3f}".replace('.',',')
                        else:
                            req_row[elem] = f"{mn:.2f}-{mx:.2f}".replace('.',',')
                    elif mn is not None:
                        req_row[elem] = f"≥{mn:.2f}".replace('.',',')
                    elif mx is not None:
                        req_row[elem] = f"≤{mx:.2f}".replace('.',',')
                    else:
                        req_row[elem] = "н/д"
                else:
                    req_row[elem] = "-"
            data.append(req_row)
            tables[grade] = {"data": pd.DataFrame(data), "samples": sorted_samples}
        return tables

def create_word_report(samples, analyzer, report_tables):
    try:
        doc = Document()
        title = doc.add_heading('Протокол анализа химического состава', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        for grade, tbl in report_tables.items():
            doc.add_heading(f"Марка стали: {grade}", level=1)
            df = tbl["data"]
            word_table = doc.add_table(rows=len(df)+1, cols=len(df.columns))
            word_table.style = 'Table Grid'
            for j, col in enumerate(df.columns):
                word_table.cell(0, j).text = str(col)
            for i, row in df.iterrows():
                for j, col in enumerate(df.columns):
                    word_table.cell(i+1, j).text = str(row[col])
            doc.add_paragraph()
        doc.save("report.docx")
        with open("report.docx", "rb") as f:
            st.download_button("Скачать отчет", f, file_name=f"химический_анализ_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx")
    except Exception as e:
        st.error(f"Ошибка создания Word: {e}")

def main():
    st.set_page_config(page_title="Анализатор химсостава", layout="wide")
    st.title("🔬 Анализатор химического состава металла")
    analyzer = ChemicalAnalyzer()

    # Инициализация session_state
    if 'samples' not in st.session_state:
        st.session_state.samples = []
    if 'correct_samples' not in st.session_state:
        st.session_state.correct_samples = []
    if 'manual_matches' not in st.session_state:
        st.session_state.manual_matches = {}
    if 'parsed_samples' not in st.session_state:
        st.session_state.parsed_samples = []
    if 'correct_hash' not in st.session_state:
        st.session_state.correct_hash = None
    if 'protocols_hash' not in st.session_state:
        st.session_state.protocols_hash = None

    # Загрузка правильных названий
    correct_file = st.file_uploader("Файл с правильными названиями (.docx)", type=["docx"], key="correct")
    if correct_file:
        cur_hash = (correct_file.name, correct_file.size)
        if st.session_state.correct_hash != cur_hash:
            st.session_state.correct_samples = analyzer.name_matcher.parse_correct_names(correct_file.getvalue())
            st.session_state.correct_hash = cur_hash
            st.session_state.manual_matches = {}
            st.session_state.samples = []
        st.success(f"Загружено {len(st.session_state.correct_samples)} названий")

    # Загрузка протоколов
    protocol_files = st.file_uploader("Файлы протоколов (.docx)", type=["docx"], accept_multiple_files=True, key="protocols")
    if protocol_files:
        cur_hash = tuple((f.name, f.size) for f in protocol_files)
        if st.session_state.protocols_hash != cur_hash:
            all_samples = []
            for f in protocol_files:
                all_samples.extend(analyzer.parse_protocol_file(f.getvalue()))
            st.session_state.parsed_samples = all_samples
            st.session_state.protocols_hash = cur_hash
            st.session_state.samples = []
            st.session_state.manual_matches = {}
        st.success(f"Загружено {len(st.session_state.parsed_samples)} образцов")

    # Автоматическое сопоставление
    if st.session_state.correct_samples and st.session_state.parsed_samples and not st.session_state.samples:
        matched_pairs, unmatched = analyzer.name_matcher.match_samples(st.session_state.parsed_samples, st.session_state.correct_samples)
        matched_samples = []
        for proto, correct, _ in matched_pairs:
            new = proto.copy()
            new['original_name'] = proto['name']
            new['name'] = correct['original']
            new['correct_number'] = correct['number']
            new['automatically_matched'] = True
            matched_samples.append(new)
        for proto in unmatched:
            proto['original_name'] = proto['name']
            proto['correct_number'] = None
            proto['automatically_matched'] = False
            matched_samples.append(proto)
        st.session_state.samples = matched_samples

    # Ручное сопоставление
    if st.session_state.samples and st.session_state.correct_samples:
        st.header("Ручное сопоставление")
        correct_dict = {c['original']: c for c in st.session_state.correct_samples}
        for idx, sample in enumerate(st.session_state.samples):
            if sample.get('correct_number') is not None:
                continue
            orig = sample['original_name']
            options = ["Не сопоставлен"] + [c['original'] for c in st.session_state.correct_samples]
            current = st.session_state.manual_matches.get(orig, "Не сопоставлен")
            selected = st.selectbox(f"Образец: {orig}", options, index=options.index(current) if current in options else 0, key=f"manual_{idx}")
            if selected != "Не сопоставлен":
                st.session_state.manual_matches[orig] = selected
            elif orig in st.session_state.manual_matches:
                del st.session_state.manual_matches[orig]

        if st.button("Применить ручное сопоставление"):
            st.session_state.samples = analyzer.apply_manual_matches(st.session_state.samples, correct_dict, st.session_state.manual_matches)
            st.rerun()

    # Отображение результатов
    if st.session_state.samples:
        st.header("Результаты анализа")
        # Применяем ручные сопоставления перед созданием таблиц
        if st.session_state.correct_samples:
            correct_dict = {c['original']: c for c in st.session_state.correct_samples}
            updated = analyzer.apply_manual_matches(st.session_state.samples, correct_dict, st.session_state.manual_matches)
            st.session_state.samples = updated
        report_tables = analyzer.create_report_tables(st.session_state.samples)
        if report_tables:
            for grade, tbl in report_tables.items():
                st.subheader(f"Марка стали: {grade}")
                st.dataframe(tbl["data"], use_container_width=True, hide_index=True)
            if st.button("Создать Word отчет"):
                create_word_report(st.session_state.samples, analyzer, report_tables)
        else:
            st.warning("Нет сопоставленных образцов. Выполните ручное сопоставление.")

if __name__ == "__main__":
    main()
