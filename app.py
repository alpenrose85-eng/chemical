import streamlit as st
from docx import Document
import re
import pandas as pd

def extract_means_from_table(table):
    """Извлекает средние значения из одной таблицы."""
    means = {}
    current_headers = []

    for i, row in enumerate(table.rows):
        first_cell = row.cells[0].text.strip()
        first_cell_clean = re.sub(r'\s+', ' ', first_cell).strip()

        # Извлекаем заголовки: если строка содержит буквенные обозначения (C, Si и т.д.)
        if i == 0 or (i > 0 and not first_cell_clean.isdigit() and not first_cell_clean.startswith(('Среднее', '1', '2', '3', '±'))):
            potential_headers = []
            for cell in row.cells[1:]:
                h = re.sub(r'\s+', ' ', cell.text).replace('%', '').strip()
                if h and not h.replace('.', '').replace(',', '').isdigit() and not h.startswith('±'):
                    potential_headers.append(h)
            if potential_headers:
                current_headers = potential_headers

        # Обрабатываем строку "Среднее:" (но не погрешности)
        if "Среднее:" in first_cell_clean and not first_cell_clean.startswith("Среднее: ±"):
            for j, cell in enumerate(row.cells[1:], start=0):
                val_text = re.sub(r'\s+', ' ', cell.text).strip()
                if val_text and not val_text.startswith(('±', '-')) and j < len(current_headers):
                    try:
                        val = float(val_text.replace(',', '.'))
                        elem = current_headers[j]
                        means[elem] = val
                    except ValueError:
                        continue
    return means

def extract_all_samples(doc):
    """Проходит по документу и извлекает все образцы и их средние значения."""
    samples = {}
    current_sample = None

    for para in doc.paragraphs:
        if "Наименование образца :" in para.text:
            # Извлекаем имя образца
            current_sample = para.text.split("Наименование образца :")[-1].strip()

    # Но! В вашем файле таблицы идут сразу после строки с именем образца.
    # Поэтому будем искать таблицы и смотреть, какой образец был указан ПЕРЕД ней.
    current_sample = None
    for element in doc.element.body:
        if element.tag.endswith('p'):  # параграф
            para_text = element.text.strip()
            if "Наименование образца :" in para_text:
                current_sample = para_text.split("Наименование образца :")[-1].strip()
        elif element.tag.endswith('tbl'):  # таблица
            if current_sample:
                # Конвертируем element в объект Table
                from docx.table import Table
                table = Table(element, doc)
                means = extract_means_from_table(table)
                if means:
                    samples[current_sample] = means
                # После таблицы сбрасываем, чтобы не дублировать
                # (на самом деле, в вашем файле после таблицы идёт новый образец)
    return samples

# --- Streamlit UI ---
st.title("Извлечение химического состава")

uploaded_file = st.file_uploader("Загрузите файл .docx", type="docx")

if uploaded_file is not None:
    doc = Document(uploaded_file)
    samples = extract_all_samples(doc)

    if samples:
        df = pd.DataFrame(samples).T  # Транспонируем: образцы — строки, элементы — столбцы
        st.dataframe(df)
    else:
        st.warning("Не удалось извлечь данные. Проверьте структуру файла.")
