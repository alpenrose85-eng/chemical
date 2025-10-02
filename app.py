import streamlit as st
from docx import Document
import re
import pandas as pd

def extract_means_from_table(table):
    """Извлекает средние значения из одной таблицы (оба блока)."""
    means = {}
    headers = []

    for i, row in enumerate(table.rows):
        first_cell = row.cells[0].text.strip()
        first_cell_clean = re.sub(r'\s+', ' ', first_cell).strip()

        # Определяем заголовки: строка, где первая ячейка пустая, а остальные — буквы
        if not first_cell_clean or first_cell_clean == "-":
            headers = []
            for cell in row.cells[1:]:
                h = re.sub(r'\s+', ' ', cell.text).replace('%', '').strip()
                if h and not h.replace('.', '').replace(',', '').isdigit() and not h.startswith(('±', '1', '2', '3', 'Среднее')):
                    headers.append(h)

        # Извлекаем строку "Среднее:" (но не погрешности)
        if "Среднее:" in first_cell_clean and not first_cell_clean.startswith("Среднее: ±"):
            for j, cell in enumerate(row.cells[1:], start=0):
                val_text = re.sub(r'\s+', ' ', cell.text).strip()
                if val_text and not val_text.startswith(('±', '-')) and j < len(headers):
                    try:
                        val = float(val_text.replace(',', '.'))
                        elem = headers[j]
                        means[elem] = val
                    except ValueError:
                        continue
    return means

def extract_all_samples(doc):
    # Шаг 1: извлекаем все имена образцов из параграфов
    sample_names = []
    for para in doc.paragraphs:
        if "Наименование образца :" in para.text:
            name = para.text.split("Наименование образца :")[-1].strip()
            if name:
                sample_names.append(name)

    # Шаг 2: извлекаем все таблицы
    tables = doc.tables

    # Шаг 3: связываем по порядку
    samples = {}
    for i, table in enumerate(tables):
        if i < len(sample_names):
            means = extract_means_from_table(table)
            if means:
                samples[sample_names[i]] = means
        else:
            break  # больше нет имён

    return samples

# --- Streamlit UI ---
st.title("Извлечение химического состава")

uploaded_file = st.file_uploader("Загрузите файл .docx", type="docx")

if uploaded_file is not None:
    doc = Document(uploaded_file)
    samples = extract_all_samples(doc)

    if samples:
        df = pd.DataFrame(samples).T
        st.dataframe(df)
    else:
        st.warning("Не удалось извлечь данные. Проверьте структуру файла.")
