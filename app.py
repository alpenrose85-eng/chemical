import streamlit as st
from docx import Document as DocxDocument  # фабричная функция для открытия
from docx.document import Document  # класс для isinstance
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
import re
import pandas as pd

def iter_block_items(parent):
    """Итератор по элементам документа (параграфы и таблицы в порядке следования)."""
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    else:
        raise ValueError("parent must be a Document instance")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_means_from_table(table):
    means = {}
    current_headers = []

    for i, row in enumerate(table.rows):
        first_cell = row.cells[0].text.strip()
        first_cell_clean = re.sub(r'\s+', ' ', first_cell).strip()

        # Проверка: строка может быть заголовочной, если первая ячейка пустая или "-"
        if not first_cell_clean or first_cell_clean == "-":
            potential_headers = []
            for cell in row.cells[1:]:
                h = re.sub(r'\s+', ' ', cell.text).replace('%', '').strip()
                # Пропускаем числа, погрешности, номера измерений
                if h and not h.replace('.', '').replace(',', '').isdigit() and not any(h.startswith(x) for x in ['±', '1', '2', '3', 'Среднее']):
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
    samples = {}
    current_sample = None

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            para_text = block.text.strip()
            if "Наименование образца :" in para_text:
                current_sample = para_text.split("Наименование образца :")[-1].strip()
        elif isinstance(block, Table):
            if current_sample:
                means = extract_means_from_table(block)
                if means:
                    samples[current_sample] = means
    return samples

# --- Streamlit UI ---
st.title("Извлечение химического состава")

uploaded_file = st.file_uploader("Загрузите файл .docx", type="docx")

if uploaded_file is not None:
    # Загружаем документ
    doc = DocxDocument(uploaded_file)
    samples = extract_all_samples(doc)

    if samples:
        df = pd.DataFrame(samples).T
        st.dataframe(df)
    else:
        st.warning("Не удалось извлечь данные. Проверьте структуру файла.")
