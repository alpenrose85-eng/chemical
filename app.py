import streamlit as st
import pandas as pd
from docx import Document
import re
import io
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt

# Нормы для сталей
NORMS = {
    "12Х1МФ": {
        "C": (0.10, 0.15),
        "Si": (0.17, 0.27),
        "Mn": (0.40, 0.70),
        "Cr": (0.90, 1.20),
        "Ni": (None, 0.25),
        "Mo": (0.25, 0.35),
        "V": (0.15, 0.30),
        "Cu": (None, 0.20),
        "S": (None, 0.025),
        "P": (None, 0.025)
    },
    "12Х18Н12Т": {
        "C": (None, 0.12),
        "Si": (None, 0.80),
        "Mn": (1.00, 2.00),
        "Cr": (17.00, 19.00),
        "Ni": (11.00, 13.00),
        "Ti": (None, 0.7),
        "Cu": (None, 0.30),
        "S": (None, 0.020),
        "P": (None, 0.035)
    }
}

# Элементы для каждой стали
ELEMENTS_BY_STEEL = {
    "12Х1МФ": ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"],
    "12Х18Н12Т": ["C", "Si", "Mn", "Cr", "Ni", "Ti", "Cu", "S", "P"]
}

def parse_protocol_docx(file):
# ================================
# Исправленная функция parse_protocol_docx
# ================================

from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.document import Document as DocxDocumentClass

def iter_block_items(parent):
    """Итератор по элементам документа (параграфы и таблицы в порядке следования)."""
    if isinstance(parent, DocxDocumentClass):
        parent_elm = parent.element.body
    else:
        raise ValueError("parent must be a Document instance")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_means_from_table(table):
    """Извлекает средние значения из одной таблицы."""
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

def parse_protocol_docx(file):
    doc = Document(file)
    samples = []
    current_sample_name = None
    current_steel = None
    table_buffer = []  # Буфер для хранения таблиц текущего образца

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            para_text = block.text.strip()
            if "Наименование образца :" in para_text:
                # Сохраняем предыдущий образец, если есть данные
                if current_sample_name and len(table_buffer) >= 2:
                    means1 = extract_means_from_table(table_buffer[0])
                    means2 = extract_means_from_table(table_buffer[1])
                    all_means = {**means1, **means2}
                    notes = "с учетом допустимых отклонений" if "с учетом допустимых отклонений" in para_text else ""
                    samples.append({
                        "name": current_sample_name,
                        "steel": current_steel,
                        "elements": all_means,
                        "notes": notes
                    })
                # Сбрасываем буфер
                table_buffer = []
                # Начинаем новый образец
                current_sample_name = para_text.split("Наименование образца :")[-1].strip()
                # Ищем марку стали в этом же параграфе
                steel_match = re.search(r"марке стали:\s*([А-Яа-я0-9Хх]+)", para_text)
                if steel_match:
                    current_steel = steel_match.group(1).strip()
                else:
                    current_steel = "Неизвестно"
            elif current_sample_name and "с учетом допустимых отклонений" in para_text:
                # Просто запоминаем, что для этого образца есть примечание
                # (мы добавим его при создании образца)
                pass
        elif isinstance(block, Table):
            if current_sample_name:
                table_buffer.append(block)
                # Если накопили 2 таблицы — обрабатываем
                if len(table_buffer) == 2:
                    means1 = extract_means_from_table(table_buffer[0])
                    means2 = extract_means_from_table(table_buffer[1])
                    all_means = {**means1, **means2}
                    # Проверяем, есть ли примечание в параграфах, относящихся к этому образцу
                    notes = ""
                    # Простой способ: проверяем, содержится ли текст "с учетом допустимых отклонений" в параграфе с именем образца
                    # или в следующих параграфах до следующего образца
                    # Для простоты, если в параграфе с именем образца есть примечание
                    if "с учетом допустимых отклонений" in current_sample_name:
                        notes = "с учетом допустимых отклонений"
                    else:
                        # Можно сделать более сложную проверку, но пока так
                        pass
                    samples.append({
                        "name": current_sample_name,
                        "steel": current_steel,
                        "elements": all_means,
                        "notes": notes
                    })
                    # Сбрасываем для следующего образца
                    current_sample_name = None
                    current_steel = None
                    table_buffer = []

    # Обработка последнего образца
    if current_sample_name and len(table_buffer) >= 2:
        means1 = extract_means_from_table(table_buffer[0])
        means2 = extract_means_from_table(table_buffer[1])
        all_means = {**means1, **means2}
        notes = "с учетом допустимых отклонений" if "с учетом допустимых отклонений" in current_sample_name else ""
        samples.append({
            "name": current_sample_name,
            "steel": current_steel,
            "elements": all_means,
            "notes": notes
        })

    return samples

def evaluate_status(value, norm_min, norm_max):
    if norm_min is not None and value < norm_min:
        return "🔴"
    if norm_max is not None and value > norm_max:
        return "🔴"
    return ""

def format_value(val, elem):
    if elem in ["S", "P"]:
        return f"{val:.3f}".replace(".", ",")
    else:
        return f"{val:.2f}".replace(".", ",")

def format_norm(norm_min, norm_max):
    if norm_min is None:
        return f"≤{norm_max:.2f}".replace(".", ",")
    elif norm_max is None:
        return f"≥{norm_min:.2f}".replace(".", ",")
    else:
        return f"{norm_min:.2f}–{norm_max:.2f}".replace(".", ",")

# ================================
# Генерация Word-отчёта для одной стали
# ================================
def create_word_report_for_steel(samples, steel):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    doc.add_heading(f'Отчёт по химическому составу металла — сталь {steel}', 0)
    doc.add_paragraph('Источник: Протокол № 46/10 от 02.10.2025, ОАО «ВТИ»')

    elements = ELEMENTS_BY_STEEL.get(steel, [])
    if not elements:
        doc.add_paragraph("Для этой стали нет нормативов")
        return doc

    cols = ["Образец"] + elements
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'

    # Заголовок
    for i, c in enumerate(cols):
        table.rows[0].cells[i].text = c
        table.rows[0].cells[i].paragraphs[0].runs[0].font.name = 'Times New Roman'

    # Данные
    for sample in samples:
        if sample["steel"] != steel:
            continue
        row = table.add_row().cells
        row[0].text = sample["name"]
        row[0].paragraphs[0].runs[0].font.name = 'Times New Roman'
        for j, elem in enumerate(elements, start=1):
            val = sample["elements"].get(elem)
            cell = row[j]
            if val is not None:
                txt = format_value(val, elem)
                cell.text = txt
                status = evaluate_status(val, *NORMS[steel][elem])
                if status == "🔴":
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:fill'), 'ffcccc')
                    cell._element.get_or_add_tcPr().append(shading)
            else:
                cell.text = "–"
            cell.paragraphs[0].runs[0].font.name = 'Times New Roman'

    # Строка требований
    req_row = table.add_row().cells
    req_row[0].text = f"Требования ТУ 14-3Р-55-2001 [3] для стали марки {steel}"
    req_row[0].paragraphs[0].runs[0].font.name = 'Times New Roman'
    for j, elem in enumerate(elements, start=1):
        nmin, nmax = NORMS[steel][elem]
        req_row[j].text = format_norm(nmin, nmax)
        req_row[j].paragraphs[0].runs[0].font.name = 'Times New Roman'

    # Выводы
    doc.add_heading('Выводы', level=1)
    for s in samples:
        if s["steel"] != steel:
            continue
        doc.add_heading(s["name"], level=2)
        for elem in elements:
            val = s["elements"].get(elem)
            if val is not None:
                nmin, nmax = NORMS[steel][elem]
                status = evaluate_status(val, nmin, nmax)
                if status == "🔴":
                    doc.add_paragraph(f"🔴 {elem} = {format_value(val, elem)} — не соответствует норме ({format_norm(nmin, nmax)})")
                else:
                    doc.add_paragraph(f"✅ {elem} = {format_value(val, elem)} — соответствует норме")
        if s["notes"]:
            doc.add_paragraph(f"📌 Примечание: {s['notes']}")

    doc.add_heading('Легенда', level=1)
    doc.add_paragraph("🔴 — несоответствие нормам\n✅ — соответствие нормам")

    return doc

# ================================
# Streamlit UI
# ================================
st.set_page_config(page_title="Анализ химсостава", layout="wide")
st.title("Анализ химического состава металла")

uploaded_files = st.file_uploader("Загрузите протоколы (.docx)", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    all_samples = []
    for file in uploaded_files:
        try:
            samples = parse_protocol_docx(file)
            all_samples.extend(samples)
        except Exception as e:
            st.error(f"Ошибка при обработке файла {file.name}: {e}")

    if not all_samples:
        st.info("Не удалось обработать ни один файл")
        st.stop()

    # Группируем образцы по маркам сталей
    steel_groups = {}
    for s in all_samples:
        steel = s["steel"]
        if steel not in steel_groups:
            steel_groups[steel] = []
        steel_groups[steel].append(s)

    # Показываем таблицы по каждой стали
    for steel, group_samples in steel_groups.items():
        st.subheader(f"Сталь: {steel}")
        elements = ELEMENTS_BY_STEEL.get(steel, [])
        if not elements:
            st.warning("Для этой стали нет нормативов")
            continue

        # Подготовка данных
        data = []
        for s in group_samples:
            row = {"Образец": s["name"]}
            for elem in elements:
                val = s["elements"].get(elem)
                row[elem] = format_value(val, elem) if val is not None else "–"
            data.append(row)

        df = pd.DataFrame(data)
        cols_order = ["Образец"] + elements
        df = df[cols_order]

        # HTML-таблица
        html_rows = ["<tr>" + "".join(f"<th style='font-family: Times New Roman;'>{c}</th>" for c in cols_order) + "</tr>"]
        for _, r in df.iterrows():
            row_html = f"<td style='font-family: Times New Roman;'>{r['Образец']}</td>"
            for elem in elements:
                val_str = r[elem]
                if val_str == "–":
                    row_html += f'<td style="font-family: Times New Roman;">{val_str}</td>'
                else:
                    try:
                        val_num = float(val_str.replace(",", "."))
                        nmin, nmax = NORMS[steel][elem]
                        status = evaluate_status(val_num, nmin, nmax)
                        if status == "🔴":
                            row_html += f'<td style="background-color:#ffcccc; font-family: Times New Roman;">{val_str}</td>'
                        else:
                            row_html += f'<td style="font-family: Times New Roman;">{val_str}</td>'
                    except:
                        row_html += f'<td style="font-family: Times New Roman;">{val_str}</td>'
            html_rows.append("<tr>" + row_html + "</tr>")

        # Строка требований
        req_cells = [f"Требования ТУ 14-3Р-55-2001 [3] для стали марки {steel}"]
        for elem in elements:
            nmin, nmax = NORMS[steel][elem]
            req_cells.append(format_norm(nmin, nmax))
        req_row = "<tr>" + "".join(f"<td style='font-family: Times New Roman;'>{c}</td>" for c in req_cells) + "</tr>"
        html_rows.append(req_row)

        html_table = f'<table border="1" style="border-collapse:collapse; font-family: Times New Roman;">{"".join(html_rows)}</table>'
        st.markdown("##### Сводная таблица (копируйте в Word):")
        st.markdown(html_table, unsafe_allow_html=True)

        # Кнопка экспорта
        if st.button(f"📥 Скачать отчёт для стали {steel}", key=f"download_{steel}"):
            doc = create_word_report_for_steel(group_samples, steel)
            bio = io.BytesIO()
            doc.save(bio)
            st.download_button(
                label=f"Скачать отчёт_{steel}.docx",
                data=bio.getvalue(),
                file_name=f"Отчёт_химсостав_{steel}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # Детальный анализ
    st.subheader("Детальный анализ")
    for s in all_samples:
        with st.expander(f"🔍 {s['name']} ({s['steel']})"):
            elements = ELEMENTS_BY_STEEL.get(s["steel"], [])
            for elem in elements:
                val = s["elements"].get(elem)
                if val is not None:
                    nmin, nmax = NORMS[s["steel"]][elem]
                    status = evaluate_status(val, nmin, nmax)
                    if status == "🔴":
                        st.error(f"{elem} = {format_value(val, elem)} — не соответствует норме ({format_norm(nmin, nmax)})")
                    else:
                        st.success(f"{elem} = {format_value(val, elem)} — соответствует норме")
            if s["notes"]:
                st.info(f"📌 Примечание: {s['notes']}")

else:
    st.info("Загрузите файлы протоколов в формате .docx")
