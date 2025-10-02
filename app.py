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

# Элементы для каждой стали (только те, что проверяются по нормам)
ELEMENTS_BY_STEEL = {
    "12Х1МФ": ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"],
    "12Х18Н12Т": ["C", "Si", "Mn", "Cr", "Ni", "Ti", "Cu", "S", "P"]
}

def extract_means_ignore_errors(table):
    """Извлекает только первую строку 'Среднее:' из таблицы."""
    if len(table.rows) < 2:
        return {}
    headers = []
    for cell in table.rows[0].cells[1:]:
        h = cell.text.strip().replace("\n", "").replace("%", "").strip()
        if h:
            headers.append(h)
    means = {}
    found = False
    for row in table.rows:
        first_cell = row.cells[0].text.strip()
        if first_cell == "Среднее:":
            if found:
                # Вторая строка "Среднее:" — погрешности, пропускаем
                break
            for j, elem in enumerate(headers):
                if j + 1 < len(row.cells):
                    try:
                        val_text = row.cells[j + 1].text.strip().replace(",", ".").replace(" ", "")
                        if val_text and val_text not in ("-", "±"):
                            val = float(val_text)
                            means[elem] = val
                    except Exception:
                        pass
            found = True
    return means

def parse_protocol_docx(file):
    doc = Document(file)
    samples = []

    # Собираем все элементы документа: параграфы и таблицы
    content = []
    for elem in doc.element.body:
        tag = elem.tag
        if tag.endswith('p'):
            content.append(('paragraph', elem.text))
        elif tag.endswith('tbl'):
            # Создаём временную таблицу
            temp_doc = Document()
            new_table = temp_doc.add_table(0, 0)
            new_table._element = elem
            content.append(('table', new_table))

    current_sample_name = None
    current_steel = None
    current_notes = ""
    pending_tables = []

    for typ, val in content:
        if typ == 'paragraph':
            text = val.strip()
            if "Наименование образца" in text:
                # Сохраняем предыдущий образец, если есть две таблицы
                if current_sample_name and len(pending_tables) >= 2:
                    means1 = extract_means_ignore_errors(pending_tables[0])
                    means2 = extract_means_ignore_errors(pending_tables[1])
                    all_means = {**means1, **means2}
                    samples.append({
                        "name": current_sample_name,
                        "steel": current_steel,
                        "elements": all_means,
                        "notes": current_notes
                    })
                    pending_tables = []

                # Новый образец
                match = re.search(r"Наименование образца\s*[:\s]*(.+)", text)
                current_sample_name = match.group(1).strip() if match else "Неизвестно"
                current_steel = None
                current_notes = "с учетом допустимых отклонений" if "с учетом допустимых отклонений" in text else ""

                # Попытка найти марку стали в этом параграфе
                steel_match = re.search(r"марке стали:\s*([А-Яа-я0-9ХхМФТ]+)", text)
                if steel_match:
                    current_steel = steel_match.group(1).strip()

        elif typ == 'table':
            if current_sample_name:
                pending_tables.append(val)
                if len(pending_tables) == 2:
                    means1 = extract_means_ignore_errors(pending_tables[0])
                    means2 = extract_means_ignore_errors(pending_tables[1])
                    all_means = {**means1, **means2}
                    samples.append({
                        "name": current_sample_name,
                        "steel": current_steel,
                        "elements": all_means,
                        "notes": current_notes
                    })
                    # Сбрасываем после сохранения
                    current_sample_name = None
                    pending_tables = []

    # Обработка последнего образца (на случай, если документ заканчивается таблицами)
    if current_sample_name and len(pending_tables) >= 2:
        means1 = extract_means_ignore_errors(pending_tables[0])
        means2 = extract_means_ignore_errors(pending_tables[1])
        all_means = {**means1, **means2}
        samples.append({
            "name": current_sample_name,
            "steel": current_steel,
            "elements": all_means,
            "notes": current_notes
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

    st.success(f"Обработано образцов: {len(all_samples)}")

    # Группируем по маркам сталей
    steel_groups = {}
    for s in all_samples:
        steel = s["steel"]
        if steel not in steel_groups:
            steel_groups[steel] = []
        steel_groups[steel].append(s)

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
        bio = io.BytesIO()
        doc = create_word_report_for_steel(group_samples, steel)
        doc.save(bio)
        st.download_button(
            label=f"📥 Скачать отчёт для стали {steel}",
            data=bio.getvalue(),
            file_name=f"Отчёт_химсостав_{steel}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"download_{steel}"
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
