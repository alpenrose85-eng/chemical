import streamlit as st
import pandas as pd
from docx import Document
import re
import io
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt

# Нормы для 12Х1МФ (ТУ 14-3Р-55-2001)
NORMS_12X1MF = {
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
}

def parse_protocol_docx(file):
    doc = Document(file)
    full_text = "\n".join([p.text for p in doc.paragraphs])
    # Разделяем по заголовкам "Наименование образца"
    blocks = re.split(r"Наименование образца\s*:", full_text)[1:]
    tables = doc.tables
    samples = []
    table_idx = 0

    for block in blocks:
        lines = [line.strip() for line in block.split("\n") if line.strip()]
        if not lines:
            continue
        sample_name = lines[0]

        # 🔧 Извлекаем марку стали: "12Х1МФ" (игнорируем запятые и примечания)
        steel_match = re.search(r"марке стали:\s*([А-Яа-я0-9Хх]+)", block)
        steel = steel_match.group(1).strip() if steel_match else "Неизвестно"

        notes = "с учетом допустимых отклонений" if "с учетом допустимых отклонений" in block else ""

        if table_idx + 1 >= len(tables):
            break

        def extract_means_from_table(table):
            headers = []
            for cell in table.rows[0].cells[1:]:  # Пропускаем первую пустую ячейку
                h = cell.text.strip().replace("\n", "").replace("%", "").strip()
                if h:
                    headers.append(h)

            means = {}
            for row in table.rows:
                if row.cells[0].text.strip() == "Среднее:":
                    for j, elem in enumerate(headers):
                        if j + 1 < len(row.cells):
                            try:
                                val = float(row.cells[j + 1].text.replace(",", ".").strip())
                                means[elem] = val
                            except:
                                pass
                    break  # Берём только ПЕРВУЮ строку "Среднее:"
            return means

        # Парсим две таблицы подряд
        means1 = extract_means_from_table(tables[table_idx])
        means2 = extract_means_from_table(tables[table_idx + 1])
        table_idx += 2

        # Объединяем значения из двух таблиц
        all_means = {**means1, **means2}
        samples.append({
            "name": sample_name,
            "steel": steel,
            "elements": all_means,
            "notes": notes
        })

    return samples

def evaluate_status_simple(value, norm_min, norm_max):
    """Сравнение без учёта погрешности"""
    if norm_min is not None and value < norm_min:
        return "🔴"
    if norm_max is not None and value > norm_max:
        return "🔴"
    return ""  # Соответствует

def format_value(val, elem):
    if elem in ["S", "P"]:
        return f"{val:.3f}".replace(".", ",")
    elif elem == "Cu":
        return f"{val:.2f}".replace(".", ",")  # Округляем медь до сотых
    else:
        return f"{val:.2f}".replace(".", ",")

def format_norm(norm_min, norm_max):
    if norm_min is None:
        return f"≤{norm_max:.3f}".replace(".", ",")
    elif norm_max is None:
        return f"≥{norm_min:.3f}".replace(".", ",")
    else:
        return f"{norm_min:.2f}–{norm_max:.2f}".replace(".", ",")

# ================================
# Генерация Word-отчёта
# ================================
def create_word_report(samples):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    doc.add_heading('Отчёт по химическому составу металла', 0)
    doc.add_paragraph('Источник: Протокол № 27/05 от 26.05.2025, ОАО «ВТИ»')

    # Элементы для таблицы — только те, что есть в нормах
    cols = ["Образец", "C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"]
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'

    # Заголовок
    hdr = table.rows[0].cells
    for i, c in enumerate(cols):
        hdr[i].text = c
        hdr[i].paragraphs[0].runs[0].font.name = 'Times New Roman'

    # Данные по образцам
    for sample in samples:
        row = table.add_row().cells
        row[0].text = sample["name"]
        row[0].paragraphs[0].runs[0].font.name = 'Times New Roman'
        for j, elem in enumerate(cols[1:], start=1):
            val = sample["elements"].get(elem)
            if val is not None:
                txt = format_value(val, elem)
                row[j].text = txt
                row[j].paragraphs[0].runs[0].font.name = 'Times New Roman'
                # Цвет
                nmin, nmax = NORMS_12X1MF[elem]
                if evaluate_status_simple(val, nmin, nmax) == "🔴":
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:fill'), 'ffcccc')
                    row[j]._element.get_or_add_tcPr().append(shading)
            else:
                row[j].text = "–"
                row[j].paragraphs[0].runs[0].font.name = 'Times New Roman'

    # Строка требований
    req_row = table.add_row().cells
    req_row[0].text = "Требования ТУ 14-3Р-55-2001 [3] для стали марки 12Х1МФ"
    req_row[0].paragraphs[0].runs[0].font.name = 'Times New Roman'
    for j, elem in enumerate(cols[1:], start=1):
        nmin, nmax = NORMS_12X1MF[elem]
        req_row[j].text = format_norm(nmin, nmax)
        req_row[j].paragraphs[0].runs[0].font.name = 'Times New Roman'

    # Анализ
    doc.add_heading('Выводы', level=1)
    for sample in samples:
        doc.add_heading(sample["name"], level=2)
        for elem in cols[1:]:
            val = sample["elements"].get(elem)
            if val is not None:
                nmin, nmax = NORMS_12X1MF[elem]
                status = evaluate_status_simple(val, nmin, nmax)
                if status == "🔴":
                    doc.add_paragraph(f"🔴 {elem} = {format_value(val, elem)} — не соответствует норме ({format_norm(nmin, nmax)})")
                else:
                    doc.add_paragraph(f"✅ {elem} = {format_value(val, elem)} — соответствует норме")
        if sample["notes"]:
            doc.add_paragraph(f"📌 Примечание: {sample['notes']}")

    doc.add_heading('Легенда', level=1)
    doc.add_paragraph("🔴 — несоответствие нормам\n✅ — соответствие нормам")

    return doc

# ================================
# Streamlit UI
# ================================
st.set_page_config(page_title="Анализ химсостава", layout="wide")
st.title("Анализ химического состава металла")

uploaded = st.file_uploader("Загрузите протокол (.docx)", type=["docx"])

if uploaded:
    try:
        samples = parse_protocol_docx(uploaded)
        st.success(f"Загружено образцов: {len(samples)}")

        # Подготовка данных для таблицы
        data = []
        for s in samples:
            row = {"Образец": s["name"]}
            for elem in ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"]:
                val = s["elements"].get(elem)
                if val is not None:
                    row[elem] = format_value(val, elem)
                else:
                    row[elem] = "–"
            data.append(row)

        df = pd.DataFrame(data)
        cols_order = ["Образец", "C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"]
        df = df[cols_order]

        # HTML-таблица
        html_rows = ["<tr>" + "".join(f"<th style='font-family: Times New Roman;'>{c}</th>" for c in cols_order) + "</tr>"]
        for _, r in df.iterrows():
            row_html = f"<td style='font-family: Times New Roman;'>{r['Образец']}</td>"
            for elem in cols_order[1:]:
                val_str = r[elem]
                val_num = None
                try:
                    val_num = float(val_str.replace(",", "."))
                except:
                    pass

                if val_num is not None:
                    nmin, nmax = NORMS_12X1MF[elem]
                    status = evaluate_status_simple(val_num, nmin, nmax)
                    if status == "🔴":
                        row_html += f'<td style="background-color:#ffcccc; font-family: Times New Roman;">{val_str}</td>'
                    else:
                        row_html += f'<td style="font-family: Times New Roman;">{val_str}</td>'
                else:
                    row_html += f'<td style="font-family: Times New Roman;">{val_str}</td>'
            html_rows.append("<tr>" + row_html + "</tr>")

        # Строка требований
        req_cells = ["Требования ТУ 14-3Р-55-2001 [3] для стали марки 12Х1МФ"]
        for elem in cols_order[1:]:
            nmin, nmax = NORMS_12X1MF[elem]
            req_cells.append(format_norm(nmin, nmax))
        req_row = "<tr>" + "".join(f"<td style='font-family: Times New Roman;'>{c}</td>" for c in req_cells) + "</tr>"
        html_rows.append(req_row)

        html_table = f'<table border="1" style="border-collapse:collapse; font-family: Times New Roman;">{"".join(html_rows)}</table>'
        st.markdown("### Сводная таблица (копируйте в Word):")
        st.markdown(html_table, unsafe_allow_html=True)

        # Экспорт в Word
        if st.button("📥 Скачать отчёт в Word"):
            doc = create_word_report(samples)
            bio = io.BytesIO()
            doc.save(bio)
            st.download_button(
                label="Скачать отчёт.docx",
                data=bio.getvalue(),
                file_name="Отчёт_химсостав_12Х1МФ.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        # Детальный анализ
        st.subheader("Детальный анализ")
        for s in samples:
            with st.expander(f"🔍 {s['name']}"):
                for elem in ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"]:
                    val = s["elements"].get(elem)
                    if val is not None:
                        nmin, nmax = NORMS_12X1MF[elem]
                        status = evaluate_status_simple(val, nmin, nmax)
                        if status == "🔴":
                            st.error(f"{elem} = {format_value(val, elem)} — не соответствует норме ({format_norm(nmin, nmax)})")
                        else:
                            st.success(f"{elem} = {format_value(val, elem)} — соответствует норме")
                if s["notes"]:
                    st.info(f"📌 Примечание: {s['notes']}")

    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
else:
    st.info("Загрузите файл протокола в формате .docx")
