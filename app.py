import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io

# ================================
# ИНИЦИАЛИЗАЦИЯ СЕССИИ + БАЗОВЫЕ НОРМЫ
# ================================
if "steel_norms" not in st.session_state:
    st.session_state.steel_norms = {
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
        "10Х13Г12БС2Н2Д2 (ДИ59)": {
            "C": (0.06, 0.10),
            "Si": (1.8, 2.2),
            "Mn": (12.00, 13.50),
            "Cr": (11.50, 13.00),
            "Ni": (1.8, 2.5),
            "Nb": (0.60, 1.00),
            "Cu": (2.00, 2.50),
            "S": (None, 0.02),
            "P": (None, 0.03)
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
        },
        "20": {
            "C": (0.17, 0.24),
            "Si": (0.17, 0.37),
            "Mn": (0.35, 0.65),
            "Cr": (None, 0.25),
            "Ni": (None, 0.25),
            "Cu": (None, 0.30),
            "P": (None, 0.03),
            "S": (None, 0.025)
        },
        "10Х9МФБ (ДИ82)": {
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
            "P": (None, 0.03)
        }
    }

# ================================
# ПАРСЕР ПРОТОКОЛА
# ================================
def parse_protocol_docx(file):
    doc = Document(file)
    full_text = "\n".join([p.text for p in doc.paragraphs])
    tables = doc.tables

    samples = []
    sample_blocks = re.split(r"Наименование образца\s*:", full_text)
    content_blocks = sample_blocks[1:]

    table_iter = iter(tables)
    for block in content_blocks:
        lines = block.strip().split("\n")
        sample_name = lines[0].strip()

        steel_match = re.search(r"марке стали:\s*([А-Яа-я0-9Хх\(\)\s\-]+)", block)
        if steel_match:
            steel_grade = steel_match.group(1).strip()
            if "," in steel_grade:
                steel_grade = steel_grade.split(",")[0].strip()
        else:
            steel_grade = "Неизвестно"

        notes = ""
        if "с учетом допустимых отклонений" in block:
            notes = "с учетом допустимых отклонений и погрешности измерения"

        try:
            table1 = next(table_iter)
            table2 = next(table_iter)
        except StopIteration:
            break

        all_elements = {}
        for tbl in [table1, table2]:
            headers = []
            for cell in tbl.rows[0].cells:
                txt = cell.text.strip().replace("\n", "").replace("%", "").strip()
                if txt and txt not in ["", "1", "2", "3"]:
                    headers.append(txt)

            mean_row = None
            unc_row = None
            for row in tbl.rows:
                first = row.cells[0].text.strip()
                if first == "Среднее:":
                    mean_row = row
                elif "±" in first:
                    unc_row = row

            if mean_row and unc_row:
                for j, elem in enumerate(headers):
                    if j + 1 < len(mean_row.cells):
                        try:
                            mean_val = float(mean_row.cells[j + 1].text.replace(",", ".").strip())
                            unc_text = unc_row.cells[j + 1].text.replace("±", "").replace(",", ".").strip()
                            unc_val = float(unc_text)
                            all_elements[elem] = {"mean": mean_val, "unc": unc_val}
                        except (ValueError, IndexError):
                            continue

        samples.append({
            "name": sample_name,
            "steel": steel_grade,
            "elements": all_elements,
            "notes": notes
        })

    return samples

# ================================
# ФУНКЦИИ АНАЛИЗА
# ================================
def evaluate_status(value, unc, norm_min, norm_max):
    low = value - unc
    high = value + unc
    if norm_min is not None and high < norm_min:
        return "🔴"
    if norm_max is not None and low > norm_max:
        return "🔴"
    if (norm_min is not None and low < norm_min <= high) or (norm_max is not None and low <= norm_max < high):
        return "🟡"
    return ""

def format_value(val, elem):
    return f"{val:.3f}" if elem in ["S", "P"] else f"{val:.2f}"

# ================================
# ГЕНЕРАЦИЯ WORD-ОТЧЁТА
# ================================
def create_word_report(all_samples, steel_norms):
    doc = Document()
    doc.add_heading('Отчёт по химическому составу металла', 0)
    doc.add_paragraph('Источник: загруженные протоколы лаборатории')

    # Собираем все нормируемые элементы
    norm_elements = set()
    for norms in steel_norms.values():
        norm_elements.update(norms.keys())
    norm_elements = sorted(norm_elements, key=lambda x: ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"].index(x) if x in ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"] else 999)

    # Таблица
    cols = ["Образец"] + norm_elements
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    for i, c in enumerate(cols):
        hdr[i].text = c

    # Заполняем строки
    for sample in all_samples:
        steel = sample["steel"]
        norms = steel_norms.get(steel, {})
        if not norms:
            continue
        row_cells = table.add_row().cells
        row_cells[0].text = sample["name"]
        for j, elem in enumerate(norm_elements, start=1):
            if elem in sample["elements"]:
                val = sample["elements"][elem]["mean"]
                unc = sample["elements"][elem]["unc"]
                nmin, nmax = norms.get(elem, (None, None))
                status = evaluate_status(val, unc, nmin, nmax)
                txt = format_value(val, elem)
                row_cells[j].text = txt
                if status == "🔴":
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:fill'), 'ffcccc')
                    row_cells[j]._element.get_or_add_tcPr().append(shading)
                elif status == "🟡":
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:fill'), 'fffacd')
                    row_cells[j]._element.get_or_add_tcPr().append(shading)
            else:
                row_cells[j].text = "–"

    # Детальный анализ
    doc.add_heading('Детальный анализ', level=1)
    for sample in all_samples:
        steel = sample["steel"]
        norms = steel_norms.get(steel, {})
        if not norms:
            continue
        doc.add_heading(f"{sample['name']} (сталь {steel})", level=2)
        for elem, (nmin, nmax) in norms.items():
            if elem in sample["elements"]:
                val = sample["elements"][elem]["mean"]
                unc = sample["elements"][elem]["unc"]
                status = evaluate_status(val, unc, nmin, nmax)
                interval = f"[{val - unc:.3f}; {val + unc:.3f}]"
                if status == "🔴":
                    doc.add_paragraph(f"🔴 {elem}: {format_value(val, elem)} ± {unc:.3f} → {interval} — ВНЕ норм")
                elif status == "🟡":
                    doc.add_paragraph(f"🟡 {elem}: {format_value(val, elem)} ± {unc:.3f} → {interval} — пограничное значение")
                else:
                    doc.add_paragraph(f"✅ {elem}: {format_value(val, elem)} ± {unc:.3f} → {interval} — в пределах норм")
        if sample["notes"]:
            doc.add_paragraph(f"📌 Примечание: {sample['notes']}")

    # Легенда
    doc.add_heading('Легенда', level=1)
    doc.add_paragraph("🔴 — явное несоответствие нормам\n🟡 — пограничное значение\n✅ — соответствие нормам")

    return doc

# ================================
# ИНТЕРФЕЙС УПРАВЛЕНИЯ МАРКАМИ
# ================================
st.sidebar.title("Управление марками сталей")
steel_to_edit = st.sidebar.selectbox(
    "Выберите марку для редактирования или введите новую",
    options=[""] + list(st.session_state.steel_norms.keys()),
    format_func=lambda x: x if x else "➕ Новая марка"
)
new_steel_name = st.sidebar.text_input("Название марки", value=steel_to_edit or "")
if new_steel_name:
    current_norms = st.session_state.steel_norms.get(new_steel_name, {})
    elements = ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P", "Al", "Co", "Nb", "Ti", "W"]
    edited_norms = {}
    for elem in elements:
        col1, col2 = st.sidebar.columns(2)
        min_val = col1.text_input(f"{elem} min", value=str(current_norms.get(elem, (None, None))[0] or ""))
        max_val = col2.text_input(f"{elem} max", value=str(current_norms.get(elem, (None, None))[1] or ""))
        min_f = float(min_val) if min_val else None
        max_f = float(max_val) if max_val else None
        if min_f is not None or max_f is not None:
            edited_norms[elem] = (min_f, max_f)
    if st.sidebar.button("💾 Сохранить нормы"):
        st.session_state.steel_norms[new_steel_name] = edited_norms
        st.sidebar.success(f"Нормы для {new_steel_name} сохранены!")

# ================================
# ОСНОВНОЙ ИНТЕРФЕЙС
# ================================
st.title("Анализ химического состава металла")
uploaded_files = st.file_uploader("Загрузите протоколы (.docx)", type=["docx"], accept_multiple_files=True)

all_samples = []
if uploaded_files:
    for f in uploaded_files:
        samples = parse_protocol_docx(f)
        all_samples.extend(samples)

if not all_samples:
    st.info("Загрузите хотя бы один протокол в формате .docx")
    st.stop()

# Подготовка данных для таблицы
rows = []
for sample in all_samples:
    steel = sample["steel"]
    norms = st.session_state.steel_norms.get(steel, {})
    if not norms:
        continue
    row = {"Образец": sample["name"], "Марка": steel}
    for elem in norms:
        if elem in sample["elements"]:
            val = sample["elements"][elem]["mean"]
            unc = sample["elements"][elem]["unc"]
            row[elem] = val
            row[f"{elem}_unc"] = unc
        else:
            row[elem] = None
    rows.append(row)

if not rows:
    st.error("Нет данных для обработки по известным маркам.")
    st.stop()

# HTML-таблица
norm_elements = set()
for norms in st.session_state.steel_norms.values():
    norm_elements.update(norms.keys())
norm_elements = sorted(norm_elements, key=lambda x: ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"].index(x) if x in ["C", "Si", "Mn", "Cr", "Ni", "Mo", "V", "Cu", "S", "P"] else 999)

df_display = pd.DataFrame(rows)
cols_order = ["Образец"] + [e for e in norm_elements if e in df_display.columns]
df_display = df_display[cols_order]

html_rows = ["<tr>" + "".join(f"<th>{c}</th>" for c in cols_order) + "</tr>"]
for _, r in df_display.iterrows():
    row_html = f"<td>{r['Образец']}</td>"
    steel = next((s["steel"] for s in all_samples if s["name"] == r["Образец"]), "Неизвестно")
    norms = st.session_state.steel_norms.get(steel, {})
    for elem in cols_order[1:]:
        val = r.get(elem, None)
        if pd.isna(val):
            row_html += "<td>–</td>"
        else:
            unc = r.get(f"{elem}_unc", 0)
            nmin, nmax = norms.get(elem, (None, None))
            status = evaluate_status(val, unc, nmin, nmax)
            txt = format_value(val, elem)
            if status == "🔴":
                row_html += f'<td style="background-color:#ffcccc">{txt}</td>'
            elif status == "🟡":
                row_html += f'<td style="background-color:#fffacd">{txt}</td>'
            else:
                row_html += f"<td>{txt}</td>"
    html_rows.append("<tr>" + row_html + "</tr>")

html_table = f'<table border="1" style="border-collapse:collapse;">{"".join(html_rows)}</table>'
st.markdown("### Сводная таблица (копируйте в Word):")
st.markdown(html_table, unsafe_allow_html=True)

# Кнопка экспорта
if st.button("📥 Скачать полный отчёт в Word (.docx)"):
    doc = create_word_report(all_samples, st.session_state.steel_norms)
    bio = io.BytesIO()
    doc.save(bio)
    st.download_button(
        label="Скачать отчёт.docx",
        data=bio.getvalue(),
        file_name="Отчёт_химсостав_металла.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# Детальный анализ
st.subheader("Детальный анализ")
for sample in all_samples:
    steel = sample["steel"]
    norms = st.session_state.steel_norms.get(steel, {})
    if not norms:
        continue
    with st.expander(f"🔍 {sample['name']} (сталь {steel})"):
        for elem, (nmin, nmax) in norms.items():
            if elem in sample["elements"]:
                val = sample["elements"][elem]["mean"]
                unc = sample["elements"][elem]["unc"]
                status = evaluate_status(val, unc, nmin, nmax)
                interval = f"[{val - unc:.3f}; {val + unc:.3f}]"
                if status == "🔴":
                    st.error(f"{elem}: {format_value(val, elem)} ± {unc:.3f} → {interval} — ВНЕ норм")
                elif status == "🟡":
                    st.warning(f"{elem}: {format_value(val, elem)} ± {unc:.3f} → {interval} — пограничное")
                else:
                    st.success(f"{elem}: {format_value(val, elem)} ± {unc:.3f} → {interval} — в норме")
        if sample["notes"]:
            st.info(f"📌 Примечание: {sample['notes']}")