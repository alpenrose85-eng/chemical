import streamlit as st
import pandas as pd
import re
from docx import Document
import io

def extract_sample_data(doc_text):
    """Извлекает данные о образцах и химическом составе из текста документа"""
    
    # Регулярные выражения для извлечения данных
    sample_name_pattern = r'\[Наименование образца:\]\{\.underline\} (.+)'
    steel_grade_pattern = r'\[Химический состав металла образца (?:соответствует|близок) марке стали:\]\{\.underline\} (.+)'
    
    # Разделяем текст на секции по образцам
    samples_sections = re.split(r'\[Наименование образца:\]\{\.underline\}', doc_text)
    
    samples_data = []
    
    for section in samples_sections[1:]:  # Пропускаем первую секцию (заголовок)
        sample_data = {}
        
        # Извлекаем название образца
        sample_name_match = re.search(r'^ (.+)', section)
        if sample_name_match:
            sample_data['Наименование образца'] = sample_name_match.group(1).strip()
        
        # Извлекаем марку стали
        steel_grade_match = re.search(steel_grade_pattern, section)
        if steel_grade_match:
            sample_data['Марка стали'] = steel_grade_match.group(1).strip()
        
        # Ищем таблицу с химическим составом
        table_match = re.search(r'Среднее:\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*', section)
        
        if table_match:
            elements = ['C', 'Si', 'Mn', 'P', 'S', 'Cr', 'Mo', 'Ni']
            for i, element in enumerate(elements, 1):
                sample_data[element] = float(table_match.group(i))
        
        # Ищем вторую часть таблицы
        table2_match = re.search(r'\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*\s*\*\*([\d\.]+)\*\*', section)
        
        if table2_match:
            elements2 = ['Cu', 'Al', 'Co', 'Nb', 'Ti', 'V', 'W', 'Fe']
            for i, element in enumerate(elements2, 1):
                sample_data[element] = float(table2_match.group(i))
        
        if sample_data:
            samples_data.append(sample_data)
    
    return samples_data

def main():
    st.title("Анализ химического состава металла")
    st.subheader("Извлечение данных из протокола испытаний")
    
    uploaded_file = st.file_uploader("Загрузите файл .docx с протоколом испытаний", type="docx")
    
    if uploaded_file is not None:
        try:
            # Читаем DOCX файл
            doc = Document(uploaded_file)
            
            # Извлекаем весь текст
            full_text = ""
            for paragraph in doc.paragraphs:
                full_text += paragraph.text + "\n"
            
            # Для таблиц в docx
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        full_text += cell.text + " | "
                    full_text += "\n"
                full_text += "\n"
            
            # Извлекаем данные
            samples_data = extract_sample_data(full_text)
            
            if samples_data:
                # Создаем DataFrame
                df = pd.DataFrame(samples_data)
                
                # Отображаем результаты
                st.success(f"Успешно извлечено данных для {len(samples_data)} образцов")
                
                # Показываем таблицу
                st.dataframe(df, use_container_width=True)
                
                # Показываем статистику
                st.subheader("Статистика по химическим элементам")
                numeric_cols = df.select_dtypes(include=['float64']).columns
                if len(numeric_cols) > 0:
                    st.dataframe(df[numeric_cols].describe(), use_container_width=True)
                
                # Кнопка для скачивания данных
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="Скачать данные в CSV",
                    data=csv,
                    file_name="химический_состав_металла.csv",
                    mime="text/csv"
                )
            else:
                st.error("Не удалось извлечь данные из файла. Проверьте формат документа.")
                
        except Exception as e:
            st.error(f"Ошибка при обработке файла: {str(e)}")
    
    else:
        st.info("Пожалуйста, загрузите файл формата .docx с протоколом испытаний")

if __name__ == "__main__":
    main()
