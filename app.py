import streamlit as st
import pandas as pd
import re
from docx import Document

def safe_extract_text_from_docx(uploaded_file):
    """Безопасно извлекает текст из DOCX файла"""
    try:
        doc = Document(uploaded_file)
        full_text = ""
        
        # Извлекаем текст из параграфов
        for paragraph in doc.paragraphs:
            full_text += paragraph.text + "\n"
        
        # Извлекаем текст из таблиц
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text += cell.text + " | "
                full_text += "\n"
            full_text += "\n"
        
        return full_text
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {str(e)}")
        return ""

def parse_samples_simple(text):
    """Простой парсер для извлечения данных образцов"""
    samples = []
    
    # Разделяем текст на секции по образцам
    sections = re.split(r'Наименование образца:', text)
    
    for section in sections[1:]:  # Пропускаем первую секцию (заголовок)
        sample_data = {}
        
        # Извлекаем название образца (первая строка после разделителя)
        first_line = section.split('\n')[0].strip()
        if first_line:
            sample_data['Наименование образца'] = first_line
        
        # Ищем марку стали
        steel_match = re.search(r'12Х[^\s]+', section)
        if steel_match:
            sample_data['Марка стали'] = steel_match.group(0)
        
        # Ищем все числовые значения в формате 0.123
        numbers = re.findall(r'\b\d+\.\d+\b', section)
        
        # Обычно в каждом образце есть несколько измерений и средние значения
        # Берем последние 16 чисел (предполагая, что это средние значения)
        if len(numbers) >= 16:
            elements = ['C', 'Si', 'Mn', 'P', 'S', 'Cr', 'Mo', 'Ni', 
                       'Cu', 'Al', 'Co', 'Nb', 'Ti', 'V', 'W', 'Fe']
            
            # Берем последние 16 чисел как средние значения
            avg_numbers = numbers[-16:]
            
            for i, element in enumerate(elements):
                try:
                    sample_data[element] = float(avg_numbers[i])
                except (ValueError, IndexError):
                    sample_data[element] = None
        
        if sample_data.get('Наименование образца'):
            samples.append(sample_data)
    
    return samples

def main():
    st.set_page_config(page_title="Анализатор химического состава", layout="wide")
    
    st.title("🔬 Анализатор химического состава металла")
    st.markdown("---")
    
    uploaded_file = st.file_uploader("Загрузите файл протокола испытаний (.docx)", type="docx")
    
    if uploaded_file is not None:
        try:
            with st.spinner("Чтение файла..."):
                text_content = safe_extract_text_from_docx(uploaded_file)
            
            if not text_content:
                st.error("Не удалось извлечь текст из файла")
                return
            
            # Показываем превью текста
            with st.expander("Просмотр содержимого файла (первые 1000 символов)"):
                st.text(text_content[:1000])
            
            with st.spinner("Анализ данных..."):
                samples_data = parse_samples_simple(text_content)
            
            if samples_data:
                st.success(f"✅ Успешно обработано {len(samples_data)} образцов")
                
                # Создаем DataFrame
                df = pd.DataFrame(samples_data)
                
                # Отображаем таблицу
                st.subheader("Результаты анализа")
                st.dataframe(df, use_container_width=True)
                
                # Статистика
                st.subheader("Статистика по элементам")
                numeric_cols = df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    st.dataframe(df[numeric_cols].describe(), use_container_width=True)
                
                # Скачивание
                st.subheader("Экспорт данных")
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    "📥 Скачать данные в CSV",
                    data=csv,
                    file_name="химический_состав.csv",
                    mime="text/csv"
                )
                
            else:
                st.error("Не удалось найти данные образцов в файле")
                st.info("""
                **Рекомендации:**
                - Убедитесь, что файл содержит таблицы с химическим составом
                - Проверьте, что формат файла соответствует примеру
                - Попробуйте преобразовать документ в текстовый формат
                """)
                
        except Exception as e:
            st.error(f"Произошла ошибка: {str(e)}")
            st.info("Попробуйте альтернативный вариант ниже")
    
    # Альтернативный метод - загрузка текста
    st.markdown("---")
    st.subheader("Альтернативный метод: вставка текста")
    
    text_input = st.text_area(
        "Если загрузка файла не работает, скопируйте и вставьте текст протокола:",
        height=300,
        placeholder="Вставьте сюда текст из вашего документа..."
    )
    
    if st.button("Анализировать текст") and text_input:
        with st.spinner("Анализ текста..."):
            samples_data = parse_samples_simple(text_input)
        
        if samples_data:
            st.success(f"✅ Найдено {len(samples_data)} образцов")
            df = pd.DataFrame(samples_data)
            st.dataframe(df, use_container_width=True)
            
            csv = df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                "📥 Скачать CSV",
                data=csv,
                file_name="chemical_composition.csv",
                mime="text/csv"
            )
        else:
            st.error("Не удалось найти данные в тексте")

if __name__ == "__main__":
    main()
