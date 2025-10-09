import streamlit as st
import pandas as pd
import re

def parse_protocol_text(text):
    """Парсит протокол из текстового представления"""
    
    samples = []
    current_sample = {}
    
    lines = text.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        # Начало нового образца
        if 'Наименование образца:' in line:
            if current_sample:
                samples.append(current_sample)
            current_sample = {'Наименование образца': line.split('Наименование образца:')[-1].strip()}
        
        # Марка стали
        elif 'Химический состав металла образца' in line and 'марке стали:' in line:
            steel_grade = line.split('марке стали:')[-1].strip()
            current_sample['Марка стали'] = steel_grade
        
        # Таблица с химическим составом
        elif 'Среднее:' in line and i + 1 < len(lines):
            # Ищем числовые значения в текущей и следующих строках
            numbers = re.findall(r'\d+\.\d+', line + ' ' + lines[i+1] if i+1 < len(lines) else line)
            
            if len(numbers) >= 8:
                elements = ['C', 'Si', 'Mn', 'P', 'S', 'Cr', 'Mo', 'Ni']
                for idx, elem in enumerate(elements):
                    if idx < len(numbers):
                        current_sample[elem] = float(numbers[idx])
            
            # Пропускаем дополнительные строки таблицы
            i += 2
        
        i += 1
    
    if current_sample:
        samples.append(current_sample)
    
    return samples

def main_simple():
    st.title("Анализатор протокола химического состава")
    
    st.write("""
    Если DOCX парсер не работает, попробуйте этот вариант:
    1. Откройте ваш DOCX файл
    2. Скопируйте весь текст (Ctrl+A, Ctrl+C)
    3. Вставьте в поле ниже
    """)
    
    text_input = st.text_area("Вставьте текст протокола сюда:", height=400)
    
    if st.button("Анализировать текст") and text_input:
        samples = parse_protocol_text(text_input)
        
        if samples:
            df = pd.DataFrame(samples)
            st.success(f"Найдено {len(samples)} образцов")
            st.dataframe(df, use_container_width=True)
            
            # Скачивание
            csv = df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                "Скачать CSV",
                csv,
                "chemical_composition.csv",
                "text/csv"
            )
        else:
            st.error("Не удалось найти данные образцов в тексте")

# Запустите main_simple() если основной не работает
if __name__ == "__main__":
    main()  # или main_simple() для текстового варианта
