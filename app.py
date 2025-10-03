def match_sample_names(self, samples, correct_names_file):
    """Сопоставление названий образцов с правильными названиями - УЛУЧШЕННАЯ ВЕРСИЯ"""
    if not correct_names_file:
        return samples, []
    
    correct_samples = self.name_matcher.parse_correct_names(correct_names_file.getvalue())
    
    if not correct_samples:
        st.warning("Не удалось загрузить правильные названия образцов")
        return samples, []
    
    matched_samples = []
    unmatched_samples = []
    
    for sample in samples:
        # ВАЖНО: сохраняем оригинальное название до любых изменений
        original_protocol_name = sample['name']
        
        protocol_sample_info = self.name_matcher.parse_protocol_sample_name(original_protocol_name)
        best_match = self.name_matcher.find_best_match(protocol_sample_info, correct_samples)
        
        if best_match:
            # Создаем копию образца с исправленным названием и номером
            corrected_sample = sample.copy()
            corrected_sample['original_name'] = original_protocol_name  # Сохраняем оригинальное название
            corrected_sample['name'] = best_match['original']           # Заменяем на правильное
            corrected_sample['correct_number'] = best_match['number']   # Сохраняем номер для сортировки
            corrected_sample['automatically_matched'] = True
            matched_samples.append(corrected_sample)
        else:
            # Если совпадение не найдено, оставляем оригинальное название
            sample['original_name'] = original_protocol_name  # Сохраняем для информации
            sample['correct_number'] = None                   # Нет номера для сортировки
            sample['automatically_matched'] = False
            unmatched_samples.append(sample)
    
    # Выводим информацию о сопоставлении
    if matched_samples:
        st.success(f"Успешно сопоставлено {len(matched_samples)} образцов")
        
        with st.expander("📋 Просмотр сопоставленных образцов"):
            match_data = []
            for sample in matched_samples:
                match_data.append({
                    'Номер': sample['correct_number'],
                    'Исходное название': sample['original_name'],
                    'Правильное название': sample['name']
                })
            # Сортируем по номеру
            match_data.sort(key=lambda x: x['Номер'])
            st.table(pd.DataFrame(match_data))
    
    if unmatched_samples:
        st.warning(f"Не удалось сопоставить {len(unmatched_samples)} образцов")
        
        with st.expander("⚠️ Просмотр несопоставленных образцов"):
            unmatched_data = []
            for sample in unmatched_samples:
                unmatched_data.append({
                    'Образец': sample['original_name'],
                    'Марка стали': sample['steel_grade']
                })
            st.table(pd.DataFrame(unmatched_data))
    
    # Сортируем сопоставленные образцы по номеру, несопоставленные оставляем в конце
    matched_samples.sort(key=lambda x: x['correct_number'])
    return matched_samples + unmatched_samples, correct_samples
