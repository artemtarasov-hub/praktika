import streamlit as st
import pandas as pd
import io

# --- 1. ЗАГРУЗКА ДАННЫХ (Этот блок обязателен) ---
st.title("Финансовый Анализ (2022-2024)")

uploaded_file = st.file_uploader("Загрузите Excel-файл с отчетностью", type=["xlsx", "xls"])

if uploaded_file:
    # Загружаем данные.
    # ВАЖНО: Здесь я использую df_balance и df_results, как вы просили.
    # Убедитесь, что в Excel листы называются именно так, или исправьте 'Balance' и 'Results' на ваши названия листов.
    try:
        df_balance = pd.read_excel(uploaded_file, sheet_name='Balance') # Или 'Баланс', проверьте имя листа!
        df_results = pd.read_excel(uploaded_file, sheet_name='Results') # Или 'ОФР'
        
        # Приводим индексы к строковому типу, чтобы искать по кодам '1600' и т.д.
        # Предполагаем, что коды строк находятся в первой колонке или индексе.
        # Если коды строк в колонке 'Code', раскомментируйте строчку ниже:
        # df_balance.set_index('Code', inplace=True)
        # df_results.set_index('Code', inplace=True)
        
    except Exception as e:
        st.error(f"Ошибка чтения листов Excel. Проверьте названия листов. Детали: {e}")
        st.stop()

    # Убеждаемся, что колонки называются '2022', '2023', '2024'. 
    # Если в файле они называются иначе, нужно их переименовать здесь.
    # Пример: df_balance.columns = ['Code', 'Name', '2024', '2023', '2022'] - подстройте под ваш файл.

    # --- 2. ФУНКЦИИ АНАЛИЗА (Задание 1 и 2) ---
    
    def analyze_dynamics_3_years(df, entity_name="Показатель"):
        analysis = df.copy()
        # Выбираем только нужные колонки с годами
        cols = ['2022', '2023', '2024']
        
        # Проверка наличия колонок
        if not all(col in analysis.columns for col in cols):
            st.error(f"В таблице {entity_name} нет колонок 2022, 2023 или 2024. Проверьте файл.")
            return analysis

        # Абсолютные изменения
        analysis['Изм. 2023/2022'] = analysis['2023'] - analysis['2022']
        analysis['Изм. 2024/2023'] = analysis['2024'] - analysis['2023']
        
        # Темпы роста (%)
        analysis['Темп роста 2023/2022 (%)'] = (analysis['2023'] / analysis['2022'] * 100).fillna(0)
        analysis['Темп роста 2024/2023 (%)'] = (analysis['2024'] / analysis['2023'] * 100).fillna(0)
        
        return analysis

    def calculate_ratios_3_years(df_bal, df_res):
        years = ['2022', '2023', '2024']
        ratios = pd.DataFrame(index=years)
        
        # ВАЖНО: Убедитесь, что коды строк (1600, 1300...) являются индексами dataframe
        # Если нет, используйте: val = df_bal.loc[df_bal['Код'] == 1600, year].values[0]
        
        for year in years:
            try:
                # Пример упрощенного доступа (если индекс = код строки)
                assets = df_bal.loc[1600, year] 
                equity = df_bal.loc[1300, year]
                curr_liab = df_bal.loc[1500, year]
                curr_assets = df_bal.loc[1200, year]
                
                revenue = df_res.loc[2110, year]
                net_profit = df_res.loc[2400, year]

                ratios.loc[year, 'Рентабельность активов (ROA), %'] = (net_profit / assets * 100)
                ratios.loc[year, 'Коэф. текущей ликвидности (CR)'] = curr_assets / curr_liab
                ratios.loc[year, 'Коэф. автономии'] = equity / assets
            except KeyError:
                st.warning(f"Не найдены коды строк для года {year}. Проверьте структуру Excel.")
        
        return ratios.T

    # --- 3. ВЫПОЛНЕНИЕ И ВЫВОД (Где была ошибка) ---

    st.subheader("Задание 1: Горизонтальный и вертикальный анализ")
    
    # Теперь df_balance точно существует
    df_balance_analysis = analyze_dynamics_3_years(df_balance, entity_name="Бухгалтерский баланс")
    st.write("Анализ Баланса:")
    st.dataframe(df_balance_analysis)

    st.subheader("Задание 2: Финансовые коэффициенты")
    df_ratios = calculate_ratios_3_years(df_balance, df_results)
    st.write("Коэффициенты (2022-2024):")
    st.dataframe(df_ratios)

else:
    st.info("Пожалуйста, загрузите файл для начала работы.")
