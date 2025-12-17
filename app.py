import pandas as pd

# Предполагаем, что данные уже загружены в df_balance и df_results (или ваши исходные названия)
# Убедитесь, что в ваших исходных файлах есть колонки '2022', '2023', '2024'

years = ['2022', '2023', '2024']

# --- ЗАДАНИЕ 1: Структурно-динамический анализ (Горизонтальный и Вертикальный) ---

def analyze_dynamics_3_years(df, entity_name="Показатель"):
    """
    Функция рассчитывает изменения за 3 года (2022-2024).
    """
    analysis = df.copy()
    
    # 1. Абсолютные изменения
    analysis['Изм. 2023/2022 (абс.)'] = analysis['2023'] - analysis['2022']
    analysis['Изм. 2024/2023 (абс.)'] = analysis['2024'] - analysis['2023']
    
    # 2. Темпы роста (%)
    # Используем fillna(0) для избежания деления на ноль, если данные могут быть пустыми
    analysis['Темп роста 2023/2022 (%)'] = (analysis['2023'] / analysis['2022'] * 100).fillna(0)
    analysis['Темп роста 2024/2023 (%)'] = (analysis['2024'] / analysis['2023'] * 100).fillna(0)
    
    # Опционально: Общий тренд 2024 к 2022
    analysis['Темп роста 2024/2022 (%)'] = (analysis['2024'] / analysis['2022'] * 100).fillna(0)

    print(f"--- Анализ динамики: {entity_name} (2022-2024) ---")
    return analysis

# Пример вызова для Баланса (Задание 1)
# Возвращаем название таблицы как "Бухгалтерский баланс" (или как было изначально)
df_balance_analysis = analyze_dynamics_3_years(df_balance, entity_name="Бухгалтерский баланс")
display(df_balance_analysis)


# --- ЗАДАНИЕ 2: Расчет финансовых коэффициентов (за 3 года) ---

def calculate_ratios_3_years(df_bal, df_res):
    """
    Расчет коэффициентов для каждого года из списка years.
    """
    ratios = pd.DataFrame(index=years)
    
    for year in years:
        # Пример получения данных (здесь нужно подставить ваши реальные коды строк)
        # Допустим: Активы (1600), Капитал (1300), Обязательства (1500+1400), Выручка (2110)
        
        # Внимание: Убедитесь, что индексация или поиск строк настроен верно под ваш файл
        # Ниже приведен пример логики доступа к данным
        assets = df_bal.loc['1600', year] 
        equity = df_bal.loc['1300', year]
        current_liabilities = df_bal.loc['1500', year]
        current_assets = df_bal.loc['1200', year]
        revenue = df_res.loc['2110', year]
        net_profit = df_res.loc['2400', year]

        # 1. Рентабельность (Profitability)
        ratios.loc[year, 'Рентабельность активов (ROA), %'] = (net_profit / assets * 100)
        ratios.loc[year, 'Рентабельность продаж (ROS), %'] = (net_profit / revenue * 100)
        
        # 2. Ликвидность (Liquidity)
        ratios.loc[year, 'Коэф. текущей ликвидности (CR)'] = current_assets / current_liabilities
        
        # 3. Финансовая устойчивость
        ratios.loc[year, 'Коэф. автономии'] = equity / assets

    return ratios.T # Транспонируем, чтобы года были столбцами, как в исходных таблицах

# Пример вызова для Коэффициентов (Задание 2)
print("--- Таблица: Финансовые коэффициенты (2022-2024) ---")
df_ratios = calculate_ratios_3_years(df_balance, df_results)
display(df_ratios)
