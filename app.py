import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from openai import OpenAI

# ==========================================
# 1. НАСТРОЙКИ СТРАНИЦЫ
# ==========================================
st.set_page_config(
    page_title="Практическая работа №1 (Smart Analysis)",
    page_icon="🧠",
    layout="wide"
)

st.title("📊 Практическая работа: Комплексный анализ (Авто-определение года)")

# ==========================================
# 2. ФУНКЦИИ (ПАРСИНГ И ЛОГИКА)
# ==========================================

def clean_number(val):
    """Превращает строку вида '(16 671)' в число -16671."""
    if pd.isna(val): return 0
    s = str(val).strip()
    if s == '-' or s == '': return 0
    
    # Убираем пробелы
    s = s.replace(' ', '').replace('\xa0', '')
    
    # Обработка скобок как минуса
    if '(' in s and ')' in s:
        s = s.replace('(', '-').replace(')', '')
    
    try:
        return float(s)
    except:
        return 0

def get_year_and_data_from_excel(file_obj):
    """
    Магическая функция: 
    1. Читает Excel.
    2. Ищет лист с Фин. результатами.
    3. Ищет год внутри ячеек (За 20xx г.).
    4. Собирает данные (Код -> Значение).
    """
    try:
        dfs = pd.read_excel(file_obj, sheet_name=None, header=None)
    except:
        return None, None, "Ошибка чтения Excel"

    # 1. Поиск нужного листа (Отчет о фин результатах)
    target_df = None
    for name, df in dfs.items():
        if 'фин' in name.lower() and 'результ' in name.lower():
            target_df = df
            break
    
    # Если не нашли по имени, берем 3-й лист (обычно это Форма 2) или 1-й
    if target_df is None:
        if len(dfs) >= 3: target_df = list(dfs.values())[2]
        elif len(dfs) > 0: target_df = list(dfs.values())[0]
        else: return None, None, "Пустой файл"

    # 2. Поиск года в тексте (сканируем первые 20 строк)
    detected_year = None
    
    # Регулярка ищет "За 2024 г." или "2024 г." или "На 31 декабря 2024"
    year_pattern = re.compile(r'\b(20\d{2})\b')
    
    # Сначала ищем в заголовках (обычно строки 0-10)
    for r in range(min(20, len(target_df))):
        row_values = target_df.iloc[r].astype(str).tolist()
        row_str = " ".join(row_values)
        
        # Ищем фразы типа "За 2024" (приоритет)
        if "за" in row_str.lower() and "г." in row_str.lower():
            match = year_pattern.search(row_str)
            if match:
                detected_year = int(match.group(1))
                break
        # Если нет "За", ищем просто дату в контексте заголовка
        elif "на" in row_str.lower() and "декабря" in row_str.lower():
            match = year_pattern.search(row_str)
            if match:
                detected_year = int(match.group(1))
                break

    if not detected_year:
        return None, None, "Не удалось найти год внутри файла (искал 'За 20xx г.')"

    # 3. Сбор данных (ищем коды строк 2110, 2400 и т.д.)
    # Обычно структура: [Пояснения, Название, КОД, Значение_Тек, Значение_Пред]
    # Нам нужно найти колонку с кодами и колонку с текущим значением.
    
    data_map = {}
    
    # Ищем индекс колонки с кодами (обычно там 4-значные числа)
    code_col_idx = -1
    value_col_idx = -1
    
    # Пробежимся, чтобы найти колонку, где много кодов (2110, 2120...)
    for c in range(len(target_df.columns)):
        col_data = pd.to_numeric(target_df.iloc[:, c], errors='coerce')
        # Если в колонке есть 2110 и 2400 - это она
        if col_data.isin([2110, 2400, 2120]).sum() >= 2:
            code_col_idx = c
            # Обычно значение текущего года идет СЛЕДУЮЩЕЙ колонкой (c+1)
            # Но иногда бывают пустые колонки. Ищем первую непустую числовую справа.
            value_col_idx = c + 1
            break
            
    if code_col_idx == -1:
        return None, None, "Не найдена колонка с кодами строк (2110, 2120...)"

    # Извлекаем данные
    for index, row in target_df.iterrows():
        try:
            code_val = pd.to_numeric(row[code_col_idx], errors='coerce')
            if pd.notna(code_val) and code_val > 1000:
                # Берем значение
                raw_val = row[value_col_idx]
                clean_val = clean_number(raw_val)
                data_map[int(code_val)] = clean_val
        except:
            continue

    return detected_year, data_map, None

def render_task(task_num, topic, goal, task_text):
    st.markdown(f"""
    <div style="background-color: #ffffff; padding: 20px; border: 1px solid #e0e0e0; border-radius: 5px; margin-top: 30px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <h4 style="color: #2c3e50; margin-top: 0;">📝 Задание {task_num}</h4>
        <p><b>Тема:</b> {topic}</p>
        <p><b>Цель:</b> {goal}</p>
        <hr style="border-top: 1px solid #eee;">
        <p><b>Что сделать:</b> {task_text}</p>
    </div>
    """, unsafe_allow_html=True)

def render_table_header(table_num, analysis_full_name, subject_genitive, period):
    header_text = f"<b>Таблица {table_num}.</b> {analysis_full_name} {subject_genitive}, xxx, xxx, {period}"
    st.markdown(f"""
    <div style="background-color: #f8f9fa; padding: 10px 15px; border-radius: 4px; margin-bottom: 5px; border-left: 5px solid #6c757d; color: #333; font-size: 15px;">
        {header_text}
    </div>
    """, unsafe_allow_html=True)

def get_ai_analysis(table_df, task_context, api_key):
    try:
        client = OpenAI(api_key=api_key, base_url="https://openai.api.proxyapi.ru/v1")
        prompt = f"Ты студент. Контекст: {task_context}. Данные:\n{table_df.to_string()}\nНапиши вывод (3-4 предложения) в академическом стиле на русском."
        response = client.chat.completions.create(
            model="anthropic/claude-sonnet-4-20250514",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e: return f"❌ Ошибка API: {e}"

# ==========================================
# 3. БОКОВАЯ ПАНЕЛЬ
# ==========================================
with st.sidebar:
    st.header("⚙️ Настройки")
    api_key = st.text_input("API Key (ProxyAPI)", type="password")
    use_ai = st.checkbox("✍️ Добавлять выводы ИИ", value=True)
    
    st.info("📂 Загрузите файлы (любые имена):")
    st.caption("Программа сама найдет год внутри файла.")
    uploaded_files = st.file_uploader("Загрузчик файлов", type=["xlsx"], accept_multiple_files=True)

# ==========================================
# 4. ОСНОВНАЯ ЛОГИКА
# ==========================================

# Словарь для хранения данных всех лет: {2022: {Code: Val}, 2023: {...}}
GLOBAL_DATA = {}

if uploaded_files:
    # 1. ПАРСИНГ ФАЙЛОВ
    logs = []
    for file in uploaded_files:
        year, data, error = get_year_and_data_from_excel(file)
        
        if year and data:
            GLOBAL_DATA[year] = data
            logs.append(f"✅ {file.name} -> Обнаружен {year} год")
        else:
            logs.append(f"❌ {file.name} -> {error}")
    
    # Отображение статуса загрузки
    with st.expander("Статус загрузки файлов", expanded=False):
        for log in logs:
            st.write(log)

    # 2. ПРОВЕРКА ДАННЫХ
    years = sorted(GLOBAL_DATA.keys())
    
    if len(years) < 2:
        st.warning("⚠️ Для анализа нужно минимум 2 года. Загрузите больше файлов.")
    else:
        current_year = years[-1]
        prev_year = years[-2]
        
        # Превращаем в DataFrame для удобства
        # Строки - показатели, Колонки - годы
        codes_map = {
            'Выручка': 2110, 'Себестоимость продаж': 2120, 'Валовая прибыль': 2100,
            'Коммерческие расходы': 2210, 'Управленческие расходы': 2220, 'Прибыль от продаж': 2200,
            'Прочие доходы': 2340, 'Прочие расходы': 2350, 'Налог на прибыль': 2410, 
            'Чистая прибыль': 2400
        }
        
        # Собираем сводную таблицу
        df_res = pd.DataFrame(index=codes_map.keys(), columns=years)
        for name, code in codes_map.items():
            for y in years:
                # Берем значение из глобального хранилища, если нет - 0
                df_res.loc[name, y] = GLOBAL_DATA[y].get(code, 0)
        
        # Принудительно конвертируем в float
        df_res = df_res.apply(pd.to_numeric)

        st.success(f"✅ Анализ выполняется за период: {years[0]} - {years[-1]} гг.")

        # ==========================================
        # ЗАДАНИЕ 1: ВЕРТИКАЛЬНЫЙ АНАЛИЗ (Все годы)
        # ==========================================
        years_display = years[-3:] if len(years) >= 3 else years # Показываем последние 3 года, чтоб не растягивать
        
        render_task("1", "Анализ структуры", "Изучение структуры доходов и расходов.", f"Анализ за {years_display[0]}-{years_display[-1]} гг.")
        
        df_v = df_res[years_display].copy()
        
        # Расчет удельных весов
        for y in years_display:
            base_rev = df_v.loc['Выручка', y]
            if base_rev != 0:
                df_v[f'Уд. вес {y} (%)'] = (df_v[y] / base_rev * 100)
            else:
                df_v[f'Уд. вес {y} (%)'] = 0
        
        # Сортировка: Год, %, Год, %
        cols_v = []
        for y in years_display:
            cols_v.append(y)
            cols_v.append(f'Уд. вес {y} (%)')
        
        render_table_header("1", "Вертикальный анализ", "финансовых результатов", f"{years_display[0]}-{years_display[-1]} гг.")
        st.dataframe(df_v[cols_v].style.format("{:,.2f}"))
        if api_key and use_ai: st.info(get_ai_analysis(df_v, "Структура показателей", api_key))

        # ==========================================
        # ЗАДАНИЕ 2: ГОРИЗОНТАЛЬНЫЙ АНАЛИЗ
        # ==========================================
        render_task("2", "Анализ динамики", "Оценка темпов роста.", f"Сравнение относительно {current_year} года.")
        
        df_h = df_res[years_display].copy()
        cols_h = years_display.copy()
        
        # Считаем отклонения для всех лет кроме последнего, сравнивая с последним
        # Например: 2024-2023, 2024-2022
        # (или как в вашем задании: Отклонение и Темп)
        
        # Сравниваем Предпоследний с Последним (2024/2023)
        if prev_year in df_h.columns:
            df_h[f'Откл. {current_year}-{prev_year}'] = df_h[current_year] - df_h[prev_year]
            df_h[f'Темп {current_year}/{prev_year} (%)'] = (df_h[current_year] / df_h[prev_year] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            cols_h.extend([f'Откл. {current_year}-{prev_year}', f'Темп {current_year}/{prev_year} (%)'])
            
        # Сравниваем Пред-предпоследний с Последним (2024/2022)
        if len(years_display) > 2:
            prev_prev = years_display[-3]
            df_h[f'Откл. {current_year}-{prev_prev}'] = df_h[current_year] - df_h[prev_prev]
            df_h[f'Темп {current_year}/{prev_prev} (%)'] = (df_h[current_year] / df_h[prev_prev] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            cols_h.extend([f'Откл. {current_year}-{prev_prev}', f'Темп {current_year}/{prev_prev} (%)'])

        render_table_header("2", "Горизонтальный анализ", "финансовых результатов", f"{years_display[0]}-{years_display[-1]} гг.")
        st.dataframe(df_h[cols_h].style.format("{:,.2f}"))
        if api_key and use_ai: st.info(get_ai_analysis(df_h, "Динамика прибыли", api_key))

        # ==========================================
        # ЗАДАНИЕ 3: ТРЕНДОВЫЙ АНАЛИЗ (ВСЕ ГОДЫ)
        # ==========================================
        render_task("3", "Трендовый анализ", "Анализ тенденций за весь период.", f"Динамика Чистой прибыли ({years[0]}-{years[-1]}).")
        
        trend_rows = []
        base_first_year = df_res.loc['Чистая прибыль', years[0]]
        prev_val = None
        
        for y in years:
            val = df_res.loc['Чистая прибыль', y]
            
            # Цепной (к прошлому)
            if prev_val is not None and prev_val != 0:
                chain_gr = (val / prev_val * 100)
            else:
                chain_gr = 100.0 if prev_val is None else 0
                
            # Базисный (к первому)
            if base_first_year != 0:
                base_gr = (val / base_first_year * 100)
            else:
                base_gr = 0
                
            trend_rows.append({
                'Год': y,
                'Чистая прибыль': val,
                'Темп роста (цепной) %': chain_gr,
                'Темп роста (базисный) %': base_gr
            })
            prev_val = val
            
        df_trend = pd.DataFrame(trend_rows).set_index('Год')
        render_table_header("3", "Трендовый анализ", "чистой прибыли", f"{years[0]}-{years[-1]} гг.")
        st.table(df_trend.style.format("{:,.2f}"))

        # ==========================================
        # ЗАДАНИЕ 4: ФАКТОРНЫЙ АНАЛИЗ (Последние 2 года)
        # ==========================================
        render_task("4", "Факторный анализ", "Влияние факторов на прибыль.", f"{current_year} к {prev_year} г.")
        
        # Получаем абсолютные значения для правильной математики
        v0 = df_res[prev_year].abs()
        v1 = df_res[current_year].abs()
        
        # Расчет влияний
        infl_vr = v1['Выручка'] - v0['Выручка']
        infl_sb = -(v1['Себестоимость продаж'] - v0['Себестоимость продаж'])
        infl_ur = -(v1['Управленческие расходы'] - v0['Управленческие расходы'])
        infl_kr = -(v1['Коммерческие расходы'] - v0['Коммерческие расходы'])
        infl_pd = v1['Прочие доходы'] - v0['Прочие доходы']
        infl_pr = -(v1['Прочие расходы'] - v0['Прочие расходы'])
        infl_np = -(v1['Налог на прибыль'] - v0['Налог на прибыль'])
        
        f_data = [
            ['Выручка', v0['Выручка'], v1['Выручка'], infl_vr],
            ['Себестоимость', v0['Себестоимость продаж'], v1['Себестоимость продаж'], infl_sb],
            ['Упр. расходы', v0['Управленческие расходы'], v1['Управленческие расходы'], infl_ur],
            ['Комм. расходы', v0['Коммерческие расходы'], v1['Коммерческие расходы'], infl_kr],
            ['Прочие доходы', v0['Прочие доходы'], v1['Прочие доходы'], infl_pd],
            ['Прочие расходы', v0['Прочие расходы'], v1['Прочие расходы'], infl_pr],
            ['Налог на прибыль', v0['Налог на прибыль'], v1['Налог на прибыль'], infl_np]
        ]
        
        total_inf = sum([x[3] for x in f_data])
        # Добавляем итог
        f_data.append(['ИТОГО влияние', 0, 0, total_inf])
        # Проверка по факту
        fact_change = df_res.loc['Чистая прибыль', current_year] - df_res.loc['Чистая прибыль', prev_year]
        f_data.append(['Изм. ЧП (факт)', df_res.loc['Чистая прибыль', prev_year], df_res.loc['Чистая прибыль', current_year], fact_change])
        
        df_fact = pd.DataFrame(f_data, columns=['Фактор', f'Базис ({prev_year})', f'Факт ({current_year})', 'Влияние'])
        
        render_table_header("4", "Факторный анализ", "чистой прибыли", f"{current_year} к {prev_year} г.")
        st.table(df_fact.style.format({col: "{:,.2f}" for col in df_fact.columns if col != 'Фактор'}))
        if api_key and use_ai: st.info(get_ai_analysis(df_fact, "Факторный анализ", api_key))

        # ==========================================
        # ЗАДАНИЕ 5: АНАЛИЗ ЗАТРАТ
        # ==========================================
        render_task("5", "Анализ затрат", "Динамика расходов.", f"Анализ за {current_year} и {prev_year} гг.")
        
        cost_items = ['Себестоимость продаж', 'Коммерческие расходы', 'Управленческие расходы']
        df_costs = df_res.loc[cost_items, [prev_year, current_year]].abs().copy()
        df_costs.loc['ИТОГО'] = df_costs.sum()
        
        df_costs['Абс. откл.'] = df_costs[current_year] - df_costs[prev_year]
        df_costs['Темп роста %'] = (df_costs[current_year] / df_costs[prev_year] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        
        render_table_header("5", "Комплексный анализ", "затрат", f"{current_year} к {prev_year} г.")
        st.dataframe(df_costs.style.format("{:,.2f}"))

        # ==========================================
        # ЗАДАНИЕ 6: CVP
        # ==========================================
        render_task("6", "CVP-анализ", "Точка безубыточности.", "Калькулятор (ввод вручную).")
        
        cvp_cols = st.columns(2)
        p = cvp_cols[0].number_input("Цена (P)", 1000.0)
        avc = cvp_cols[0].number_input("VC на ед.", 600.0)
        tfc = cvp_cols[1].number_input("TFC (Пост. затраты)", 200000.0)
        
        if (p - avc) > 0:
            bep = tfc / (p - avc)
            st.success(f"Точка безубыточности: {bep:,.0f} шт. / {bep*p:,.2f} руб.")
        else:
            st.error("Цена ниже переменных затрат!")

        # СКАЧИВАНИЕ
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_v.to_excel(writer, sheet_name='Вертикальный')
            df_h.to_excel(writer, sheet_name='Горизонтальный')
            df_trend.to_excel(writer, sheet_name='Трендовый')
            df_fact.to_excel(writer, sheet_name='Факторный', index=False)
        st.download_button("📥 Скачать Excel", data=output.getvalue(), file_name="analysis.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("👈 Загрузите файлы в меню слева. Имена файлов не важны!")
