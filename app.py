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
    page_title="Практическая работа № 1",
    page_icon="📊",
    layout="wide"
)

# ИСПРАВЛЕННЫЙ ЗАГОЛОВОК
st.title("Практическая работа № 1: Анализ финансовых результатов")

# ==========================================
# 2. ФУНКЦИИ (ОФОРМЛЕНИЕ И ЛОГИКА)
# ==========================================

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

def render_table_header(table_num, analysis_full_name, subject_genitive, period=""):
    """Формирует стандартный заголовок таблицы: Таблица X. Название..."""
    header_text = f"<b>Таблица {table_num}.</b> {analysis_full_name} {subject_genitive} {period}"
    st.markdown(f"""
    <div style="background-color: #f8f9fa; padding: 10px 15px; border-radius: 4px; margin-bottom: 5px; border-left: 5px solid #6c757d; color: #333; font-size: 16px;">
        {header_text}
    </div>
    """, unsafe_allow_html=True)

def get_ai_analysis(table_df, task_context, api_key):
    if not api_key: return "⚠️ Введите API Key для получения выводов."
    try:
        client = OpenAI(api_key=api_key, base_url="https://openai.api.proxyapi.ru/v1")
        prompt = f"Ты финансовый аналитик. Контекст: {task_context}. Данные:\n{table_df.to_string()}\nНапиши краткий аналитический вывод (3-4 предложения) с цифрами."
        response = client.chat.completions.create(
            model="gpt-4o-mini", # Можно менять модель
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e: return f"❌ Ошибка API: {e}"

def load_excel_sheet(file):
    """Ищет лист с финансовыми результатами."""
    try:
        dfs = pd.read_excel(file, sheet_name=None, header=None)
        for name, df in dfs.items():
            if 'фин' in name.lower() or 'результ' in name.lower() or 'форма 2' in name.lower(): return df
        # Если явного названия нет, пробуем эвристику: ищем лист с кодом 2110
        for name, df in dfs.items():
            s = df.astype(str).to_string()
            if '2110' in s: return df
        return list(dfs.values())[0]
    except: return None

def detect_year_in_df(df):
    """
    Пытается найти год отчета в первых строках файла (ищет 2020-2030).
    Возвращает самый большой найденный год (считаем его отчетным).
    """
    if df is None: return None
    
    # Преобразуем первые 20 строк в текст
    header_part = df.head(20).astype(str).to_string()
    # Ищем года (2020-2029)
    years = re.findall(r'202[0-9]', header_part)
    
    if years:
        years = [int(y) for y in years]
        return max(years) # Возвращаем самый свежий год
    return None

def get_values_by_code(df, code):
    """Возвращает (Value_Current_Year, Value_Previous_Year)"""
    if df is None: return (0, 0)
    values_found = []
    for index, row in df.iterrows():
        for i, cell in enumerate(row):
            try:
                # Ищем код строки (например, 2110)
                if pd.to_numeric(cell, errors='coerce') == code:
                    # Как только нашли код, ищем справа от него два числа
                    for next_cell in row[i+1:]:
                        if pd.notna(next_cell) and str(next_cell).strip() not in ['', '-', '(-)']:
                            val_str = str(next_cell).replace(' ', '').replace('\xa0', '')
                            if val_str.startswith('(') and val_str.endswith(')'):
                                val_str = '-' + val_str[1:-1]
                            val = pd.to_numeric(val_str, errors='coerce')
                            
                            if pd.notna(val): values_found.append(val)
                            if len(values_found) == 2: return tuple(values_found)
            except: continue
    
    if len(values_found) == 1: return (values_found[0], 0)
    return (0, 0)

# ==========================================
# 3. БОКОВАЯ ПАНЕЛЬ
# ==========================================
with st.sidebar:
    st.header("⚙️ Настройки")
    api_key = st.text_input("API Key (ProxyAPI)", type="password")
    use_ai = st.checkbox("✍️ Добавлять выводы ИИ", value=True)
    
    st.info("📂 Загрузка файлов")
    # МУЛЬТИ-ЗАГРУЗКА
    uploaded_files = st.file_uploader(
        "Загрузите все отчеты (xlsx)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )

# ==========================================
# 4. ОБРАБОТКА ДАННЫХ
# ==========================================

codes_map = {
    'Выручка': 2110, 'Себестоимость продаж': 2120, 'Валовая прибыль': 2100,
    'Коммерческие расходы': 2210, 'Управленческие расходы': 2220, 'Прибыль от продаж': 2200,
    'Прочие доходы': 2340, 'Прочие расходы': 2350, 'Налог на прибыль': 2410, 
    'Чистая прибыль': 2400
}

df_res = pd.DataFrame()

if uploaded_files:
    # Словарь для сбора данных: {Год: {Показатель: Значение}}
    master_data = {} 

    for file in uploaded_files:
        df_raw = load_excel_sheet(file)
        if df_raw is not None:
            # 1. Пытаемся найти год в файле
            detected_year = detect_year_in_df(df_raw)
            
            # Если год не найден внутри, пробуем вытащить из названия файла
            if not detected_year:
                fname_years = re.findall(r'202[0-9]', file.name)
                if fname_years:
                    detected_year = int(max(fname_years))
                else:
                    st.warning(f"⚠️ Не удалось определить год для файла: {file.name}. Пропускаем.")
                    continue
            
            year_curr = detected_year
            year_prev = detected_year - 1
            
            # 2. Извлекаем данные по кодам
            for metric, code in codes_map.items():
                val_curr, val_prev = get_values_by_code(df_raw, code)
                
                # Сохраняем текущий год
                if year_curr not in master_data: master_data[year_curr] = {}
                master_data[year_curr][metric] = val_curr
                
                # Сохраняем предыдущий год (только если его еще нет или мы перезаписываем более старые данные)
                if year_prev not in master_data: master_data[year_prev] = {}
                
                # Если значения для prev года еще нет, записываем
                if metric not in master_data[year_prev]:
                    master_data[year_prev][metric] = val_prev

    # 3. Превращаем в DataFrame
    if master_data:
        df_res = pd.DataFrame(master_data).sort_index(axis=1) # Сортируем колонки-года по возрастанию
        # Упорядочиваем строки по логическому порядку (как в codes_map)
        df_res = df_res.reindex(codes_map.keys())
        
        years_avail = sorted([str(y) for y in df_res.columns])
        st.success(f"✅ Данные успешно загружены за периоды: {', '.join(years_avail)}")
        
        # Определяем "базовые" года для анализа (два последних)
        if len(years_avail) >= 2:
            last_year = years_avail[-1]
            prev_year = years_avail[-2]
        else:
            last_year = years_avail[0]
            prev_year = years_avail[0] # Fallback

        # ==========================================
        # ЗАДАНИЕ 1: ВЕРТИКАЛЬНЫЙ АНАЛИЗ (ДИНАМИЧЕСКИЙ)
        # ==========================================
        render_task("1", "Вертикальный анализ", "Структура доходов и расходов", 
                   f"Анализ за весь доступный период ({years_avail[0]} - {years_avail[-1]}).")

        df_v = df_res.copy()
        display_cols = []
        
        for y in df_res.columns:
            y_str = str(y)
            base_val = df_v.loc['Выручка', y]
            col_share = f'{y} (%)'
            df_v[col_share] = (df_v[y] / base_val * 100).fillna(0)
            display_cols.extend([y, col_share]) # Чередуем: Сумма, Процент

        # Переставляем колонки для красоты
        df_v_display = df_v[display_cols]
        
        render_table_header("1", "Динамика структуры", "финансовых результатов")
        st.dataframe(df_v_display.style.format("{:,.2f}"))
        
        if api_key and use_ai:
            st.info(get_ai_analysis(df_v_display, f"Вертикальный анализ за {years_avail}", api_key))

        # ==========================================
        # ЗАДАНИЕ 2: ГОРИЗОНТАЛЬНЫЙ АНАЛИЗ (ПОСЛЕДНИЕ 2 ГОДА)
        # ==========================================
        if len(years_avail) >= 2:
            render_task("2", "Горизонтальный анализ", "Динамика показателей", 
                       f"Сравнение {last_year} года относительно {prev_year}.")

            df_h = df_res[[int(prev_year), int(last_year)]].copy()
            y1, y2 = int(prev_year), int(last_year)
            
            df_h['Абс. откл.'] = df_h[y2] - df_h[y1]
            df_h['Темп роста (%)'] = (df_h[y2] / df_h[y1] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            
            render_table_header("2", "Горизонтальный анализ", f"{last_year}/{prev_year}")
            st.dataframe(df_h.style.format("{:,.2f}"))
            
            if api_key and use_ai:
                st.info(get_ai_analysis(df_h, f"Сравнение {last_year} к {prev_year}", api_key))
        else:
            st.warning("Для горизонтального анализа нужно минимум 2 года данных.")

        # ==========================================
        # ЗАДАНИЕ 3: ТРЕНДОВЫЙ АНАЛИЗ (ВСЕ ГОДА)
        # ==========================================
        render_task("3", "Трендовый анализ", "Тенденция чистой прибыли", "Цепные и базисные темпы роста.")

        trend_list = []
        base_year_val = df_res.loc['Чистая прибыль', df_res.columns[0]]
        prev_val_trend = None
        
        for y in df_res.columns:
            curr = df_res.loc['Чистая прибыль', y]
            
            abs_ch = (curr - prev_val_trend) if prev_val_trend is not None else 0
            rate_ch = (curr / prev_val_trend * 100) if (prev_val_trend and prev_val_trend != 0) else 100.0
            rate_bs = (curr / base_year_val * 100) if base_year_val != 0 else 0
            
            trend_list.append({
                'Год': str(y),
                'Чистая прибыль': curr,
                'Цепной темп %': rate_ch if y != df_res.columns[0] else 100,
                'Базисный темп %': rate_bs
            })
            prev_val_trend = curr
            
        df_trend = pd.DataFrame(trend_list).set_index('Год')
        render_table_header("3", "Трендовый анализ", "чистой прибыли")
        st.table(df_trend.style.format("{:,.2f}"))
        
        # График
        st.line_chart(df_trend['Чистая прибыль'])

        # ==========================================
        # ЗАДАНИЕ 4: ФАКТОРНЫЙ АНАЛИЗ (ПОСЛЕДНИЕ 2 ГОДА)
        # ==========================================
        if len(years_avail) >= 2:
            render_task("4", "Факторный анализ", "Влияние на прибыль", 
                       f"Модель: ЧП = В - С - КР - УР + ПД - ПР - НП. ({last_year} к {prev_year})")
            
            def get_abs(row, yr): return abs(df_res.loc[row, int(yr)])

            v0 = {k: get_abs(k, prev_year) for k in df_res.index}
            v1 = {k: get_abs(k, last_year) for k in df_res.index}
            
            factors = [
                ('Выручка', v1['Выручка'] - v0['Выручка']),
                ('Себестоимость', -(v1['Себестоимость продаж'] - v0['Себестоимость продаж'])),
                ('Упр. расходы', -(v1['Управленческие расходы'] - v0['Управленческие расходы'])),
                ('Комм. расходы', -(v1['Коммерческие расходы'] - v0['Коммерческие расходы'])),
                ('Прочие доходы', v1['Прочие доходы'] - v0['Прочие доходы']),
                ('Прочие расходы', -(v1['Прочие расходы'] - v0['Прочие расходы'])),
                ('Налог на прибыль', -(v1['Налог на прибыль'] - v0['Налог на прибыль']))
            ]
            
            total_inf = sum([f[1] for f in factors])
            
            df_fact = pd.DataFrame(factors, columns=['Фактор', 'Влияние'])
            df_fact.loc[len(df_fact)] = ['ИТОГО', total_inf]
            
            render_table_header("4", "Факторный анализ", f"{last_year} к {prev_year}")
            st.table(df_fact.style.format({"Влияние": "{:,.2f}"}))
            
            if api_key and use_ai:
                st.info(get_ai_analysis(df_fact, "Факторы изменения прибыли", api_key))

        # ==========================================
        # ЗАДАНИЕ 5: АНАЛИЗ ЗАТРАТ
        # ==========================================
        if len(years_avail) >= 2:
            render_task("5", "Анализ затрат", "Структура расходов", f"За {prev_year} и {last_year} гг.")
            
            cost_rows = ['Себестоимость продаж', 'Коммерческие расходы', 'Управленческие расходы']
            df_c = df_res.loc[cost_rows, [int(prev_year), int(last_year)]].apply(abs)
            df_c.loc['ИТОГО'] = df_c.sum()
            
            y1, y2 = int(prev_year), int(last_year)
            df_c['Темп роста %'] = (df_c[y2] / df_c[y1] * 100).fillna(0)
            df_c[f'Доля {y1} %'] = (df_c[y1] / df_c.loc['ИТОГО', y1] * 100)
            df_c[f'Доля {y2} %'] = (df_c[y2] / df_c.loc['ИТОГО', y2] * 100)
            
            render_table_header("5", "Анализ затрат", "")
            st.dataframe(df_c.style.format("{:,.2f}"))

        # ==========================================
        # ЗАДАНИЕ 6: CVP (Берем последний год)
        # ==========================================
        render_task("6", "CVP-анализ", "Безубыточность", "Расчет на основе введенных данных (моделирование).")
        
        cvp_type = st.radio("Тип:", ["Однопродуктовое", "Многопродуктовое"], horizontal=True)
        
        if cvp_type == "Однопродуктовое":
            c1, c2 = st.columns(2)
            p = c1.number_input("Цена (P)", 1000.0)
            avc = c1.number_input("Перем. затраты (AVC)", 600.0)
            # Пытаемся взять управленческие расходы последнего года как базу для TFC
            default_tfc = df_res.loc['Управленческие расходы', int(last_year)] if len(years_avail)>0 else 200000.0
            tfc = c2.number_input("Пост. затраты (TFC)", abs(float(default_tfc)))
            
            md = p - avc
            if md > 0:
                bep = tfc / md
                st.metric("Точка безубыточности (шт)", f"{bep:,.0f}")
                st.metric("Точка безубыточности (руб)", f"{bep*p:,.2f}")
            else:
                st.error("Маржа отрицательная (Цена < AVC)")

        # ==========================================
        # СКАЧИВАНИЕ
        # ==========================================
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_res.to_excel(writer, sheet_name='Сводные_Данные')
            if 'df_v_display' in locals(): df_v_display.to_excel(writer, sheet_name='Вертикальный')
            if 'df_h' in locals(): df_h.to_excel(writer, sheet_name='Горизонтальный')
            if 'df_trend' in locals(): df_trend.to_excel(writer, sheet_name='Трендовый')
        
        st.download_button(
            "📥 Скачать сводный отчет", 
            data=output.getvalue(), 
            file_name="multi_year_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
else:
    st.info("👈 Пожалуйста, загрузите файлы отчетов (можно выбрать сразу несколько).")
