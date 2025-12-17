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
    page_title="Практическая работа №1",
    page_icon="📊",
    layout="wide"
)

# ЕДИНСТВЕННЫЙ ЗАГОЛОВОК
st.title("📊 Практическая работа №1: Анализ финансовых результатов")

# ==========================================
# 2. ФУНКЦИИ
# ==========================================

def render_task(task_num, topic, goal, task_text):
    """Выводит блок 'Задание'."""
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
    """Заголовок таблицы в одну строку."""
    # Пример: Таблица 1. Вертикальный анализ финансовых результатов, xxx, xxx, 2022-2024 гг.
    header_text = f"<b>Таблица {table_num}.</b> {analysis_full_name} {subject_genitive}, xxx, xxx, {period}"
    st.markdown(f"""
    <div style="background-color: #f8f9fa; padding: 10px 15px; border-radius: 4px; margin-bottom: 5px; border-left: 5px solid #6c757d; color: #333; font-size: 15px;">
        {header_text}
    </div>
    """, unsafe_allow_html=True)

def get_ai_analysis(table_df, task_context, api_key):
    if not api_key: return "⚠️ Введите API Key для получения выводов."
    try:
        client = OpenAI(api_key=api_key, base_url="https://openai.api.proxyapi.ru/v1")
        prompt = f"Ты студент. Контекст задания: {task_context}. Данные таблицы:\n{table_df.to_string()}\nНапиши аналитический вывод (3-4 предложения) в академическом стиле на русском языке."
        response = client.chat.completions.create(
            model="anthropic/claude-sonnet-4-20250514", # или gpt-4o-mini
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e: return f"❌ Ошибка API: {e}"

def load_excel_sheet(file):
    """Ищет подходящий лист."""
    try:
        dfs = pd.read_excel(file, sheet_name=None, header=None)
        for name, df in dfs.items():
            if 'фин' in name.lower() or 'результ' in name.lower(): return df
        if len(dfs) >= 3: return list(dfs.values())[2]
        return list(dfs.values())[0]
    except: return None

def detect_year_in_df(df):
    """Ищет год в содержимом (2020-2029)."""
    if df is None: return None
    header_text = df.head(20).astype(str).to_string()
    years = re.findall(r'202[0-9]', header_text)
    if years:
        return max([int(y) for y in years])
    return None

def get_values_by_code(df, code):
    """Возвращает (Текущий, Предыдущий) по коду строки."""
    if df is None: return (0, 0)
    values_found = []
    for index, row in df.iterrows():
        for i, cell in enumerate(row):
            try:
                if pd.to_numeric(cell, errors='coerce') == code:
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
    uploaded_files = st.file_uploader("Загрузите все отчеты (xlsx)", type=["xlsx"], accept_multiple_files=True)

# ==========================================
# 4. ОБРАБОТКА И АНАЛИЗ
# ==========================================

codes_map = {
    'Выручка': 2110, 'Себестоимость продаж': 2120, 'Валовая прибыль': 2100,
    'Коммерческие расходы': 2210, 'Управленческие расходы': 2220, 'Прибыль от продаж': 2200,
    'Прочие доходы': 2340, 'Прочие расходы': 2350, 'Налог на прибыль': 2410, 
    'Чистая прибыль': 2400
}

if uploaded_files:
    master_data = {}
    
    # 1. Читаем файлы
    for file in uploaded_files:
        df_raw = load_excel_sheet(file)
        if df_raw is not None:
            # Определяем год
            year = detect_year_in_df(df_raw)
            if not year:
                # Если не нашли внутри, ищем в имени файла
                fname_years = re.findall(r'202[0-9]', file.name)
                if fname_years: year = int(max(fname_years))
            
            if year:
                # Парсим данные
                for metric, code in codes_map.items():
                    v_curr, v_prev = get_values_by_code(df_raw, code)
                    
                    if year not in master_data: master_data[year] = {}
                    master_data[year][metric] = v_curr
                    
                    if (year-1) not in master_data: master_data[year-1] = {}
                    if metric not in master_data[year-1]: # Не перезаписываем, если уже есть более точные данные
                        master_data[year-1][metric] = v_prev

    # 2. Создаем DataFrame
    if master_data:
        df_res = pd.DataFrame(master_data).sort_index(axis=1) # Годы по возрастанию
        df_res = df_res.reindex(codes_map.keys())
        years = sorted(df_res.columns)
        
        st.success(f"✅ Данные загружены за период: {years[0]} - {years[-1]} гг.")
        
        # Определяем "базовые" годы (последний, предпоследний)
        curr_y = years[-1]
        prev_y = years[-2] if len(years) > 1 else years[0]
        
        # ---------------------------------------------------------
        # ЗАДАНИЕ 1: ВЕРТИКАЛЬНЫЙ АНАЛИЗ
        # ---------------------------------------------------------
        render_task("1", "Анализ структуры финансовых результатов", "Изучение структуры доходов и расходов.",
                   f"Провести вертикальный анализ за {years[0]}-{years[-1]} гг.")
        
        # Берем последние 3 года для отображения (или сколько есть)
        disp_years = years[-3:] if len(years) >= 3 else years
        
        df_v = df_res[disp_years].copy()
        cols_v = []
        
        for y in disp_years:
            base = df_v.loc['Выручка', y]
            df_v[f'Уд. вес {y} (%)'] = (df_v[y] / base * 100).fillna(0)
            cols_v.extend([y, f'Уд. вес {y} (%)'])
            
        render_table_header("1", "Вертикальный сравнительный анализ", "финансовых результатов", f"{disp_years[0]}-{disp_years[-1]} гг.")
        st.dataframe(df_v[cols_v].style.format("{:,.2f}"))
        
        if api_key and use_ai: st.info(get_ai_analysis(df_v[cols_v], "Структура доходов и расходов", api_key))

        # ---------------------------------------------------------
        # ЗАДАНИЕ 2: ГОРИЗОНТАЛЬНЫЙ АНАЛИЗ
        # ---------------------------------------------------------
        render_task("2", "Анализ динамики прибыли", "Оценка темпов роста.", 
                   f"Провести горизонтальный анализ. База сравнения: {curr_y} год.")
        
        if len(years) >= 2:
            df_h = df_res[disp_years].copy()
            cols_h = disp_years.copy()
            
            # Логика как в вашем коде: Сравниваем Текущий (2024) с Пред (2023) и Пред-Пред (2022)
            # 1. Сравнение с "Пред-Пред" (если есть, например 2022)
            if len(disp_years) > 2:
                y_base_old = disp_years[-3] # 2022
                df_h[f'Откл. {curr_y}-{y_base_old}'] = df_h[curr_y] - df_h[y_base_old]
                df_h[f'Темп {curr_y}/{y_base_old} (%)'] = (df_h[curr_y] / df_h[y_base_old] * 100).replace([np.inf, -np.inf], 0).fillna(0)
                cols_h.extend([f'Откл. {curr_y}-{y_base_old}', f'Темп {curr_y}/{y_base_old} (%)'])
            
            # 2. Сравнение с "Пред" (2023)
            y_prev = disp_years[-2] # 2023
            df_h[f'Откл. {curr_y}-{y_prev}'] = df_h[curr_y] - df_h[y_prev]
            df_h[f'Темп {curr_y}/{y_prev} (%)'] = (df_h[curr_y] / df_h[y_prev] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            cols_h.extend([f'Откл. {curr_y}-{y_prev}', f'Темп {curr_y}/{y_prev} (%)'])
            
            render_table_header("2", "Горизонтальный сравнительный анализ", "финансовых результатов", f"{disp_years[0]}-{curr_y} гг.")
            st.dataframe(df_h[cols_h].style.format("{:,.2f}"))
            
            if api_key and use_ai: st.info(get_ai_analysis(df_h[cols_h], "Динамика прибыли", api_key))
        else:
            st.warning("Недостаточно данных для горизонтального анализа (нужно минимум 2 года).")

        # ---------------------------------------------------------
        # ЗАДАНИЕ 3: ТРЕНДОВЫЙ АНАЛИЗ
        # ---------------------------------------------------------
        render_task("3", "Трендовый анализ показателей", "Выявление тенденций.", "Трендовый анализ Чистой прибыли.")
        
        trend_rows = []
        base_val_start = df_res.loc['Чистая прибыль', years[0]]
        prev_val = None
        
        for y in years:
            val = df_res.loc['Чистая прибыль', y]
            chain = (val/prev_val*100) if (prev_val and prev_val!=0) else (100 if prev_val is None else 0)
            base = (val/base_val_start*100) if base_val_start!=0 else 0
            
            trend_rows.append({
                'Год': str(y), 'Чистая прибыль': val,
                'Цепной темп %': chain if y != years[0] else 100,
                'Базисный темп %': base
            })
            prev_val = val
            
        df_tr = pd.DataFrame(trend_rows).set_index('Год')
        render_table_header("3", "Трендовый анализ", "чистой прибыли", f"{years[0]}-{years[-1]} гг.")
        st.table(df_tr.style.format("{:,.2f}"))

        # ---------------------------------------------------------
        # ЗАДАНИЕ 4: ФАКТОРНЫЙ АНАЛИЗ
        # ---------------------------------------------------------
        if len(years) >= 2:
            render_task("4", "Факторный анализ прибыли", "Оценка влияния факторов.", f"Анализ {curr_y} к {prev_y} г.")
            
            def g(row, yr): return abs(df_res.loc[row, yr]) # Берем модуль для формул
            
            v0 = {k: g(k, prev_y) for k in df_res.index}
            v1 = {k: g(k, curr_y) for k in df_res.index}
            
            factors = [
                ('Выручка', v1['Выручка'] - v0['Выручка']),
                ('Себестоимость', -(v1['Себестоимость продаж'] - v0['Себестоимость продаж'])),
                ('Упр. расходы', -(v1['Управленческие расходы'] - v0['Управленческие расходы'])),
                ('Комм. расходы', -(v1['Коммерческие расходы'] - v0['Коммерческие расходы'])),
                ('Прочие доходы', v1['Прочие доходы'] - v0['Прочие доходы']),
                ('Прочие расходы', -(v1['Прочие расходы'] - v0['Прочие расходы'])),
                ('Налог на прибыль', -(v1['Налог на прибыль'] - v0['Налог на прибыль']))
            ]
            
            f_rows = []
            tot = 0
            for name, val in factors:
                key = name if name in v0 else name + ' продаж' if name+' продаж' in v0 else name
                # Костыль для сопоставления имен
                if name == 'Себестоимость': key = 'Себестоимость продаж'
                if name == 'Упр. расходы': key = 'Управленческие расходы'
                if name == 'Комм. расходы': key = 'Коммерческие расходы'
                
                f_rows.append([name, v0.get(key, 0), v1.get(key, 0), val])
                tot += val
                
            f_rows.append(['ИТОГО влияние', 0, 0, tot])
            
            df_fact = pd.DataFrame(f_rows, columns=['Фактор', f'Базис ({prev_y})', f'Факт ({curr_y})', 'Влияние'])
            
            render_table_header("4", "Факторный анализ", "чистой прибыли", f"{curr_y} к {prev_y} г.")
            st.table(df_fact.style.format({col: "{:,.2f}" for col in df_fact.columns if col != 'Фактор'}))
            
            if api_key and use_ai: st.info(get_ai_analysis(df_fact, "Факторы прибыли", api_key))

        # ---------------------------------------------------------
        # ЗАДАНИЕ 5: АНАЛИЗ ЗАТРАТ
        # ---------------------------------------------------------
        if len(years) >= 2:
            render_task("5", "Анализ затрат на производство", "Динамика расходов.", f"{curr_y} к {prev_y} г.")
            
            c_items = ['Себестоимость продаж', 'Коммерческие расходы', 'Управленческие расходы']
            df_c = df_res.loc[c_items, [prev_y, curr_y]].apply(abs)
            df_c.loc['ИТОГО'] = df_c.sum()
            
            df_c['Абс. откл.'] = df_c[curr_y] - df_c[prev_y]
            df_c['Темп %'] = (df_c[curr_y] / df_c[prev_y] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            
            # Доля
            tot_p, tot_c = df_c.loc['ИТОГО', prev_y], df_c.loc['ИТОГО', curr_y]
            df_c[f'Доля {prev_y}%'] = (df_c[prev_y]/tot_p*100).fillna(0)
            df_c[f'Доля {curr_y}%'] = (df_c[curr_y]/tot_c*100).fillna(0)
            
            render_table_header("5", "Комплексный анализ", "затрат на производство")
            st.dataframe(df_c.style.format("{:,.2f}"))

        # ---------------------------------------------------------
        # ЗАДАНИЕ 6: CVP
        # ---------------------------------------------------------
        render_task("6", "CVP-анализ", "Точка безубыточности.", "Калькулятор.")
        
        col1, col2 = st.columns(2)
        p = col1.number_input("Цена (P)", 1000.0)
        avc = col1.number_input("VC на ед.", 600.0)
        tfc = col2.number_input("TFC", 200000.0)
        
        if p > avc:
            bep = tfc / (p - avc)
            st.success(f"BEP: {bep:,.0f} шт. | {bep*p:,.2f} руб.")
        else:
            st.error("Убыток с единицы!")

        # ---------------------------------------------------------
        # СКАЧИВАНИЕ
        # ---------------------------------------------------------
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_v.to_excel(writer, sheet_name='Вертикальный')
            if 'df_h' in locals(): df_h.to_excel(writer, sheet_name='Горизонтальный')
            df_tr.to_excel(writer, sheet_name='Трендовый')
            if 'df_fact' in locals(): df_fact.to_excel(writer, sheet_name='Факторный', index=False)
            if 'df_c' in locals(): df_c.to_excel(writer, sheet_name='Затраты')
            
        st.download_button("📥 Скачать Excel", data=output.getvalue(), file_name="report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("👈 Загрузите файлы в меню слева (например, 2021.xlsx, 2022.xlsx, 2023.xlsx...).")
