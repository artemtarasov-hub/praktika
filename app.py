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
    page_title="Комплексный анализ (Динамический)",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Практическая работа: Комплексный анализ (Многолетний)")

# ==========================================
# 2. ФУНКЦИИ
# ==========================================

def extract_year_from_filename(filename):
    """Ищет 4 цифры (год) в названии файла."""
    match = re.search(r'\d{4}', filename)
    if match:
        return int(match.group(0))
    return None

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

def get_single_value_by_code(df, code):
    """Ищет первое значение в строке с кодом."""
    if df is None: return 0
    for index, row in df.iterrows():
        for i, cell in enumerate(row):
            try:
                if pd.to_numeric(cell, errors='coerce') == code:
                    # Ищем первое число справа
                    for next_cell in row[i+1:]:
                        if pd.notna(next_cell) and str(next_cell).strip() not in ['', '-', '(-)']:
                            val_str = str(next_cell).replace(' ', '').replace('\xa0', '')
                            if val_str.startswith('(') and val_str.endswith(')'):
                                val_str = '-' + val_str[1:-1]
                            val = pd.to_numeric(val_str, errors='coerce')
                            if pd.notna(val): return val
            except: continue
    return 0

def load_excel_sheet(file):
    try:
        dfs = pd.read_excel(file, sheet_name=None, header=None)
        for name, df in dfs.items():
            if 'фин' in name.lower() or 'результ' in name.lower(): return df
        if len(dfs) >= 3: return list(dfs.values())[2]
        return list(dfs.values())[0]
    except: return None

# ==========================================
# 3. БОКОВАЯ ПАНЕЛЬ (ЗАГРУЗКА МНОГИХ ФАЙЛОВ)
# ==========================================
with st.sidebar:
    st.header("⚙️ Настройки")
    api_key = st.text_input("API Key (ProxyAPI)", type="password")
    use_ai = st.checkbox("✍️ Добавлять выводы ИИ", value=True)
    
    st.info("📂 Загрузите файлы за ВСЕ годы (например, 2015.xlsx ... 2024.xlsx):")
    uploaded_files = st.file_uploader("Перетащите файлы сюда", type=["xlsx"], accept_multiple_files=True)

# ==========================================
# 4. СБОРКА ОБЩЕЙ ТАБЛИЦЫ
# ==========================================

df_res = pd.DataFrame()

if uploaded_files:
    data_store = {} # {Year: {Indicator: Value}}
    
    codes = {
        'Выручка': 2110, 'Себестоимость продаж': 2120, 'Валовая прибыль': 2100,
        'Коммерческие расходы': 2210, 'Управленческие расходы': 2220, 'Прибыль от продаж': 2200,
        'Прочие доходы': 2340, 'Прочие расходы': 2350, 'Налог на прибыль': 2410, 
        'Чистая прибыль': 2400
    }

    # Чтение файлов
    for file in uploaded_files:
        year = extract_year_from_filename(file.name)
        if year is None:
            st.error(f"⚠️ Не удалось определить год в названии файла: {file.name}. Переименуйте файл (например 'Otchet_2020.xlsx').")
            continue
            
        df_sheet = load_excel_sheet(file)
        if df_sheet is not None:
            year_data = {}
            for name, code in codes.items():
                year_data[name] = get_single_value_by_code(df_sheet, code)
            data_store[year] = year_data

    # Создание DataFrame
    if data_store:
        df_res = pd.DataFrame(data_store).sort_index(axis=1) # Сортируем колонки по годам (2015, 2016...)
        years = df_res.columns.tolist() # Список доступных лет
        
        if len(years) < 2:
            st.warning("⚠️ Загружено менее 2-х лет. Для анализа нужно минимум 2 года.")
        else:
            # Определяем Базовый (последний) и Предыдущий год для стандартных таблиц
            current_year = years[-1]
            prev_year = years[-2]
            base_period_str = f"{years[0]}-{years[-1]} гг."
            
            st.success(f"✅ Данные успешно загружены за период: {base_period_str}. (Всего лет: {len(years)})")

            # ==========================================
            # ЗАДАНИЕ 1: ВЕРТИКАЛЬНЫЙ (Последние 3 года)
            # ==========================================
            years_for_vert = years[-3:] if len(years) >= 3 else years # Берем последние 3 года или сколько есть
            period_vert = f"{years_for_vert[0]}-{years_for_vert[-1]} гг."

            render_task("1", "Анализ структуры", "Изучение структуры доходов и расходов.", f"Провести вертикальный анализ за последние доступные годы ({period_vert}).")
            
            df_v = df_res[years_for_vert].copy()
            for y in years_for_vert:
                base_val = df_v.loc['Выручка', y]
                df_v[f'Уд. вес {y} (%)'] = (df_v[y] / base_val * 100).fillna(0)
            
            # Сортировка колонок для красивого вывода
            cols_v = []
            for y in years_for_vert:
                cols_v.append(y)
                cols_v.append(f'Уд. вес {y} (%)')
            df_v = df_v[cols_v]
            
            render_table_header("1", "Вертикальный анализ", "финансовых результатов", period_vert)
            st.dataframe(df_v.style.format("{:,.2f}"))
            if api_key and use_ai: st.info(get_ai_analysis(df_v, "Структура финансовых результатов", api_key))

            # ==========================================
            # ЗАДАНИЕ 2: ГОРИЗОНТАЛЬНЫЙ (Последние 2-3 года)
            # ==========================================
            # Сравниваем Current vs Prev и Current vs Pre-Prev (как в задании 2024-2023 и 2024-2022)
            render_task("2", "Анализ динамики", "Оценка темпов роста.", f"Провести горизонтальный анализ относительно {current_year} года.")
            
            df_h = df_res[years_for_vert].copy()
            cols_h_display = years_for_vert.copy()
            
            # Расчет отклонений относительно ПОСЛЕДНЕГО года (current_year)
            # Идем по годам, кроме последнего
            for y in years_for_vert[:-1]:
                diff_col = f'Откл. {current_year}-{y}'
                rate_col = f'Темп роста {current_year}/{y} (%)'
                
                df_h[diff_col] = df_h[current_year] - df_h[y]
                df_h[rate_col] = (df_h[current_year] / df_h[y] * 100).replace([np.inf, -np.inf], 0).fillna(0)
                
                cols_h_display.extend([diff_col, rate_col])
                
            render_table_header("2", "Горизонтальный анализ", "финансовых результатов", period_vert)
            st.dataframe(df_h[cols_h_display].style.format("{:,.2f}"))
            if api_key and use_ai: st.info(get_ai_analysis(df_h, "Динамика показателей", api_key))

            # ==========================================
            # ЗАДАНИЕ 3: ТРЕНДОВЫЙ (ЗА 10 ЛЕТ ИЛИ СКОЛЬКО ЕСТЬ)
            # ==========================================
            render_task("3", "Трендовый анализ", "Анализ тенденций за весь период.", f"Трендовый анализ Чистой прибыли за {base_period_str}.")
            
            trend_data = []
            base_of_all = df_res.loc['Чистая прибыль', years[0]] # Самый первый год
            prev = None
            
            for y in years:
                curr = df_res.loc['Чистая прибыль', y]
                
                # Цепной (к прошлому году)
                abs_ch = (curr - prev) if prev is not None else 0
                rate_ch = (curr / prev * 100) if (prev and prev != 0) else 100.0
                
                # Базисный (к самому первому году)
                rate_base = (curr / base_of_all * 100) if base_of_all != 0 else 0
                
                trend_data.append({
                    'Год': y,
                    'Чистая прибыль': curr,
                    'Абс. откл. (цепное)': abs_ch if y != years[0] else 0,
                    'Темп (цепной) %': rate_ch,
                    'Темп (базисный к ' + str(years[0]) + ') %': rate_base
                })
                prev = curr
            
            df_trend = pd.DataFrame(trend_data).set_index('Год')
            # Транспонируем для компактности, если лет много, или оставляем так. Таблица тренда обычно вертикальная (годы в строках).
            
            render_table_header("3", "Трендовый анализ", "чистой прибыли", base_period_str)
            st.table(df_trend.style.format("{:,.2f}"))

            # ==========================================
            # ЗАДАНИЕ 4: ФАКТОРНЫЙ (ПОСЛЕДНИЕ 2 ГОДА)
            # ==========================================
            render_task("4", "Факторный анализ", "Влияние факторов на прибыль.", f"Анализ за отчетный {current_year} год по сравнению с {prev_year}.")
            
            # Берем данные двух последних лет
            v0 = {k: abs(df_res.loc[k, prev_year]) for k in df_res.index}
            v1 = {k: abs(df_res.loc[k, current_year]) for k in df_res.index}
            
            # Расчет
            factors = [
                ('Выручка', v1['Выручка'] - v0['Выручка']),
                ('Себестоимость', -(v1['Себестоимость продаж'] - v0['Себестоимость продаж'])),
                ('Упр. расходы', -(v1['Управленческие расходы'] - v0['Управленческие расходы'])),
                ('Комм. расходы', -(v1['Коммерческие расходы'] - v0['Коммерческие расходы'])),
                ('Прочие доходы', v1['Прочие доходы'] - v0['Прочие доходы']),
                ('Прочие расходы', -(v1['Прочие расходы'] - v0['Прочие расходы'])),
                ('Налог на прибыль', -(v1['Налог на прибыль'] - v0['Налог на прибыль']))
            ]
            
            factor_rows = []
            total_inf = 0
            for name, val in factors:
                factor_rows.append([name, v0.get(name) or v0.get(name+' продаж', 0), v1.get(name) or v1.get(name+' продаж', 0), val])
                total_inf += val
                
            factor_rows.append(['ИТОГО влияние', 0, 0, total_inf])
            factor_rows.append(['Изм. ЧП (Факт)', v0['Чистая прибыль'], v1['Чистая прибыль'], v1['Чистая прибыль']-v0['Чистая прибыль']])
            
            df_fact = pd.DataFrame(factor_rows, columns=['Фактор', f'Базис ({prev_year})', f'Факт ({current_year})', 'Влияние'])
            
            render_table_header("4", "Факторный анализ", "чистой прибыли", f"{current_year} к {prev_year} г.")
            st.table(df_fact.style.format({col: "{:,.2f}" for col in df_fact.columns if col != 'Фактор'}))
            if api_key and use_ai: st.info(get_ai_analysis(df_fact, "Факторный анализ", api_key))

            # ==========================================
            # ЗАДАНИЕ 5: ЗАТРАТЫ (ПОСЛЕДНИЕ 2 ГОДА)
            # ==========================================
            render_task("5", "Анализ затрат", "Динамика и структура.", f"Анализ затрат за {current_year} и {prev_year} гг.")
            
            cost_cols = ['Себестоимость продаж', 'Коммерческие расходы', 'Управленческие расходы']
            df_costs = df_res.loc[cost_cols, [prev_year, current_year]].apply(abs).copy()
            df_costs.loc['ИТОГО'] = df_costs.sum()
            
            df_costs['Абс. откл.'] = df_costs[current_year] - df_costs[prev_year]
            df_costs['Темп роста %'] = (df_costs[current_year] / df_costs[prev_year] * 100).replace([np.inf], 0)
            
            tot_p = df_costs.loc['ИТОГО', prev_year]
            tot_c = df_costs.loc['ИТОГО', current_year]
            df_costs[f'Уд. вес {prev_year} %'] = (df_costs[prev_year] / tot_p * 100).fillna(0)
            df_costs[f'Уд. вес {current_year} %'] = (df_costs[current_year] / tot_c * 100).fillna(0)
            
            render_table_header("5", "Комплексный анализ", "затрат", f"{current_year} к {prev_year} г.")
            st.dataframe(df_costs.style.format("{:,.2f}"))

            # ==========================================
            # ЗАДАНИЕ 6: CVP (Калькулятор)
            # ==========================================
            render_task("6", "CVP-анализ", "Точка безубыточности.", "Рассчитать BEP (ввод данных вручную).")
            
            cvp_type = st.radio("Тип:", ["Однопродуктовое", "Многопродуктовое"], horizontal=True)
            if cvp_type == "Однопродуктовое":
                c1, c2 = st.columns(2)
                p = c1.number_input("Цена (P)", 1000.0)
                avc = c1.number_input("VC на ед.", 600.0)
                tfc = c2.number_input("TFC (Пост. затраты)", 200000.0)
                q = c2.number_input("Объем (Q)", 1000.0)
                md = p - avc
                if md > 0:
                    bep = tfc/md
                    st.table(pd.DataFrame([
                        ["BEP (шт)", bep], 
                        ["BEP (руб)", bep*p], 
                        ["Запас прочности (%)", ((q*p - bep*p)/(q*p)*100) if q else 0]
                    ], columns=["Показатель", "Значение"]).style.format({"Значение": "{:,.2f}"}))
                else: st.error("Убыток с единицы!")
            else:
                st.write("3 товара:")
                tfc_m = st.number_input("Общие TFC", 150000.0)
                prods = []
                cols = st.columns(3)
                for i in range(3):
                    with cols[i]:
                        r = st.number_input(f"Выручка {i+1}", 100000.0)
                        v = st.number_input(f"VC {i+1}", 60000.0)
                        prods.append((r,v))
                if st.button("Рассчитать"):
                    tot_r = sum(x[0] for x in prods)
                    w_k = sum([(r-v)/r * (r/tot_r) for r,v in prods if r > 0])
                    st.success(f"Точка безубыточности: {tfc_m/w_k:,.2f} руб.")

            # СКАЧИВАНИЕ
            st.markdown("---")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_v.to_excel(writer, sheet_name='Вертикальный')
                df_h.to_excel(writer, sheet_name='Горизонтальный')
                df_trend.to_excel(writer, sheet_name='Трендовый')
                df_fact.to_excel(writer, sheet_name='Факторный')
                df_costs.to_excel(writer, sheet_name='Затраты')
            st.download_button("📥 Скачать Excel", data=output.getvalue(), file_name="full_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("👈 Загрузите файлы в меню слева (например, 2020.xlsx, 2021.xlsx...). Год берется из названия файла!")
