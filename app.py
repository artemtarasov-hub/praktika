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

# ЕДИНСТВЕННЫЙ ЗАГОЛОВОК ПРАКТИЧЕСКОЙ РАБОТЫ (КАК БЫЛО)
st.title("📊 Практическая работа №1: Анализ финансовых результатов")

# ==========================================
# 2. ФУНКЦИИ (ОФОРМЛЕНИЕ И ЛОГИКА)
# ==========================================

def render_task(task_num, topic, goal, task_text):
    """
    Выводит блок 'Задание' (вместо Практической работы).
    """
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
    header_text = f"<b>Таблица {table_num}.</b> {analysis_full_name} {subject_genitive} {period}"
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
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e: return f"❌ Ошибка API: {e}"

def load_excel_sheet(file):
    try:
        dfs = pd.read_excel(file, sheet_name=None, header=None)
        # Ищем лист по ключевым словам
        for name, df in dfs.items():
            if 'фин' in name.lower() or 'результ' in name.lower() or 'форма 2' in name.lower(): return df
        # Или ищем код 2110 внутри листа
        for name, df in dfs.items():
            s = df.astype(str).to_string()
            if '2110' in s: return df
        return list(dfs.values())[0]
    except: return None

def detect_year_in_df(df):
    """Ищет год (2020-2030) в шапке файла."""
    if df is None: return None
    header_part = df.head(20).astype(str).to_string()
    years = re.findall(r'202[0-9]', header_part)
    if years:
        return max([int(y) for y in years])
    return None

def get_values_by_code(df, code):
    """Возвращает пару значений (Current, Previous) для найденного кода."""
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
    
    st.info("📂 Исходные данные (Excel):")
    # МУЛЬТИ-ЗАГРУЗКА
    uploaded_files = st.file_uploader(
        "Загрузите файлы отчетов (xlsx)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )

# ==========================================
# 4. ОБРАБОТКА ДАННЫХ И ВЫВОД
# ==========================================

codes_map = {
    'Выручка': 2110, 'Себестоимость продаж': 2120, 'Валовая прибыль': 2100,
    'Коммерческие расходы': 2210, 'Управленческие расходы': 2220, 'Прибыль от продаж': 2200,
    'Прочие доходы': 2340, 'Прочие расходы': 2350, 'Налог на прибыль': 2410, 
    'Чистая прибыль': 2400
}

if uploaded_files:
    master_data = {} 
    
    # 1. Считываем все файлы
    for file in uploaded_files:
        df_raw = load_excel_sheet(file)
        if df_raw is not None:
            detected_year = detect_year_in_df(df_raw)
            # Если год не нашли внутри, пробуем из имени файла
            if not detected_year:
                fname_years = re.findall(r'202[0-9]', file.name)
                if fname_years: detected_year = int(max(fname_years))
                else: continue # Пропускаем, если год неизвестен
            
            year_curr = detected_year
            year_prev = detected_year - 1
            
            for metric, code in codes_map.items():
                val_curr, val_prev = get_values_by_code(df_raw, code)
                
                if year_curr not in master_data: master_data[year_curr] = {}
                master_data[year_curr][metric] = val_curr
                
                # Записываем прошлый год, только если его еще нет (приоритет у свежих отчетов)
                if year_prev not in master_data: master_data[year_prev] = {}
                if metric not in master_data[year_prev]:
                    master_data[year_prev][metric] = val_prev

    # 2. Формируем единый DataFrame
    if master_data:
        df_res = pd.DataFrame(master_data).sort_index(axis=1)
        df_res = df_res.reindex(codes_map.keys())
        years_avail = sorted([str(y) for y in df_res.columns])
        
        # Определяем базовые года (последние два)
        if len(years_avail) >= 2:
            last_year = years_avail[-1]
            prev_year = years_avail[-2]
        else:
            last_year = years_avail[0]
            prev_year = years_avail[0]

        # ---------------------------------------------------------
        # ЗАДАНИЕ 1: ВЕРТИКАЛЬНЫЙ АНАЛИЗ
        # ---------------------------------------------------------
        render_task(
            task_num="1",
            topic="Анализ структуры финансовых результатов",
            goal="Изучение структуры и структурной динамики доходов и расходов организации.",
            task_text=f"На основе данных годовой отчетности провести <b>вертикальный сравнительный анализ</b> финансовых результатов за {', '.join(years_avail)} гг. Рассчитать удельные веса показателей к Выручке."
        )

        df_v = df_res.copy()
        display_cols = []
        for y in df_res.columns:
            base_val = df_v.loc['Выручка', y]
            df_v[f'Уд. вес {y} (%)'] = (df_v[y] / base_val * 100).fillna(0)
            display_cols.extend([y, f'Уд. вес {y} (%)'])
        
        df_v_display = df_v[display_cols]
        render_table_header("1", "Вертикальный сравнительный анализ", "финансовых результатов")
        st.dataframe(df_v_display.style.format("{:,.2f}"))
        
        if api_key and use_ai:
            st.info(get_ai_analysis(df_v_display, "Вывод по структуре доходов и расходов", api_key))

        # ---------------------------------------------------------
        # ЗАДАНИЕ 2: ГОРИЗОНТАЛЬНЫЙ АНАЛИЗ
        # ---------------------------------------------------------
        if len(years_avail) >= 2:
            render_task(
                task_num="2",
                topic="Анализ динамики прибыли",
                goal="Оценка темпов изменения показателей финансовых результатов.",
                task_text=f"Провести <b>горизонтальный сравнительный анализ</b>. Рассчитать абсолютные отклонения и темпы роста показателей за {prev_year}-{last_year} гг. (базисным методом относительно {last_year} года)."
            )

            df_h = df_res[[int(prev_year), int(last_year)]].copy()
            y1, y2 = int(prev_year), int(last_year)
            
            df_h[f'Откл. {y2}-{y1}'] = df_h[y2] - df_h[y1]
            df_h[f'Темп роста {y2}/{y1} (%)'] = (df_h[y2] / df_h[y1] * 100).replace([np.inf, -np.inf], 0).fillna(0)
            
            render_table_header("2", "Горизонтальный сравнительный анализ", "финансовых результатов")
            st.dataframe(df_h.style.format("{:,.2f}"))
            
            if api_key and use_ai:
                st.info(get_ai_analysis(df_h, "Вывод по динамике прибыли", api_key))
        else:
            st.warning("Для горизонтального анализа требуется минимум 2 года данных.")

        # ---------------------------------------------------------
        # ЗАДАНИЕ 3: ТРЕНДОВЫЙ АНАЛИЗ
        # ---------------------------------------------------------
        render_task(
            task_num="3",
            topic="Трендовый анализ показателей",
            goal="Выявление основной тенденции динамики показателя.",
            task_text="Составить таблицу <b>трендового анализа</b> Чистой прибыли за доступный период. Рассчитать цепные и базисные темпы роста."
        )

        trend_data = []
        base_val_start = df_res.loc['Чистая прибыль', df_res.columns[0]]
        prev_val = None
        
        for y in df_res.columns:
            curr = df_res.loc['Чистая прибыль', y]
            abs_ch = (curr - prev_val) if prev_val is not None else 0
            rate_ch = (curr / prev_val * 100) if (prev_val and prev_val != 0) else 100.0
            rate_bs = (curr / base_val_start * 100) if base_val_start != 0 else 0
            
            trend_data.append({
                'Год': y,
                'Чистая прибыль': curr,
                'Абс. откл. (цепное)': abs_ch if y != df_res.columns[0] else 0,
                'Темп (цепной) %': rate_ch if y != df_res.columns[0] else 100,
                'Темп (базисный) %': rate_bs
            })
            prev_val = curr
            
        df_trend = pd.DataFrame(trend_data).set_index('Год')
        render_table_header("3", "Трендовый анализ", "чистой прибыли")
        st.table(df_trend.style.format("{:,.2f}"))

        # ---------------------------------------------------------
        # ЗАДАНИЕ 4: ФАКТОРНЫЙ АНАЛИЗ
        # ---------------------------------------------------------
        if len(years_avail) >= 2:
            render_task(
                task_num="4",
                topic="Факторный анализ прибыли",
                goal="Оценка влияния факторов на изменение результативного показателя.",
                task_text=f"Провести <b>факторный анализ</b> Чистой прибыли методом цепных подстановок ({last_year} к {prev_year}). <br>Модель: <i>ЧП = Выручка - Себестоимость - Упр. - Комм. + Прочие Дох. - Прочие Расх. - Налог</i>."
            )

            def get_abs(row, yr): return abs(df_res.loc[row, int(yr)])
            v0 = {k: get_abs(k, prev_year) for k in df_res.index}
            v1 = {k: get_abs(k, last_year) for k in df_res.index}
            
            factors = [
                ['Выручка', v0['Выручка'], v1['Выручка'], v1['Выручка'] - v0['Выручка']],
                ['Себестоимость', v0['Себестоимость продаж'], v1['Себестоимость продаж'], -(v1['Себестоимость продаж'] - v0['Себестоимость продаж'])],
                ['Упр. расходы', v0['Управленческие расходы'], v1['Управленческие расходы'], -(v1['Управленческие расходы'] - v0['Управленческие расходы'])],
                ['Комм. расходы', v0['Коммерческие расходы'], v1['Коммерческие расходы'], -(v1['Коммерческие расходы'] - v0['Коммерческие расходы'])],
                ['Прочие доходы', v0['Прочие доходы'], v1['Прочие доходы'], v1['Прочие доходы'] - v0['Прочие доходы']],
                ['Прочие расходы', v0['Прочие расходы'], v1['Прочие расходы'], -(v1['Прочие расходы'] - v0['Прочие расходы'])],
                ['Налог на прибыль', v0['Налог на прибыль'], v1['Налог на прибыль'], -(v1['Налог на прибыль'] - v0['Налог на прибыль'])]
            ]
            
            total_inf = sum([r[3] for r in factors])
            np_0, np_1 = v0['Чистая прибыль'], v1['Чистая прибыль']
            factors.append(['ИТОГО влияние', 0, 0, total_inf])
            factors.append(['Изм. ЧП (Факт)', np_0, np_1, np_1 - np_0])
            
            df_fact = pd.DataFrame(factors, columns=['Фактор', f'Базис ({prev_year})', f'Факт ({last_year})', 'Влияние'])
            
            render_table_header("4", "Факторный анализ", "чистой прибыли", f"{last_year} к {prev_year} г.")
            st.table(df_fact.style.format({f'Базис ({prev_year})': "{:,.2f}", f'Факт ({last_year})': "{:,.2f}", 'Влияние': "{:,.2f}"}))
            
            if api_key and use_ai:
                st.info(get_ai_analysis(df_fact, "Какие факторы снизили или увеличили прибыль?", api_key))

        # ---------------------------------------------------------
        # ЗАДАНИЕ 5: АНАЛИЗ ЗАТРАТ
        # ---------------------------------------------------------
        if len(years_avail) >= 2:
            render_task(
                task_num="5",
                topic="Анализ затрат на производство",
                goal="Оценка динамики и структуры расходов по обычным видам деятельности.",
                task_text="Провести горизонтальный и вертикальный анализ затрат (Себестоимость + Коммерческие + Управленческие)."
            )

            cost_cols = ['Себестоимость продаж', 'Коммерческие расходы', 'Управленческие расходы']
            y1, y2 = int(prev_year), int(last_year)
            df_costs = df_res.loc[cost_cols, [y1, y2]].apply(abs)
            df_costs.loc['ИТОГО'] = df_costs.sum()
            
            df_costs['Абс. откл.'] = df_costs[y2] - df_costs[y1]
            df_costs['Темп роста %'] = (df_costs[y2] / df_costs[y1] * 100).replace([np.inf], 0)
            df_costs[f'Уд. вес {y1} %'] = (df_costs[y1] / df_costs.loc['ИТОГО', y1] * 100).fillna(0)
            df_costs[f'Уд. вес {y2} %'] = (df_costs[y2] / df_costs.loc['ИТОГО', y2] * 100).fillna(0)
            
            render_table_header("5", "Комплексный анализ", "затрат на производство", f"{last_year} к {prev_year} г.")
            st.dataframe(df_costs.style.format("{:,.2f}"))

        # ---------------------------------------------------------
        # ЗАДАНИЕ 6: CVP АНАЛИЗ
        # ---------------------------------------------------------
        render_task(
            task_num="6",
            topic="CVP-анализ (Анализ безубыточности)",
            goal="Определение точки безубыточности и запаса финансовой прочности.",
            task_text="Рассчитать точку безубыточности в натуральном и денежном выражении."
        )

        cvp_type = st.radio("Вариант задания:", ["Однопродуктовое", "Многопродуктовое"], horizontal=True)
        
        if cvp_type == "Однопродуктовое":
            c1, c2 = st.columns(2)
            p = c1.number_input("Цена (P)", 1000.0)
            avc = c1.number_input("Перем. затраты (AVC)", 600.0)
            # Берем управленческие расходы последнего года как подсказку для TFC
            def_tfc = abs(float(df_res.loc['Управленческие расходы', int(last_year)])) if len(years_avail)>0 else 200000.0
            tfc = c2.number_input("Пост. затраты (TFC)", def_tfc)
            q = c2.number_input("Объем (Q)", 1000.0)
            
            md = p - avc
            if md > 0:
                bep = tfc / md
                margin = (q * p) - (bep * p)
                data_c = [["BEP (шт)", bep], ["BEP (руб)", bep*p], ["Запас прочности", margin]]
                df_cvp = pd.DataFrame(data_c, columns=["Показатель", "Значение"])
                render_table_header("6", "CVP-анализ", "безубыточности")
                st.table(df_cvp.style.format({"Значение": "{:,.2f}"}))
            else:
                st.error("Цена должна быть больше переменных затрат!")
        else:
            st.write("Введите данные для 3-х товаров:")
            tfc_m = st.number_input("Общие TFC", 150000.0)
            prods = []
            cols = st.columns(3)
            for i in range(3):
                with cols[i]:
                    r = st.number_input(f"Выручка {i+1}", 100000.0)
                    v = st.number_input(f"VC {i+1}", 60000.0)
                    prods.append((r,v))
            if st.button("Рассчитать CVP"):
                tot_r = sum(x[0] for x in prods)
                w_k = sum([(r-v)/r * (r/tot_r) for r,v in prods if r > 0])
                bep_tot = tfc_m / w_k if w_k else 0
                st.success(f"Точка безубыточности: {bep_tot:,.2f} руб.")

        # ---------------------------------------------------------
        # КНОПКА СКАЧИВАНИЯ
        # ---------------------------------------------------------
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_res.to_excel(writer, sheet_name='Сводные_Данные')
            df_v_display.to_excel(writer, sheet_name='Вертикальный')
            if 'df_h' in locals(): df_h.to_excel(writer, sheet_name='Горизонтальный')
            if 'df_trend' in locals(): df_trend.to_excel(writer, sheet_name='Трендовый')
            if 'df_fact' in locals(): df_fact.to_excel(writer, sheet_name='Факторный', index=False)
            if 'df_costs' in locals(): df_costs.to_excel(writer, sheet_name='Затраты')
        
        st.download_button(
            "📥 Скачать полный отчет (Excel)", 
            data=output.getvalue(), 
            file_name="financial_analysis_report.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
else:
    st.info("👈 Загрузите файлы в меню слева (можно выбрать несколько).")
