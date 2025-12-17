import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import re
from openai import OpenAI

# ==========================================
# 1. НАСТРОЙКИ СТРАНИЦЫ
# ==========================================
st.set_page_config(
    page_title="AI Финансовый Анализ",
    page_icon="🤖",
    layout="wide"
)

st.title("📊 АРМ Аналитика: Сбор данных через AI")
st.markdown("### Загрузите файлы, а ИИ сам извлечет из них данные для анализа.")

# ==========================================
# 2. ФУНКЦИИ (AI-ПАРСИНГ)
# ==========================================

def extract_data_with_gpt(file_obj, api_key):
    """
    Отправляет содержимое Excel в GPT и просит вернуть JSON с данными.
    """
    if not api_key:
        return None, None, "Нет API ключа"

    try:
        # 1. Читаем Excel и превращаем в простой текст (CSV)
        # Читаем все листы, ищем тот, где есть слова "Выручка" или код "2110"
        dfs = pd.read_excel(file_obj, sheet_name=None, header=None)
        target_text = ""
        
        for name, df in dfs.items():
            # Конвертируем лист в строку
            text_dump = df.to_csv(index=False, sep='\t')
            # Если похоже на фин. отчет, берем его
            if "2110" in text_dump or "Выручка" in text_dump:
                target_text = text_dump[:5000] # Берем первые 5000 символов (обычно достаточно)
                break
        
        if not target_text:
            # Если не нашли, берем первый лист
            target_text = list(dfs.values())[0].to_csv(index=False, sep='\t')[:5000]

        # 2. Формируем запрос к ИИ
        client = OpenAI(api_key=api_key, base_url="https://openai.api.proxyapi.ru/v1")
        
        system_prompt = """
        Ты — бухгалтерский парсер. Твоя задача — извлечь данные из текста отчета о финансовых результатах (Форма 2).
        1. Найди ГОД отчета (например, "За 2023 г." -> 2023).
        2. Найди значения для следующих кодов строк:
           2110 (Выручка), 2120 (Себестоимость), 2100 (Валовая прибыль),
           2210 (Коммерческие), 2220 (Управленческие), 2200 (Прибыль от продаж),
           2340 (Прочие доходы), 2350 (Прочие расходы), 2410 (Налог на прибыль), 2400 (Чистая прибыль).
        
        Правила:
        - Если число в скобках (100) — это отрицательное число -100.
        - Убери пробелы между разрядами (10 000 -> 10000).
        - Если значения нет, ставь 0.
        - Верни ТОЛЬКО валидный JSON без markdown.
        
        Пример ответа:
        {
            "year": 2023,
            "data": {
                "Выручка": 10000,
                "Себестоимость продаж": -5000,
                ...
            }
        }
        """
        
        user_prompt = f"Извлеки данные из этого текста:\n\n{target_text}"

        response = client.chat.completions.create(
            model="anthropic/claude-sonnet-4-20250514", # или gpt-4o, gpt-3.5-turbo
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0
        )
        
        content = response.choices[0].message.content.strip()
        
        # Очистка от маркдауна ```json ... ```
        if "```" in content:
            content = content.split("```json")[1].split("```")[0].strip()
        elif "```" in content: # просто ```
             content = content.split("```")[1].strip()

        result = json.loads(content)
        return result.get('year'), result.get('data'), None

    except Exception as e:
        return None, None, str(e)

# --- Остальные функции оформления ---

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

def get_ai_analysis_text(table_df, task_context, api_key):
    try:
        client = OpenAI(api_key=api_key, base_url="https://openai.api.proxyapi.ru/v1")
        prompt = f"Ты студент. Контекст: {task_context}. Данные:\n{table_df.to_string()}\nНапиши вывод (3-4 предложения) в академическом стиле на русском."
        response = client.chat.completions.create(
            model="anthropic/claude-sonnet-4-20250514",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e: return f"Ошибка API: {e}"

# ==========================================
# 3. БОКОВАЯ ПАНЕЛЬ
# ==========================================
with st.sidebar:
    st.header("⚙️ Настройки")
    api_key = st.text_input("API Key (ProxyAPI)", type="password")
    use_ai_analysis = st.checkbox("✍️ Добавлять выводы к таблицам", value=True)
    
    st.info("📂 Загрузите файлы:")
    st.caption("Данные будут извлечены с помощью AI.")
    uploaded_files = st.file_uploader("Перетащите файлы сюда", type=["xlsx"], accept_multiple_files=True)

# ==========================================
# 4. ОСНОВНАЯ ЛОГИКА
# ==========================================

GLOBAL_DATA = {} # {2022: {'Выручка': 100, ...}}

if uploaded_files:
    if not api_key:
        st.error("🛑 Для автоматического извлечения данных через AI необходим API ключ!")
    else:
        # ПАРСИНГ ФАЙЛОВ ЧЕРЕЗ AI
        with st.status("🤖 ИИ анализирует файлы...", expanded=True) as status:
            for file in uploaded_files:
                st.write(f"Обработка {file.name}...")
                year, data, error = extract_data_with_gpt(file, api_key)
                
                if year and data:
                    GLOBAL_DATA[year] = data
                    st.write(f"✅ {file.name}: Год {year} найден.")
