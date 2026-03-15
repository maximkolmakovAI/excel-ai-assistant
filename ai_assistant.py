import os
import json
import re
import traceback
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import io
import base64
import httpx

MODEL = "openai/gpt-4.1"

PROXY_API_KEY = os.environ.get("PROXY_API_KEY", "sk-2dbPPD0-PWa-1GUgVfonjg")
PROXY_BASE_URL = os.environ.get("PROXY_BASE_URL", "https://litellm.1bitai.ru")

SYSTEM_PROMPT = """Ты — ИИ-аналитик данных, работающий с загруженными Excel-файлами (листы: Заявки, Заказы, Предприятия).

Тебе будут предоставлены:
1. Описание схемы данных (листы, колонки, типы, примеры значений)
2. Дополнительный контекст (информация о филиалах из префиксов номеров)
3. История диалога
4. Новый вопрос пользователя

Твоя задача — сгенерировать исполняемый Python-код с использованием pandas и matplotlib, который отвечает на вопрос.

ВАЖНЫЕ ПРАВИЛА:
- Переменные с данными уже доступны: `df_zayavki` (Заявки), `df_zakazy` (Заказы), `df_predpriyatiya` (Предприятия)
- Для ответа тебе нужно установить одну или несколько из следующих переменных:
  * `result_text` (str) — краткий текстовый ответ
  * `result_df` (pd.DataFrame) — табличные данные для отображения
  * `result_fig` (matplotlib.figure.Figure) — график/диаграмма
- ВСЕГДА устанавливай `result_text`
- Используй `result_df` для таблиц и `result_fig` для визуализации

КОНТЕКСТ О ДАННЫХ:
- В листе "Заявки": колонка "Заказано" = 100 означает, что по заявке был получен заказ (значение 0 = нет заказа)
- Колонка "Сумма" в Заявках — сумма в рублях
- Номер заявки начинается с кода подразделения: SPB=СПб, MSK=Москва, EKB=Екатеринбург, SMR=Самара, KRD=Краснодар, KZN=Казань
- В Заказах колонка "Подразделение" содержит: Москва, Москва 2, Екатеринбург, Казань, Краснодар, Самара, Главный отдел продаж, Отдел продаж 3/5/7/8, Отдел приводной техники, Отдел продаж запчастей
- Московский филиал — это подразделения "Москва" и "Москва 2"; также в Заявках — префикс MSK
- Центральный офис (СПб) имеет несколько отделов продаж
- Если колонок "Интересен ОЗ" или "Рук." нет в данных — сообщи об этом пользователю

ПРАВИЛА ГЕНЕРАЦИИ КОДА:
- Не используй plt.show()
- Всегда указывай figsize для фигур: fig, ax = plt.subplots(figsize=(12, 6))
- Используй ax (оси) для рисования вместо plt.plot(), plt.bar() и т.д.
- Обязательно установи result_fig = fig после создания графика
- Не забывай plt.close() только если явно создаёшь несколько фигур последовательно
- Форматируй числа в русской нотации (пробел как разделитель тысяч)
- Для графиков: используй читаемые шрифты, подписи осей, заголовки на русском языке
- Обрабатывай случаи с пустыми данными (проверяй .empty)
- НЕ используй f-строки с кавычками внутри без экранирования
- Используй matplotlib с русскими подписями (шрифт поддерживает Unicode)
- Для гистограмм с накоплением используй ax.bar(..., stacked=True) или сразу df.plot(kind='bar', stacked=True, ax=ax)
- ВАЖНО: tick_params() не поддерживает 'ha', 'va' — для поворота подписей осей используй ax.set_xticklabels(..., rotation=45) или ax.tick_params(axis='x', labelrotation=45)
- НЕ передавай 'ha' в ax.tick_params() — используй ax.set_xticklabels() с параметром ha вместо этого
- Для лучшей читаемости графиков вращай подписи осей: ax.set_xticklabels(..., rotation=45, ha='right')

Отвечай ТОЛЬКО валидным JSON в формате:
{
  "explanation": "краткое объяснение что делает код",
  "code": "python код"
}
"""


def build_context_prompt(schema_summary: str, prefix_info: str) -> str:
    return f"""Схема данных:
{schema_summary}

{prefix_info}
"""


def ask_ai(
    question: str,
    schema_summary: str,
    prefix_info: str,
    history: list[dict],
    sheets: dict,
) -> dict:
    """
    Ask the AI assistant a question about the loaded data.
    Returns: {text, df, fig, error, code}
    """
    context = build_context_prompt(schema_summary, prefix_info)
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": f"Контекст о данных:\n{context}"},
        {"role": "assistant", "content": "Понял. Готов анализировать данные."},
    ]

    for turn in history[-8:]:
        messages.append({"role": "user", "content": turn["user"]})
        messages.append({"role": "assistant", "content": turn["assistant_raw"]})

    messages.append({"role": "user", "content": question})

    try:
        with httpx.Client(timeout=120.0) as client:
            response = client.post(
                f"{PROXY_BASE_URL}/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {PROXY_API_KEY}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": MODEL,
                    "messages": messages,
                    "max_completion_tokens": 8192,
                },
            )
        response.raise_for_status()
        data = response.json()
        raw = data.get("choices", [{}])[0].get("message", {}).get("content", "{}")
        
        # Remove control characters and clean JSON
        raw = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', raw)
        
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as e:
            # Try to fix common issues
            raw = raw.replace('\n', '\\n').replace('\r', '\\r')
            try:
                parsed = json.loads(raw)
            except:
                return {"text": f"Ошибка обработки ответа ИИ: {e}", "df": None, "fig": None, "error": str(e), "code": ""}
        
        code = parsed.get("code", "")
        explanation = parsed.get("explanation", "")
    except Exception as e:
        return {"text": f"Ошибка обращения к ИИ: {e}", "df": None, "fig": None, "error": str(e), "code": ""}

    if not code:
        return {"text": explanation or "Не удалось сформировать ответ.", "df": None, "fig": None, "error": None, "code": code}

    result = _execute_code(code, sheets)
    result["explanation"] = explanation
    result["code"] = code
    return result


def _execute_code(code: str, sheets: dict) -> dict:
    """Safely execute generated pandas code and return results."""
    df_zayavki = sheets.get("Заявки", pd.DataFrame())
    df_zakazy = sheets.get("Заказы", pd.DataFrame())
    df_predpriyatiya = sheets.get("Предприятия", pd.DataFrame())

    plt.close("all")

    namespace = {
        "__builtins__": __builtins__,
        "pd": pd,
        "plt": plt,
        "df_zayavki": df_zayavki.copy(),
        "df_zakazy": df_zakazy.copy(),
        "df_predpriyatiya": df_predpriyatiya.copy(),
        "result_text": None,
        "result_df": None,
        "result_fig": None,
    }

    try:
        exec(code, namespace)
    except Exception as e:
        tb = traceback.format_exc()
        return {
            "text": f"Ошибка выполнения кода: {e}",
            "df": None,
            "fig": None,
            "error": tb,
        }

    text = namespace.get("result_text")
    df = namespace.get("result_df")
    fig = namespace.get("result_fig")

    if text is None:
        text = "Анализ выполнен."

    if isinstance(df, pd.Series):
        df = df.reset_index()
        df.columns = ["Категория", "Значение"]

    return {"text": str(text), "df": df, "fig": fig, "error": None}


def fig_to_bytes(fig) -> bytes:
    """Convert matplotlib figure to PNG bytes."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf.read()


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert DataFrame to Excel bytes."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Результат")
    buf.seek(0)
    return buf.read()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Convert DataFrame to CSV bytes (UTF-8 with BOM for Excel compatibility)."""
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
