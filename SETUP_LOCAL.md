# Локальное развёртывание ИИ-ассистента по Excel

## 3 шага к старту

### Шаг 1: Подготовка файлов

Скопируйте эти 6 файлов из Replit проекта в локальную папку:

```
your-project/
├── app.py                      # Главное приложение Streamlit
├── data_processor.py           # Парсинг Excel-файлов
├── ai_assistant.py             # Интеграция с ИИ
├── requirements.txt            # Зависимости
├── .streamlit/
│   └── config.toml             # Конфиг Streamlit
└── attached_assets/
    └── Сентябрь_2020_*.xlsx    # (опционально) Тестовый файл
```

**Где их найти:**
- Все .py файлы — в корне Replit
- `requirements.txt` — в корне
- `.streamlit/config.toml` — в скрытой папке `.streamlit/`
- Excel файл — в папке `attached_assets/` (не обязателен)

### Шаг 2: Установка зависимостей

```bash
cd your-project
pip install -r requirements.txt
```

Или с использованием `uv` (быстрее):
```bash
uv pip install -r requirements.txt
```

### Шаг 3: Настройка API ключа

**Вариант A: OpenAI API (платный)**

```bash
export OPENAI_API_KEY=sk-...
```

Затем в `ai_assistant.py` замените:
```python
# Строка ~20:
client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY")
)
```

**Вариант B: Другой LLM (Claude, Gemini и т.д.)**

Отредактируйте `ai_assistant.py` для вашего провайдера.

**Вариант C: Без интернета (mock)**

Для тестирования UI без реального ИИ, верните фиксированный ответ в функции `ask_ai()`.

### Запуск

```bash
streamlit run app.py
```

Приложение откроется на **http://localhost:8501**

---

## Содержимое ключевых файлов

### app.py (~320 строк)
```python
import streamlit as st
import pandas as pd
from data_processor import load_excel, get_schema_summary, get_sheet_prefix_info
from ai_assistant import ask_ai

# Основные компоненты:
# 1. Инициализация сессии (история, загруженные данные)
# 2. Боковая панель (загрузка файла, инфо о листах)
# 3. Три вкладки (Чат, Данные, Примеры)
# 4. История диалога с кешированием результатов
```

### data_processor.py (~130 строк)
```python
import pandas as pd
import openpyxl

def load_excel(file_source):
    # Загружает Excel, парсит листы, типизирует колонки
    # Возвращает dict: {"Лист": DataFrame}

def get_schema_summary(sheets):
    # Описание схемы для ИИ-промпта

def get_sheet_prefix_info(sheets):
    # Информация о филиалах (префиксы номеров)
```

### ai_assistant.py (~190 строк)
```python
from openai import OpenAI

def ask_ai(question, schema_summary, prefix_info, history, sheets):
    # Главная функция для общения с ИИ
    # 1. Строит промпт (система + контекст + история + вопрос)
    # 2. Отправляет в OpenAI
    # 3. Парсит JSON ответ (code + explanation)
    # 4. Выполняет код в безопасном namespace
    # 5. Возвращает результаты (text, df, fig)

def _execute_code(code, sheets):
    # Безопасное выполнение сгенерированного кода
    # Доступны только: pd, plt, DataFrames, встроенные функции

def fig_to_bytes(fig):
    # matplotlib.Figure → PNG bytes

def df_to_csv_bytes(df):
def df_to_excel_bytes(df):
    # DataFrame → CSV/XLSX bytes
```

### .streamlit/config.toml
```ini
[server]
headless = true
address = 0.0.0.0
port = 8501
```

### requirements.txt
```
streamlit>=1.55.0
pandas>=2.0.0
openpyxl>=3.10.0
matplotlib>=3.7.0
plotly>=5.0.0
kaleido>=0.2.0
fpdf2>=2.7.0
xlsxwriter>=3.0.0
openai>=1.0.0
tenacity>=8.0.0
```

---

## Проверка установки

```bash
# Проверить Python
python --version  # должна быть 3.11+

# Проверить зависимости
pip list | grep streamlit
pip list | grep pandas
pip list | grep openai

# Тест Streamlit
streamlit hello  # должно открыть пример приложения

# Тест вашего приложения
streamlit run app.py
```

---

## Типичные проблемы и решения

| Проблема | Решение |
|----------|---------|
| `ModuleNotFoundError: No module named 'streamlit'` | `pip install -r requirements.txt` |
| `Python 3.10 или ниже` | Установите Python 3.11+: `python3.11` |
| `Port 8501 already in use` | `streamlit run app.py --logger.level=debug --server.port 8502` |
| `OpenAI API Error 401` | Проверьте `OPENAI_API_KEY` |
| `openpyxl error` | `pip install --upgrade openpyxl` |
| Ошибка matplotlib графиков | Обновите: `pip install --upgrade matplotlib` |

---

## Для Windows

```batch
# CMD
set OPENAI_API_KEY=sk-...
python -m streamlit run app.py

REM Или используйте PowerShell:
$env:OPENAI_API_KEY="sk-..."
python -m streamlit run app.py
```

---

## Запуск в фоне (Linux/Mac)

```bash
nohup streamlit run app.py > app.log 2>&1 &
# или
tmux new-session -d 'streamlit run app.py'
```

Остановка:
```bash
pkill -f streamlit
```

---

## Дополнительно: Виртуальное окружение (рекомендуется)

```bash
# Создать
python -m venv venv

# Активировать
source venv/bin/activate  # Linux/Mac
# или
venv\Scripts\activate     # Windows

# Установить зависимости
pip install -r requirements.txt

# Запустить
streamlit run app.py
```

---

## Дальнейшее развитие

После локального старта вы можете:

1. **Модифицировать промпт** — отредактировать `SYSTEM_PROMPT` в `ai_assistant.py` для вашего домена
2. **Добавить БД** — сохранять историю анализов (SQLite, PostgreSQL)
3. **Расширить функции** — custom экспорт, планирование, уведомления
4. **Развернуть в облако** — Streamlit Cloud, Heroku, Docker
5. **Оптимизировать** — кэширование, параллелизм, логирование

---

## Готово! 🎉

Если всё установилось без ошибок, приложение полностью функционально. Начните с примеров вопросов в приложении.

Вопросы? Смотрите `ARCHITECTURE.md` для подробной документации.
