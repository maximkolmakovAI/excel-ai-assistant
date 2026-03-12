import streamlit as st
import pandas as pd
import io
import os
import json
import hashlib
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

from data_processor import load_excel, get_schema_summary, get_sheet_prefix_info
from ai_assistant import ask_ai, fig_to_bytes, df_to_excel_bytes, df_to_csv_bytes

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ИИ-ассистент по Excel",
    page_icon="📊",
    layout="wide",
)

SAMPLE_FILE = "attached_assets/Сентябрь 2020.xlsx"

# ── Session state defaults ─────────────────────────────────────────────────────
def init_state():
    defaults = {
        "sheets": {},
        "schema_summary": "",
        "prefix_info": "",
        "chat_history": [],
        "file_loaded": False,
        "file_name": "",
        "show_code": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ── Helpers ────────────────────────────────────────────────────────────────────
def load_sheets(source, name: str):
    with st.spinner("Загружаю и обрабатываю файл…"):
        sheets = load_excel(source)
        schema = get_schema_summary(sheets)
        prefix = get_sheet_prefix_info(sheets)
    st.session_state.sheets = sheets
    st.session_state.schema_summary = schema
    st.session_state.prefix_info = prefix
    st.session_state.file_loaded = True
    st.session_state.file_name = name
    st.session_state.chat_history = []


def format_number(n):
    try:
        return f"{float(n):,.0f}".replace(",", " ")
    except Exception:
        return str(n)


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📂 Данные")

    # Load sample file automatically on first run
    if not st.session_state.file_loaded and os.path.exists(SAMPLE_FILE):
        load_sheets(SAMPLE_FILE, "Сентябрь_2020.xlsx")

    if st.session_state.file_loaded:
        st.success(f"**Файл:** {st.session_state.file_name}")
        for name, df in st.session_state.sheets.items():
            st.metric(f"Лист «{name}»", f"{len(df)} строк")

    st.divider()
    # Скрыто для тестового режима
    # st.subheader("Загрузить другой файл")
    # uploaded = st.file_uploader(
    #     "Excel (.xlsx, .xls)",
    #     type=["xlsx", "xls"],
    #     label_visibility="collapsed",
    # )
    # if uploaded is not None:
    #     file_key = hashlib.md5(uploaded.getvalue()).hexdigest()
    #     if st.session_state.get("_last_upload_key") != file_key:
    #         st.session_state["_last_upload_key"] = file_key
    #         load_sheets(io.BytesIO(uploaded.getvalue()), uploaded.name)
    #         st.rerun()

    st.divider()
    st.subheader("⚙️ Настройки")
    st.session_state.show_code = st.toggle("Показывать сгенерированный код", value=st.session_state.show_code)

    if st.session_state.chat_history:
        if st.button("🗑️ Очистить историю чата", use_container_width=True):
            st.session_state.chat_history = []
            st.rerun()

    st.divider()


# ── Main area ──────────────────────────────────────────────────────────────────
st.title("📊 ИИ-ассистент по Excel")

if not st.session_state.file_loaded:
    st.info("Загрузите Excel-файл через боковую панель слева.")
    st.stop()

# Tabs
tab_chat, tab_data, tab_examples = st.tabs(["💬 Чат-аналитик", "🗂️ Данные", "💡 Примеры вопросов"])

# ── TAB: Данные ────────────────────────────────────────────────────────────────
with tab_data:
    st.subheader("Просмотр загруженных таблиц")
    if st.session_state.sheets:
        sheet_names = list(st.session_state.sheets.keys())
        selected_sheet = st.selectbox("Выберите лист", sheet_names)
        df_view = st.session_state.sheets[selected_sheet]
        st.caption(f"Всего строк: {len(df_view)}, колонок: {len(df_view.columns)}")

        col_search, col_rows = st.columns([3, 1])
        with col_search:
            search_term = st.text_input("🔍 Поиск по тексту", placeholder="Введите текст для фильтрации…", label_visibility="collapsed")
        with col_rows:
            max_rows = st.selectbox("Строк на странице", [50, 100, 200, 500], label_visibility="collapsed")

        if search_term:
            mask = df_view.apply(lambda col: col.astype(str).str.contains(search_term, case=False, na=False)).any(axis=1)
            df_filtered = df_view[mask]
        else:
            df_filtered = df_view

        st.dataframe(df_filtered.head(max_rows), use_container_width=True, height=450)

        st.download_button(
            "⬇️ Скачать лист как CSV",
            data=df_to_csv_bytes(df_filtered),
            file_name=f"{selected_sheet}.csv",
            mime="text/csv",
        )


# ── TAB: Примеры вопросов ──────────────────────────────────────────────────────
EXAMPLE_QUESTIONS = [
    "Найти заявки с суммой от 200 тыс. до 1 млн руб.",
    "Показать заявки на сумму от 500 тыс. руб., по которым были получены заказы (Заказано=100).",
    "Вывести заявки менеджера Корминкина на сумму от 1 млн руб.",
    "Определить менеджеров московского филиала, у которых количество заявок в первой половине периода было больше, чем во второй.",
    "Определить менеджера, который получил заявки от наибольшего количества предприятий; вывести список этих предприятий с количеством заявок от каждого.",
    "Построить круговую диаграмму по количеству заявок, полученных подразделениями компании — отделами центрального офиса и филиалами.",
    "Построить гистограмму полученных компанией заявок за период по дням или неделям.",
    "Построить гистограмму с накоплением по количеству заявок, полученных менеджерами московского филиала.",
    "Построить диаграмму по количеству заявок с сегментами — регионами по местонахождению предприятий.",
    "Какова общая сумма заявок по каждому менеджеру? Показать топ-10.",
]

with tab_examples:
    st.subheader("Примеры вопросов для ассистента")
    st.caption("Нажмите на вопрос, чтобы отправить его в чат.")
    for i, q in enumerate(EXAMPLE_QUESTIONS):
        if st.button(f"❓ {q}", key=f"example_{i}", use_container_width=True):
            st.session_state["_pending_question"] = q
            st.rerun()


# ── TAB: Чат ──────────────────────────────────────────────────────────────────
with tab_chat:
    # Display chat history
    chat_container = st.container()

    with chat_container:
        if not st.session_state.chat_history:
            st.info(
                "Здравствуйте! Я анализирую ваши Excel-данные. Задайте вопрос о заявках, заказах или предприятиях.\n\n"
                "Посмотрите вкладку **💡 Примеры вопросов** для идей."
            )

        for idx, turn in enumerate(st.session_state.chat_history):
            with st.chat_message("user"):
                st.write(turn["user"])

            with st.chat_message("assistant"):
                st.write(turn["text"])

                if turn.get("df") is not None:
                    df_result = turn["df"]
                    st.caption(f"Таблица результатов ({len(df_result)} строк)")
                    st.dataframe(df_result, use_container_width=True)

                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button(
                            "⬇️ CSV",
                            data=df_to_csv_bytes(df_result),
                            file_name=f"result_{idx}.csv",
                            mime="text/csv",
                            key=f"dl_csv_{idx}",
                        )
                    with c2:
                        st.download_button(
                            "⬇️ XLSX",
                            data=df_to_excel_bytes(df_result),
                            file_name=f"result_{idx}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_xlsx_{idx}",
                        )

                if turn.get("fig_bytes") is not None:
                    st.image(turn["fig_bytes"])
                    st.download_button(
                        "⬇️ Скачать график (PNG)",
                        data=turn["fig_bytes"],
                        file_name=f"chart_{idx}.png",
                        mime="image/png",
                        key=f"dl_png_{idx}",
                    )

                if st.session_state.show_code and turn.get("code"):
                    with st.expander("🔍 Сгенерированный код"):
                        st.code(turn["code"], language="python")

                if turn.get("error"):
                    with st.expander("⚠️ Детали ошибки"):
                        st.code(turn["error"])

    # Input area
    pending = st.session_state.pop("_pending_question", None)

    user_input = st.chat_input("Задайте вопрос о ваших данных…")
    question = pending or user_input

    if question:
        with st.chat_message("user"):
            st.write(question)

        with st.chat_message("assistant"):
            with st.spinner("Анализирую данные…"):
                result = ask_ai(
                    question=question,
                    schema_summary=st.session_state.schema_summary,
                    prefix_info=st.session_state.prefix_info,
                    history=st.session_state.chat_history,
                    sheets=st.session_state.sheets,
                )

            st.write(result["text"])

            fig_bytes = None
            if result.get("fig") is not None:
                try:
                    fig_bytes = fig_to_bytes(result["fig"])
                    st.image(fig_bytes)
                    st.download_button(
                        "⬇️ Скачать график (PNG)",
                        data=fig_bytes,
                        file_name=f"chart_{len(st.session_state.chat_history)}.png",
                        mime="image/png",
                        key="dl_png_new",
                    )
                except Exception as e:
                    st.warning(f"Не удалось отобразить график: {e}")

            df_result = result.get("df")
            if df_result is not None and not (isinstance(df_result, pd.DataFrame) and df_result.empty):
                st.caption(f"Таблица результатов ({len(df_result)} строк)")
                st.dataframe(df_result, use_container_width=True)

                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        "⬇️ CSV",
                        data=df_to_csv_bytes(df_result),
                        file_name=f"result_{len(st.session_state.chat_history)}.csv",
                        mime="text/csv",
                        key="dl_csv_new",
                    )
                with c2:
                    st.download_button(
                        "⬇️ XLSX",
                        data=df_to_excel_bytes(df_result),
                        file_name=f"result_{len(st.session_state.chat_history)}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_xlsx_new",
                    )

            if st.session_state.show_code and result.get("code"):
                with st.expander("🔍 Сгенерированный код"):
                    st.code(result["code"], language="python")

            if result.get("error"):
                with st.expander("⚠️ Детали ошибки"):
                    st.code(result["error"])

        # Save to history
        st.session_state.chat_history.append({
            "user": question,
            "text": result["text"],
            "df": df_result if df_result is not None and not (isinstance(df_result, pd.DataFrame) and df_result.empty) else None,
            "fig_bytes": fig_bytes,
            "code": result.get("code", ""),
            "error": result.get("error"),
            "assistant_raw": result.get("explanation", result["text"]),
        })

        if result.get("fig") is not None:
            try:
                plt.close(result["fig"])
            except Exception:
                pass
