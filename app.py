import io
import re

import pandas as pd
import streamlit as st


APP_TITLE = "Парсинг населенных пунктов"
DEFAULT_PARSE_COLUMN = "НП очищенный"


def compact_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("ё", "е").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def format_snt_quotes(value: object) -> str:
    text = compact_text(value)
    if not text:
        return ""

    text = text.replace("“", "«").replace("”", "»").replace('"', "«")
    text = re.sub(r"\s+", " ", text).strip()

    match = re.match(r"(?i)^снт\s+[«\"]?(.+?)[»\"]?$", text)
    if match:
        name = match.group(1).strip(" «»\"")
        return f"СНТ «{name}»"

    return text


def clean_np(value: object) -> str:
    text = format_snt_quotes(value)
    replacements = [
        (r"(?i)^н\s*\.?\s*п\s*\.?\s+", "н.п. "),
        (r"(?i)^п\s*\.?\s*г\s*\.?\s*т\s*\.?\s+", "пгт "),
        (r"(?i)^ж\s*/?\s*д\s*\.?\s*ст\s*\.?\s+", "ж/д ст "),
        (r"(?i)^г\s*\.?\s+", "г. "),
        (r"(?i)^с\s*\.?\s+", "с. "),
        (r"(?i)^д\s*\.?\s+", "д. "),
        (r"(?i)^п\s*\.?\s+", "п. "),
    ]
    for pattern, repl in replacements:
        text = re.sub(pattern, repl, text)
    return re.sub(r"\s+", " ", text).strip()


def read_excel_sheets(uploaded_file) -> dict[str, pd.DataFrame]:
    return pd.read_excel(uploaded_file, sheet_name=None)


def choose_columns(df: pd.DataFrame) -> tuple[str, str]:
    columns = list(df.columns)

    st.subheader("Выбор столбцов")
    left, right = st.columns(2)

    with left:
        source_column = st.selectbox(
            "Где находится название НП",
            columns,
            key="source_column",
        )

    with right:
        output_mode = st.radio(
            "Куда записать результат",
            ["Создать новую колонку", "Записать в существующую колонку"],
            horizontal=True,
        )

        if output_mode == "Создать новую колонку":
            result_column = st.text_input("Название колонки результата", value=DEFAULT_PARSE_COLUMN)
        else:
            result_column = st.selectbox("Столбец для записи результата", columns, key="result_column")

    result_column = compact_text(result_column)
    if not result_column:
        st.error("Укажите колонку для результата парсинга.")
        st.stop()

    return source_column, result_column


def parse_dataframe(df: pd.DataFrame, source_column: str, result_column: str) -> pd.DataFrame:
    result = df.copy()
    result[result_column] = result[source_column].apply(clean_np)
    return result


def build_summary(df: pd.DataFrame, source_column: str, result_column: str) -> pd.DataFrame:
    empty_result_count = int((df[result_column].astype(str).str.strip() == "").sum())

    return pd.DataFrame(
        [
            {"Показатель": "Всего строк", "Значение": len(df)},
            {"Показатель": "Пустых результатов парсинга", "Значение": empty_result_count},
            {"Показатель": "Исходный столбец", "Значение": source_column},
            {"Показатель": "Столбец результата", "Значение": result_column},
        ]
    )


def make_excel(parsed_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        parsed_df.to_excel(writer, index=False, sheet_name="Исправленный файл")
        summary_df.to_excel(writer, index=False, sheet_name="Сводка")

        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})

        for sheet_name, sheet_df in {
            "Исправленный файл": parsed_df,
            "Сводка": summary_df,
        }.items():
            worksheet = writer.sheets[sheet_name]
            for col_num, value in enumerate(sheet_df.columns):
                worksheet.write(0, col_num, value, header_format)
                width = min(max(len(str(value)) + 4, 14), 42)
                worksheet.set_column(col_num, col_num, width)
            worksheet.freeze_panes(1, 0)

    return output.getvalue()


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.caption("Загрузите Excel файл, выберите исходный столбец с НП и столбец для результата парсинга.")

uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("Загрузите файл, чтобы начать парсинг.")
    st.stop()

try:
    sheets = read_excel_sheets(uploaded_file)
except Exception as error:
    st.error(f"Не удалось прочитать Excel файл: {error}")
    st.stop()

sheet_name = st.selectbox("Выберите лист для обработки", list(sheets.keys()))
source_df = sheets[sheet_name].copy()

if source_df.empty:
    st.warning("Выбранный лист пустой.")
    st.stop()

st.subheader("Предпросмотр файла")
st.dataframe(source_df.head(30), use_container_width=True)

source_column, result_column = choose_columns(source_df)
run_parsing = st.button("Запустить парсинг", type="primary")

if not run_parsing:
    st.info("Выберите два столбца и нажмите кнопку запуска.")
    st.stop()

parsed_df = parse_dataframe(source_df, source_column, result_column)
summary_df = build_summary(parsed_df, source_column, result_column)

left, right = st.columns(2)
left.metric("Всего строк", len(parsed_df))
right.metric(
    "Пустых результатов",
    int(summary_df.loc[summary_df["Показатель"] == "Пустых результатов парсинга", "Значение"].iloc[0]),
)

st.subheader("Сводка")
st.dataframe(summary_df, use_container_width=True)

st.subheader("Исправленный файл")
st.dataframe(parsed_df, use_container_width=True, height=480)

excel_bytes = make_excel(parsed_df, summary_df)

st.download_button(
    "Скачать результат Excel",
    data=excel_bytes,
    file_name="np_parse_result.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
