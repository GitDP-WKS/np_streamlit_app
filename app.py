import io
import re

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process


APP_TITLE = "Умный парсинг населенных пунктов"
DEFAULT_RESULT_COLUMN = "НП из справочника"
DEFAULT_SCORE_COLUMN = "Точность НП"
DEFAULT_STATUS_COLUMN = "Статус парсинга НП"
MATCH_THRESHOLD = 82

REFERENCE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1UhrxXABHp5yrtiAm7excLPicHhlgZrgN7-8R3Ada9ZY/export?format=xlsx&gid=0"
)
REFERENCE_COLUMN_CANDIDATES = [
    "сокр.Населенный пункт",
    "Населенный пункт",
    "НП",
    "Наименование населенного пункта",
]


def compact_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("ё", "е").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_for_match(value: object) -> str:
    text = compact_text(value).lower()
    text = text.replace("«", " ").replace("»", " ")
    text = re.sub(r"[\"'`.,;:()№/\\\-]+", " ", text)
    text = re.sub(r"\bн\s*п\b", "нп", text)
    text = re.sub(r"\bп\s*г\s*т\b", "пгт", text)
    text = re.sub(r"\bж\s*д\s*ст\b", "жд ст", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_name_only(value: object) -> str:
    text = normalize_for_match(value)
    text = re.sub(r"^(г|с|д|п|нп|пгт|жд ст|снт)\s+", "", text)
    return text.strip()


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


@st.cache_data(show_spinner=False)
def load_reference() -> pd.DataFrame:
    return pd.read_excel(REFERENCE_URL)


def find_reference_column(ref_df: pd.DataFrame) -> str:
    normalized_columns = {compact_text(col).lower(): col for col in ref_df.columns}
    for candidate in REFERENCE_COLUMN_CANDIDATES:
        found = normalized_columns.get(compact_text(candidate).lower())
        if found:
            return found

    text_columns = [col for col in ref_df.columns if ref_df[col].dtype == object]
    if text_columns:
        return text_columns[0]

    return ref_df.columns[0]


def build_reference_dict(ref_df: pd.DataFrame, ref_column: str) -> list[dict[str, str]]:
    values = ref_df[ref_column].dropna().map(compact_text)
    values = values[values != ""].drop_duplicates().tolist()

    reference = []
    for value in values:
        reference.append(
            {
                "original": value,
                "clean": clean_np(value),
                "full_key": normalize_for_match(value),
                "name_key": normalize_name_only(value),
            }
        )

    return reference


def match_reference_cell(raw_value: object, reference: list[dict[str, str]]) -> dict[str, object]:
    raw_text = compact_text(raw_value)
    if not raw_text:
        return {"value": "", "score": 0, "status": "пусто"}

    cell_key = normalize_for_match(raw_text)
    cell_name_key = normalize_name_only(raw_text)

    if not cell_key:
        return {"value": "", "score": 0, "status": "пусто"}

    for item in reference:
        full_key = item["full_key"]
        name_key = item["name_key"]

        if full_key and re.search(rf"\b{re.escape(full_key)}\b", cell_key):
            return {"value": item["clean"], "score": 100, "status": "найдено точно"}

        if name_key and re.search(rf"\b{re.escape(name_key)}\b", cell_name_key):
            return {"value": item["clean"], "score": 98, "status": "найдено по названию"}

    choices = {item["full_key"]: item for item in reference if item["full_key"]}
    best = process.extractOne(cell_key, list(choices.keys()), scorer=fuzz.WRatio)

    if best:
        best_key, score, _ = best
        if score >= MATCH_THRESHOLD:
            item = choices[best_key]
            status = "найдено похоже" if score >= 90 else "требует проверки"
            return {"value": item["clean"], "score": round(score, 1), "status": status}

    name_choices = {item["name_key"]: item for item in reference if item["name_key"]}
    best_name = process.extractOne(cell_name_key, list(name_choices.keys()), scorer=fuzz.WRatio)

    if best_name:
        best_key, score, _ = best_name
        if score >= MATCH_THRESHOLD:
            item = name_choices[best_key]
            status = "найдено похоже" if score >= 90 else "требует проверки"
            return {"value": item["clean"], "score": round(score, 1), "status": status}

    return {"value": "", "score": 0, "status": "не найдено"}


def choose_columns(df: pd.DataFrame) -> tuple[str, str]:
    columns = list(df.columns)

    st.subheader("Выбор столбцов")
    left, right = st.columns(2)

    with left:
        source_column = st.selectbox(
            "Где находится текст для анализа",
            columns,
            key="source_column",
        )

    with right:
        output_mode = st.radio(
            "Куда записать найденный НП",
            ["Создать новую колонку", "Записать в существующую колонку"],
            horizontal=True,
        )

        if output_mode == "Создать новую колонку":
            result_column = st.text_input("Название колонки результата", value=DEFAULT_RESULT_COLUMN)
        else:
            result_column = st.selectbox("Столбец для записи результата", columns, key="result_column")

    result_column = compact_text(result_column)
    if not result_column:
        st.error("Укажите колонку для результата парсинга.")
        st.stop()

    return source_column, result_column


def parse_dataframe(
    df: pd.DataFrame,
    source_column: str,
    result_column: str,
    reference: list[dict[str, str]],
) -> pd.DataFrame:
    result = df.copy()
    matches = result[source_column].apply(lambda value: match_reference_cell(value, reference))
    match_df = pd.DataFrame(list(matches))

    result[result_column] = match_df["value"]
    result[DEFAULT_SCORE_COLUMN] = match_df["score"]
    result[DEFAULT_STATUS_COLUMN] = match_df["status"]
    return result


def build_summary(
    df: pd.DataFrame,
    source_column: str,
    result_column: str,
    ref_column: str,
    reference_count: int,
) -> pd.DataFrame:
    empty_result_count = int((df[result_column].astype(str).str.strip() == "").sum())
    found_count = len(df) - empty_result_count
    check_count = int((df[DEFAULT_STATUS_COLUMN] == "требует проверки").sum())

    return pd.DataFrame(
        [
            {"Показатель": "Всего строк", "Значение": len(df)},
            {"Показатель": "Найдено НП", "Значение": found_count},
            {"Показатель": "Не найдено НП", "Значение": empty_result_count},
            {"Показатель": "Требует проверки", "Значение": check_count},
            {"Показатель": "Исходный столбец", "Значение": source_column},
            {"Показатель": "Столбец результата", "Значение": result_column},
            {"Показатель": "Столбец справочника", "Значение": ref_column},
            {"Показатель": "Значений в справочнике", "Значение": reference_count},
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

st.caption(
    "Приложение анализирует выбранную ячейку и вытаскивает из нее только тот населенный пункт, "
    "который есть в эталонном Google справочнике."
)

try:
    reference_df = load_reference()
except Exception as error:
    st.error(f"Не удалось загрузить Google справочник: {error}")
    st.stop()

ref_column = find_reference_column(reference_df)
reference = build_reference_dict(reference_df, ref_column)

with st.expander("Справочник Google", expanded=False):
    st.write(f"Используется столбец справочника: {ref_column}")
    st.write(f"Уникальных значений: {len(reference)}")
    st.dataframe(reference_df.head(20), use_container_width=True)

uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("Загрузите файл, чтобы начать анализ.")
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
run_parsing = st.button("Запустить анализ", type="primary")

if not run_parsing:
    st.info("Выберите исходный столбец и столбец результата, затем нажмите кнопку запуска.")
    st.stop()

parsed_df = parse_dataframe(source_df, source_column, result_column, reference)
summary_df = build_summary(parsed_df, source_column, result_column, ref_column, len(reference))

left, middle, right = st.columns(3)
left.metric("Всего строк", len(parsed_df))
middle.metric(
    "Найдено НП",
    int(summary_df.loc[summary_df["Показатель"] == "Найдено НП", "Значение"].iloc[0]),
)
right.metric(
    "Не найдено",
    int(summary_df.loc[summary_df["Показатель"] == "Не найдено НП", "Значение"].iloc[0]),
)

st.subheader("Сводка")
st.dataframe(summary_df, use_container_width=True)

st.subheader("Исправленный файл")
st.dataframe(parsed_df, use_container_width=True, height=480)

excel_bytes = make_excel(parsed_df, summary_df)

st.download_button(
    "Скачать результат Excel",
    data=excel_bytes,
    file_name="np_ai_parse_result.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
