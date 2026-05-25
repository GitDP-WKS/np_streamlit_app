import io
import re
from typing import Iterable

import pandas as pd
import streamlit as st


APP_TITLE = "Парсинг населенных пунктов"

PRIMARY_COLUMNS = {
    "filial_new": ["Филиал новый", "Новый филиал", "Филиал_новый"],
    "res_new": ["РЭС новый", "РЭС новый ", "Новый РЭС", "РЭС_новый", "РЕС новый"],
    "district": ["Район", "Муниципальный район"],
    "np": ["Населенный пункт", "Населенный  пункт", "НП", "Наименование населенного пункта"],
    "status": ["Статус населенного пункта", "Статус НП", "Тип населенного пункта"],
}

FIELD_LABELS = {
    "filial_new": "Филиал новый",
    "res_new": "РЭС новый",
    "district": "Район",
    "np": "Населенный пункт для парсинга",
    "status": "Статус населенного пункта",
}

OLD_COLUMNS = ["Филиал", "РЭС"]
DEFAULT_PARSE_COLUMN = "НП очищенный"


def normalize_column_name(value: object) -> str:
    text = str(value or "").replace("ё", "е").strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def find_column(columns: Iterable[str], variants: list[str]) -> str | None:
    normalized = {normalize_column_name(col): col for col in columns}
    for variant in variants:
        found = normalized.get(normalize_column_name(variant))
        if found:
            return found
    return None


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


def build_auto_column_map(df: pd.DataFrame) -> dict[str, str | None]:
    return {key: find_column(df.columns, variants) for key, variants in PRIMARY_COLUMNS.items()}


def column_selectbox(label: str, columns: list[str], suggested: str | None, key: str, required: bool = True) -> str | None:
    options = columns if required else ["Не использовать"] + columns
    index = options.index(suggested) if suggested in options else 0

    value = st.selectbox(label, options, index=index, key=key)
    if value == "Не использовать":
        return None
    return value


def choose_columns(df: pd.DataFrame) -> dict[str, str | None]:
    columns = list(df.columns)
    auto_map = build_auto_column_map(df)

    st.subheader("Настройка столбцов")
    st.caption("Приложение предложит найденные столбцы автоматически, но их можно изменить вручную.")

    with st.expander("Автоматически найденные столбцы", expanded=False):
        st.dataframe(
            pd.DataFrame(
                [{"Поле": FIELD_LABELS[key], "Найдено": value or "не найдено"} for key, value in auto_map.items()]
            ),
            use_container_width=True,
        )

    left, right = st.columns(2)
    with left:
        filial_new = column_selectbox(
            "Столбец с новым филиалом",
            columns,
            auto_map.get("filial_new"),
            "select_filial_new",
            required=False,
        )
        district = column_selectbox(
            "Столбец с районом",
            columns,
            auto_map.get("district"),
            "select_district",
            required=False,
        )
    with right:
        res_new = column_selectbox(
            "Столбец с новым РЭС",
            columns,
            auto_map.get("res_new"),
            "select_res_new",
            required=False,
        )
        np_col = column_selectbox(
            "Столбец, который нужно парсить",
            columns,
            auto_map.get("np"),
            "select_np_col",
            required=True,
        )

    status_col = column_selectbox(
        "Столбец со статусом НП, если есть",
        columns,
        auto_map.get("status"),
        "select_status_col",
        required=False,
    )

    return {
        "filial_new": filial_new,
        "res_new": res_new,
        "district": district,
        "np": np_col,
        "status": status_col,
    }


def choose_parse_output(df: pd.DataFrame) -> tuple[str, bool]:
    st.subheader("Куда записать результат парсинга")

    mode = st.radio(
        "Выберите способ записи результата",
        ["Создать новую колонку", "Записать в существующую колонку"],
        horizontal=True,
    )

    if mode == "Создать новую колонку":
        parse_column = st.text_input("Название новой колонки", value=DEFAULT_PARSE_COLUMN)
        overwrite_existing = False
    else:
        parse_column = st.selectbox("Колонка для записи результата", list(df.columns))
        overwrite_existing = True

    parse_column = compact_text(parse_column)
    if not parse_column:
        st.error("Укажите название колонки для результата парсинга.")
        st.stop()

    if parse_column in df.columns and not overwrite_existing:
        st.warning(
            f"Колонка '{parse_column}' уже есть в файле. Результат будет записан в нее, чтобы не создавать дубль."
        )
        overwrite_existing = True

    return parse_column, overwrite_existing


def validate_required_columns(col_map: dict[str, str | None]) -> list[str]:
    if not col_map.get("np"):
        return [FIELD_LABELS["np"]]
    return []


def prepare_dataframe(df: pd.DataFrame, col_map: dict[str, str | None], parse_column: str) -> pd.DataFrame:
    result = df.copy()
    np_col = col_map["np"]
    result[parse_column] = result[np_col].apply(clean_np)

    drop_candidates = [
        col for col in OLD_COLUMNS
        if col in result.columns and col not in {col_map.get("filial_new"), col_map.get("res_new")}
    ]
    result = result.drop(columns=drop_candidates, errors="ignore")

    return result


def build_summary(df: pd.DataFrame, col_map: dict[str, str | None], parse_column: str) -> pd.DataFrame:
    parsed_empty = int((df[parse_column].astype(str).str.strip() == "").sum())

    return pd.DataFrame(
        [
            {"Показатель": "Всего строк", "Значение": len(df)},
            {"Показатель": "Пустых результатов парсинга", "Значение": parsed_empty},
            {"Показатель": "Колонка для парсинга", "Значение": col_map["np"]},
            {"Показатель": "Колонка результата парсинга", "Значение": parse_column},
            {"Показатель": "Колонка филиала", "Значение": col_map.get("filial_new") or "не используется"},
            {"Показатель": "Колонка РЭС", "Значение": col_map.get("res_new") or "не используется"},
            {"Показатель": "Колонка района", "Значение": col_map.get("district") or "не используется"},
            {"Показатель": "Колонка статуса НП", "Значение": col_map.get("status") or "не используется"},
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
    "Приложение очищает выбранный столбец с населенными пунктами и записывает результат в выбранную колонку. "
    "Анализ дублей полностью удален."
)

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

col_map = choose_columns(source_df)
missing_columns = validate_required_columns(col_map)

if missing_columns:
    st.error("Не выбраны обязательные столбцы: " + ", ".join(missing_columns))
    st.stop()

parse_column, overwrite_existing = choose_parse_output(source_df)
run_analysis = st.button("Запустить парсинг", type="primary")

if not run_analysis:
    st.info("Выберите столбец для парсинга, укажите колонку результата и нажмите кнопку запуска.")
    st.stop()

parsed_df = prepare_dataframe(source_df, col_map, parse_column)
summary_df = build_summary(parsed_df, col_map, parse_column)

left, right = st.columns(2)
left.metric("Всего строк", len(parsed_df))
right.metric("Пустых результатов", int(summary_df.loc[summary_df["Показатель"] == "Пустых результатов парсинга", "Значение"].iloc[0]))

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
