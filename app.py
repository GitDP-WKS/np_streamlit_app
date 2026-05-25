import io
import re
from typing import Iterable

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz


APP_TITLE = "Проверка населенных пунктов"

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
FUZZY_LIMIT = 92
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


def normalize_np_key(value: object) -> str:
    text = compact_text(value).lower()
    text = text.replace("«", '"').replace("»", '"')
    text = re.sub(r"[\"'`.,;:()№]+", " ", text)
    text = re.sub(r"\bн\s*\.?\s*п\s*\.?\b", "нп", text)
    text = re.sub(r"\bп\s*\.?\s*г\s*\.?\s*т\s*\.?\b", "пгт", text)
    text = re.sub(r"\bж\s*/?\s*д\s*\.?\s*ст\s*\.?\b", "жд ст", text)
    text = re.sub(r"\b(г|с|д|п)\s*\.\s*", r"\1 ", text)
    text = re.sub(r"\s+", " ", text)
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


def build_auto_column_map(df: pd.DataFrame) -> dict[str, str | None]:
    return {key: find_column(df.columns, variants) for key, variants in PRIMARY_COLUMNS.items()}


def column_selectbox(label: str, columns: list[str], suggested: str | None, key: str, required: bool = True) -> str | None:
    options = columns if required else ["Не использовать"] + columns
    if suggested in options:
        index = options.index(suggested)
    else:
        index = 0

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
        )
        district = column_selectbox(
            "Столбец с районом",
            columns,
            auto_map.get("district"),
            "select_district",
        )
    with right:
        res_new = column_selectbox(
            "Столбец с новым РЭС",
            columns,
            auto_map.get("res_new"),
            "select_res_new",
        )
        np_col = column_selectbox(
            "Столбец, который нужно парсить",
            columns,
            auto_map.get("np"),
            "select_np_col",
        )

    status_col = column_selectbox(
        "Столбец со статусом НП, если есть. В дублях не используется",
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
    missing = []
    for key in ["filial_new", "res_new", "district", "np"]:
        if not col_map.get(key):
            missing.append(FIELD_LABELS[key])
    return missing


def prepare_dataframe(df: pd.DataFrame, col_map: dict[str, str | None], parse_column: str) -> pd.DataFrame:
    result = df.copy()

    np_col = col_map["np"]
    result[parse_column] = result[np_col].apply(clean_np)
    result["_np_key"] = result[parse_column].apply(normalize_np_key)

    for key, output_col in [
        ("filial_new", "_filial_key"),
        ("res_new", "_res_key"),
        ("district", "_district_key"),
    ]:
        source_col = col_map[key]
        result[output_col] = result[source_col].apply(lambda x: compact_text(x).lower())

    duplicate_basis = ["_filial_key", "_res_key", "_district_key", "_np_key"]
    result["Группа дубля"] = result.groupby(duplicate_basis, dropna=False).ngroup() + 1
    result["Количество дублей"] = result.groupby(duplicate_basis, dropna=False)["_np_key"].transform("size")
    result["Статус проверки"] = result["Количество дублей"].apply(
        lambda count: "дубль" if count > 1 else "уникально"
    )

    drop_candidates = [
        col for col in OLD_COLUMNS
        if col in result.columns and col not in {col_map["filial_new"], col_map["res_new"]}
    ]
    result = result.drop(columns=drop_candidates, errors="ignore")

    return result


def find_suspicious_pairs(df: pd.DataFrame, col_map: dict[str, str | None], parse_column: str) -> pd.DataFrame:
    rows = []
    group_cols = ["_filial_key", "_res_key", "_district_key"]

    for _, group in df.groupby(group_cols, dropna=False):
        unique_names = (
            group[["_np_key", parse_column]]
            .dropna()
            .drop_duplicates()
            .query("_np_key != ''")
            .to_dict("records")
        )

        for i, left in enumerate(unique_names):
            for right in unique_names[i + 1:]:
                score = fuzz.WRatio(left["_np_key"], right["_np_key"])
                if score >= FUZZY_LIMIT and left["_np_key"] != right["_np_key"]:
                    rows.append(
                        {
                            "Филиал новый": group[col_map["filial_new"]].iloc[0],
                            "РЭС новый": group[col_map["res_new"]].iloc[0],
                            "Район": group[col_map["district"]].iloc[0],
                            "Вариант 1": left[parse_column],
                            "Вариант 2": right[parse_column],
                            "Похожесть": round(score, 1),
                            "Комментарий": "возможная разница в написании",
                        }
                    )

    return pd.DataFrame(rows)


def build_summary(
    df: pd.DataFrame,
    suspicious_df: pd.DataFrame,
    col_map: dict[str, str | None],
    parse_column: str,
) -> pd.DataFrame:
    rows_total = len(df)
    duplicate_rows = int((df["Количество дублей"] > 1).sum())
    duplicate_groups = int(df.loc[df["Количество дублей"] > 1, "Группа дубля"].nunique())
    empty_np = int((df["_np_key"] == "").sum())

    return pd.DataFrame(
        [
            {"Показатель": "Всего строк", "Значение": rows_total},
            {"Показатель": "Строк в дублях", "Значение": duplicate_rows},
            {"Показатель": "Групп дублей", "Значение": duplicate_groups},
            {"Показатель": "Пустых населенных пунктов", "Значение": empty_np},
            {"Показатель": "Подозрительных похожих пар", "Значение": len(suspicious_df)},
            {"Показатель": "Колонка филиала", "Значение": col_map["filial_new"]},
            {"Показатель": "Колонка РЭС", "Значение": col_map["res_new"]},
            {"Показатель": "Колонка района", "Значение": col_map["district"]},
            {"Показатель": "Колонка для парсинга", "Значение": col_map["np"]},
            {"Показатель": "Колонка результата парсинга", "Значение": parse_column},
            {"Показатель": "Колонка статуса НП", "Значение": col_map.get("status") or "не используется"},
        ]
    )


def public_columns(df: pd.DataFrame) -> pd.DataFrame:
    service_cols = [col for col in df.columns if col.startswith("_")]
    return df.drop(columns=service_cols, errors="ignore")


def make_excel(
    clean_df: pd.DataFrame,
    duplicates_df: pd.DataFrame,
    suspicious_df: pd.DataFrame,
    summary_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        public_columns(clean_df).to_excel(writer, index=False, sheet_name="Исправленный файл")
        public_columns(duplicates_df).to_excel(writer, index=False, sheet_name="Дубли")
        suspicious_df.to_excel(writer, index=False, sheet_name="Подозрительные")
        summary_df.to_excel(writer, index=False, sheet_name="Сводка")

        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})

        for sheet_name, sheet_df in {
            "Исправленный файл": public_columns(clean_df),
            "Дубли": public_columns(duplicates_df),
            "Подозрительные": suspicious_df,
            "Сводка": summary_df,
        }.items():
            worksheet = writer.sheets[sheet_name]
            for col_num, value in enumerate(sheet_df.columns):
                worksheet.write(0, col_num, value, header_format)
                width = min(max(len(str(value)) + 4, 14), 42)
                worksheet.set_column(col_num, col_num, width)
            worksheet.freeze_panes(1, 0)

        if not duplicates_df.empty:
            writer.sheets["Дубли"].set_tab_color("#70AD47")
        if not suspicious_df.empty:
            writer.sheets["Подозрительные"].set_tab_color("#FFC000")

    return output.getvalue()


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.caption(
    "Приложение проверяет файл по актуальной логике: новые колонки филиала и РЭС, район, населенный пункт. "
    "Статус населенного пункта не участвует в поиске дублей."
)

uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("Загрузите файл, чтобы начать проверку.")
    st.stop()

try:
    sheets = read_excel_sheets(uploaded_file)
except Exception as error:
    st.error(f"Не удалось прочитать Excel файл: {error}")
    st.stop()

sheet_name = st.selectbox("Выберите лист для проверки", list(sheets.keys()))
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

run_analysis = st.button("Запустить анализ", type="primary")

if not run_analysis:
    st.info("Выберите нужные столбцы, укажите колонку результата парсинга и нажмите кнопку анализа.")
    st.stop()

clean_df = prepare_dataframe(source_df, col_map, parse_column)
duplicates_df = clean_df[clean_df["Количество дублей"] > 1].sort_values(
    ["_filial_key", "_res_key", "_district_key", "_np_key"]
)
suspicious_df = find_suspicious_pairs(clean_df, col_map, parse_column)
summary_df = build_summary(clean_df, suspicious_df, col_map, parse_column)

left, middle, right = st.columns(3)
left.metric("Всего строк", len(clean_df))
middle.metric("Строк в дублях", len(duplicates_df))
right.metric("Подозрительных пар", len(suspicious_df))

st.subheader("Сводка")
st.dataframe(summary_df, use_container_width=True)

st.subheader("Исправленный файл")
st.dataframe(public_columns(clean_df), use_container_width=True, height=420)

st.subheader("Дубли")
if duplicates_df.empty:
    st.success("Полные дубли не найдены.")
else:
    st.dataframe(public_columns(duplicates_df), use_container_width=True, height=320)

st.subheader("Подозрительные похожие написания")
if suspicious_df.empty:
    st.success("Подозрительные похожие варианты не найдены.")
else:
    st.dataframe(suspicious_df, use_container_width=True, height=320)

excel_bytes = make_excel(clean_df, duplicates_df, suspicious_df, summary_df)

st.download_button(
    "Скачать результат Excel",
    data=excel_bytes,
    file_name="np_check_result.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
