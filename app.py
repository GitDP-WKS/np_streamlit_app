import io
import re

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process


APP_TITLE = "Умный парсинг населенных пунктов"
DEFAULT_RESULT_COLUMN = "НП из справочника"
DEFAULT_SCORE_COLUMN = "Точность НП"
DEFAULT_STATUS_COLUMN = "Статус парсинга НП"
MATCH_THRESHOLD = 94

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

DELETE_EXCEL_COLUMN_INDEXES = [41]  # AP в Excel, нумерация pandas с нуля
HIDE_EXCEL_RANGES = ["A:B", "E:U", "Y:AM"]
ADDRESS_MARKERS = [
    "ул", "улица", "пер", "переулок", "пр", "проспект", "д", "дом", "корп", "корпус",
    "кв", "квартира", "зд", "здание", "стр", "строение", "ш", "шоссе", "тракт",
    "пл", "площадь", "бульвар", "б р", "наб", "набережная"
]
SETTLEMENT_TYPE_TOKENS = {"г", "с", "д", "п", "нп", "пгт", "снт", "жд"}
BLOCKED_REFERENCE_KEYS = {
    "татарстан",
    "республика татарстан",
    "респ татарстан",
    "рт",
}


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
    text = re.sub(r"\bреспублика\s+татарстан\b", " ", text)
    text = re.sub(r"\bресп\s+татарстан\b", " ", text)
    text = re.sub(r"\bрт\b", " ", text)
    text = re.sub(r"\bн\s*п\b", "нп", text)
    text = re.sub(r"\bп\s*г\s*т\b", "пгт", text)
    text = re.sub(r"\bж\s*д\s*ст\b", "жд ст", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def has_address_marker(normalized_text: str) -> bool:
    return any(re.search(rf"\b{re.escape(marker)}\b", normalized_text) for marker in ADDRESS_MARKERS)


def token_count(value: str) -> int:
    return len([part for part in value.split() if part])


def starts_with_settlement_type(value: str) -> bool:
    parts = value.split()
    return bool(parts and parts[0] in SETTLEMENT_TYPE_TOKENS)


def is_city_or_bare_short_reference(full_key: str) -> bool:
    if not full_key:
        return False
    parts = full_key.split()
    if len(parts) == 1:
        return True
    if len(parts) == 2 and parts[0] == "г":
        return True
    return False


def is_blocked_reference(full_key: str) -> bool:
    if not full_key:
        return True
    if full_key in BLOCKED_REFERENCE_KEYS:
        return True
    if "татарстан" in full_key:
        return True
    return False


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
        full_key = normalize_for_match(value)
        if is_blocked_reference(full_key):
            continue
        reference.append(
            {
                "original": value,
                "clean": clean_np(value),
                "full_key": full_key,
                "is_city_or_bare_short": is_city_or_bare_short_reference(full_key),
            }
        )

    reference = sorted(reference, key=lambda item: len(item["full_key"]), reverse=True)
    return reference


def is_reference_allowed_for_cell(item: dict[str, str], cell_key: str, address_like: bool) -> bool:
    full_key = item["full_key"]
    if is_blocked_reference(full_key):
        return False

    if address_like and item["is_city_or_bare_short"] and cell_key != full_key:
        return False

    return True


def match_reference_cell(raw_value: object, reference: list[dict[str, str]]) -> dict[str, object]:
    raw_text = compact_text(raw_value)
    if not raw_text:
        return {"value": "", "score": 0, "status": "пусто"}

    cell_key = normalize_for_match(raw_text)
    if not cell_key:
        return {"value": "", "score": 0, "status": "пусто"}

    address_like = has_address_marker(cell_key)

    for item in reference:
        full_key = item["full_key"]
        if not full_key or not is_reference_allowed_for_cell(item, cell_key, address_like):
            continue

        if cell_key == full_key or re.search(rf"\b{re.escape(full_key)}\b", cell_key):
            return {"value": item["clean"], "score": 100, "status": "точное совпадение"}

    choices = {
        item["full_key"]: item
        for item in reference
        if item["full_key"] and is_reference_allowed_for_cell(item, cell_key, address_like)
    }
    best = process.extractOne(cell_key, list(choices.keys()), scorer=fuzz.WRatio) if choices else None

    if best:
        best_key, score, _ = best
        if score >= MATCH_THRESHOLD:
            item = choices[best_key]
            return {"value": item["clean"], "score": round(score, 1), "status": "совпадение от 94%"}

    return {"value": "", "score": 0, "status": "совпадение ниже 94% или адрес без НП"}


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

    return pd.DataFrame(
        [
            {"Показатель": "Всего строк", "Значение": len(df)},
            {"Показатель": "Найдено совпадений от 94%", "Значение": found_count},
            {"Показатель": "Не найдено совпадений от 94%", "Значение": empty_result_count},
            {"Показатель": "Исходный столбец", "Значение": source_column},
            {"Показатель": "Столбец результата", "Значение": result_column},
            {"Показатель": "Столбец справочника", "Значение": ref_column},
            {"Показатель": "Значений в справочнике", "Значение": reference_count},
        ]
    )


def prepare_export_dataframe(parsed_df: pd.DataFrame) -> pd.DataFrame:
    export_df = parsed_df.copy()
    columns_to_drop = [
        export_df.columns[index]
        for index in DELETE_EXCEL_COLUMN_INDEXES
        if index < len(export_df.columns)
    ]
    if columns_to_drop:
        export_df = export_df.drop(columns=columns_to_drop)
    return export_df


def make_excel(parsed_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    export_df = prepare_export_dataframe(parsed_df)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Исправленный файл")
        summary_df.to_excel(writer, index=False, sheet_name="Сводка")

        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})

        for sheet_name, sheet_df in {
            "Исправленный файл": export_df,
            "Сводка": summary_df,
        }.items():
            worksheet = writer.sheets[sheet_name]
            for col_num, value in enumerate(sheet_df.columns):
                worksheet.write(0, col_num, value, header_format)
                width = min(max(len(str(value)) + 4, 14), 42)
                worksheet.set_column(col_num, col_num, width)
            worksheet.freeze_panes(1, 0)

        result_worksheet = writer.sheets["Исправленный файл"]
        for column_range in HIDE_EXCEL_RANGES:
            result_worksheet.set_column(column_range, None, None, {"hidden": True})

    return output.getvalue()


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.caption(
    "Приложение анализирует выбранную ячейку и записывает только конкретный населенный пункт "
    "из эталонного Google справочника. Совпадение от 94% считается точным. "
    "Татарстан как регион не записывается. Если в ячейке только адрес города без явного НП из справочника, результат остается пустым."
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
    st.write(f"Уникальных значений после фильтрации: {len(reference)}")
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
    "Найдено от 94%",
    int(summary_df.loc[summary_df["Показатель"] == "Найдено совпадений от 94%", "Значение"].iloc[0]),
)
right.metric(
    "Не найдено",
    int(summary_df.loc[summary_df["Показатель"] == "Не найдено совпадений от 94%", "Значение"].iloc[0]),
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
