import io
import re

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process

APP_TITLE = "Умный парсинг населенных пунктов"
REFERENCE_URL = "https://docs.google.com/spreadsheets/d/1UhrxXABHp5yrtiAm7excLPicHhlgZrgN7-8R3Ada9ZY/export?format=xlsx&gid=0"
REFERENCE_COLUMN_CANDIDATES = ["сокр.Населенный пункт", "Населенный пункт", "НП", "Наименование населенного пункта"]
MATCH_THRESHOLD = 94
DEFAULT_RESULT_COLUMN = "НП из справочника"
DEFAULT_SCORE_COLUMN = "Точность НП"
DEFAULT_STATUS_COLUMN = "Статус парсинга НП"
DELETE_EXCEL_COLUMN_INDEXES = [41]
HIDE_EXCEL_RANGES = ["A:B", "E:U", "Y:AM"]

ADDRESS_START_RE = re.compile(r"(?i)(?:^|[\s,;:])адрес(?:\s+[а-яa-z0-9_ -]+)?\s*[:\-–—]?\s*")
ADDRESS_STOP_RE = re.compile(r"(?i)\b(фио|заявитель|потребитель|телефон|контакт|договор|комментарий|примечание|описание|вопрос|лицевой\s+счет|л\s*с)\b")
ADDRESS_TAIL_RE = re.compile(r"\b(ул|улица|пер|переулок|пр|проспект|д|дом|корп|корпус|кв|квартира|шоссе|тракт|набережная)\b", re.IGNORECASE)
REGION_WORDS = ["республика татарстан", "респ татарстан", "татарстан", "рт", "муниципальный район", "район", "р н", "мр"]
BLOCKED_KEYS = {"татарстан", "республика татарстан", "респ татарстан", "рт"}
TYPE_PREFIX_RE = re.compile(r"^(жд\s+ст|нп|пгт|снт|г|с|д|п)\s+")
TYPED_RE = re.compile(r"^(г|с|д|п|нп|пгт|снт)\s+|^жд\s+ст\s+")


def compact_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("ё", "е").strip())


def normalize_text(value: object) -> str:
    text = compact_text(value).lower().replace("«", " ").replace("»", " ")
    text = re.sub(r"[\"'`.,;:()№/\\\-]+", " ", text)
    for word in REGION_WORDS:
        text = re.sub(rf"\b{re.escape(word)}\b", " ", text)
    text = re.sub(r"\bн\s*п\b", "нп", text)
    text = re.sub(r"\bп\s*г\s*т\b", "пгт", text)
    text = re.sub(r"\bж\s*д\s*ст\b", "жд ст", text)
    return re.sub(r"\s+", " ", text).strip()


def strip_np_type(value: str) -> str:
    return TYPE_PREFIX_RE.sub("", normalize_text(value)).strip()


def clean_reference_value(value: object) -> str:
    text = compact_text(value)
    if not text:
        return ""

    text = text.replace("“", "«").replace("”", "»").replace('"', "«")
    snt_match = re.match(r"(?i)^снт\s+[«\"]?(.+?)[»\"]?$", text)
    if snt_match:
        return f"СНТ «{snt_match.group(1).strip(' «»\"')}»"

    replacements = [
        (r"(?i)^н\s*\.?\s*п\s*\.?\s+", "н.п. "),
        (r"(?i)^п\s*\.?\s*г\s*\.?\s*т\s*\.?\s+", "пгт "),
        (r"(?i)^ж\s*/?\s*д\s*\.?\s*ст\s*\.?\s+", "ж/д ст "),
        (r"(?i)^г\s*\.?\s+", "г. "),
        (r"(?i)^с\s*\.?\s+", "с. "),
        (r"(?i)^д\s*\.?\s+", "д. "),
        (r"(?i)^п\s*\.?\s+", "п. "),
    ]
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)
    return re.sub(r"\s+", " ", text).strip()


def is_bad_reference_key(key: str) -> bool:
    return not key or key in BLOCKED_KEYS or "татарстан" in key


def is_typed_reference(key: str) -> bool:
    return bool(TYPED_RE.match(key or ""))


def is_short_city_reference(key: str) -> bool:
    parts = key.split()
    return len(parts) == 1 or (len(parts) == 2 and parts[0] == "г")


def extract_address_block(value: object) -> str:
    text = compact_text(value)
    if not text:
        return ""

    matches = list(ADDRESS_START_RE.finditer(text))
    if not matches:
        return ""

    address = text[matches[-1].end():].strip(" :;-–—")
    stop_match = ADDRESS_STOP_RE.search(address)
    if stop_match:
        address = address[:stop_match.start()].strip(" :;-–—")
    return address


def make_address_candidates(address: object) -> list[str]:
    text = compact_text(address)
    if not text:
        return []

    parts = [text, *re.split(r"[,;\n\r]+", text)]
    candidates: list[str] = []

    for part in parts:
        normalized = normalize_text(part)
        if normalized:
            candidates.append(normalized)

        head = ADDRESS_TAIL_RE.split(normalized, maxsplit=1)[0].strip() if normalized else ""
        if head:
            candidates.append(head)

    return list(dict.fromkeys(candidates))


def load_reference() -> pd.DataFrame:
    return pd.read_excel(REFERENCE_URL)


def find_reference_column(df: pd.DataFrame) -> str:
    normalized_columns = {compact_text(col).lower(): col for col in df.columns}
    for candidate in REFERENCE_COLUMN_CANDIDATES:
        column = normalized_columns.get(compact_text(candidate).lower())
        if column:
            return column

    text_columns = [col for col in df.columns if df[col].dtype == object]
    return text_columns[0] if text_columns else df.columns[0]


def build_reference_index(df: pd.DataFrame, column: str) -> dict[str, object]:
    values = df[column].dropna().map(compact_text)
    values = values[values != ""].drop_duplicates().tolist()

    items: list[dict[str, object]] = []
    exact_index: dict[str, dict[str, object]] = {}
    name_index: dict[str, dict[str, object]] = {}
    fuzzy_choices: dict[str, dict[str, object]] = {}

    for value in values:
        full_key = normalize_text(value)
        if is_bad_reference_key(full_key):
            continue

        name_key = strip_np_type(full_key)
        item = {
            "result": clean_reference_value(value),
            "full_key": full_key,
            "name_key": name_key,
            "typed": is_typed_reference(full_key),
            "short_city": is_short_city_reference(full_key),
        }
        items.append(item)
        exact_index[full_key] = item

        if item["typed"] and name_key and not item["short_city"]:
            name_index[name_key] = item
            fuzzy_choices[name_key] = item

        fuzzy_choices[full_key] = item

    items.sort(key=lambda row: len(str(row["full_key"])), reverse=True)
    fuzzy_keys = list(fuzzy_choices.keys())
    return {"items": items, "exact": exact_index, "names": name_index, "fuzzy": fuzzy_choices, "fuzzy_keys": fuzzy_keys}


def allowed_for_address(item: dict[str, object], address_key: str, address_has_street: bool) -> bool:
    if address_has_street and item["short_city"] and address_key != item["full_key"]:
        return False
    return True


def lookup_exact(candidates: list[str], ref: dict[str, object], address_key: str, address_has_street: bool) -> dict[str, object] | None:
    exact_index = ref["exact"]
    name_index = ref["names"]

    for candidate in candidates:
        item = exact_index.get(candidate)
        if item and allowed_for_address(item, address_key, address_has_street):
            return {"value": item["result"], "score": 100, "status": "найдено после адреса"}

        item = name_index.get(candidate)
        if item and allowed_for_address(item, address_key, address_has_street):
            return {"value": item["result"], "score": 99, "status": "найдено название после адреса"}

    return None


def lookup_substring(candidates: list[str], ref: dict[str, object], address_key: str, address_has_street: bool) -> dict[str, object] | None:
    for candidate in candidates:
        for item in ref["items"]:
            if not allowed_for_address(item, address_key, address_has_street):
                continue

            full_key = str(item["full_key"])
            name_key = str(item["name_key"])

            if full_key and re.search(rf"\b{re.escape(full_key)}\b", candidate):
                return {"value": item["result"], "score": 100, "status": "найдено внутри адреса"}

            if item["typed"] and name_key and not item["short_city"] and re.search(rf"\b{re.escape(name_key)}\b", candidate):
                return {"value": item["result"], "score": 99, "status": "найдено название внутри адреса"}

    return None


def lookup_fuzzy(candidates: list[str], ref: dict[str, object], address_key: str, address_has_street: bool) -> dict[str, object] | None:
    best_result = None
    fuzzy_keys = ref["fuzzy_keys"]
    fuzzy_index = ref["fuzzy"]

    for candidate in candidates:
        best = process.extractOne(candidate, fuzzy_keys, scorer=fuzz.WRatio) if fuzzy_keys else None
        if not best or best[1] < MATCH_THRESHOLD:
            continue

        item = fuzzy_index[best[0]]
        if not allowed_for_address(item, address_key, address_has_street):
            continue

        if best_result is None or best[1] > best_result["score"]:
            best_result = {"value": item["result"], "score": round(best[1], 1), "status": "совпадение от 94% после адреса"}

    return best_result


def find_settlement_in_address(address: str, ref: dict[str, object]) -> dict[str, object]:
    address_key = normalize_text(address)
    if not address_key:
        return {"value": "", "score": 0, "status": "пустой блок адреса"}

    address_has_street = bool(ADDRESS_TAIL_RE.search(address_key))
    candidates = make_address_candidates(address)

    for lookup in (lookup_exact, lookup_substring, lookup_fuzzy):
        result = lookup(candidates, ref, address_key, address_has_street)
        if result:
            return result

    return {"value": "", "score": 0, "status": "НП не найден после адреса"}


def match_cell(value: object, ref: dict[str, object]) -> dict[str, object]:
    address = extract_address_block(value)
    if not address:
        return {"value": "", "score": 0, "status": "нет блока адреса"}
    return find_settlement_in_address(address, ref)


def read_excel_sheets(uploaded_file) -> dict[str, pd.DataFrame]:
    return pd.read_excel(uploaded_file, sheet_name=None)


def choose_columns(df: pd.DataFrame) -> tuple[str, str]:
    columns = list(df.columns)
    st.subheader("Выбор столбцов")

    left, right = st.columns(2)
    with left:
        source_column = st.selectbox("Где находится текст для анализа", columns)
    with right:
        mode = st.radio("Куда записать найденный НП", ["Создать новую колонку", "Записать в существующую колонку"], horizontal=True)
        result_column = st.text_input("Название колонки результата", value=DEFAULT_RESULT_COLUMN) if mode == "Создать новую колонку" else st.selectbox("Столбец для записи результата", columns)

    result_column = compact_text(result_column)
    if not result_column:
        st.error("Укажите колонку для результата.")
        st.stop()
    return source_column, result_column


def analyze_dataframe(df: pd.DataFrame, source_column: str, result_column: str, ref: dict[str, object]) -> pd.DataFrame:
    result = df.copy()
    matches = result[source_column].apply(lambda value: match_cell(value, ref))
    match_df = pd.DataFrame(list(matches))

    result[result_column] = match_df["value"]
    result[DEFAULT_SCORE_COLUMN] = match_df["score"]
    result[DEFAULT_STATUS_COLUMN] = match_df["status"]
    return result


def build_summary(df: pd.DataFrame, source_column: str, result_column: str, ref_column: str, reference_count: int) -> pd.DataFrame:
    empty_count = int((df[result_column].astype(str).str.strip() == "").sum())
    found_count = len(df) - empty_count
    return pd.DataFrame([
        {"Показатель": "Всего строк", "Значение": len(df)},
        {"Показатель": "Найдено НП после слова адрес", "Значение": found_count},
        {"Показатель": "Не найдено НП после слова адрес", "Значение": empty_count},
        {"Показатель": "Исходный столбец", "Значение": source_column},
        {"Показатель": "Столбец результата", "Значение": result_column},
        {"Показатель": "Столбец справочника", "Значение": ref_column},
        {"Показатель": "Значений в справочнике", "Значение": reference_count},
    ])


def prepare_export_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    columns_to_drop = [result.columns[index] for index in DELETE_EXCEL_COLUMN_INDEXES if index < len(result.columns)]
    return result.drop(columns=columns_to_drop) if columns_to_drop else result


def make_excel(parsed_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    export_df = prepare_export_dataframe(parsed_df)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Исправленный файл")
        summary_df.to_excel(writer, index=False, sheet_name="Сводка")

        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})

        for sheet_name, sheet_df in {"Исправленный файл": export_df, "Сводка": summary_df}.items():
            worksheet = writer.sheets[sheet_name]
            for column_index, column_name in enumerate(sheet_df.columns):
                worksheet.write(0, column_index, column_name, header_format)
                worksheet.set_column(column_index, column_index, min(max(len(str(column_name)) + 4, 14), 42))
            worksheet.freeze_panes(1, 0)

        result_sheet = writer.sheets["Исправленный файл"]
        for column_range in HIDE_EXCEL_RANGES:
            result_sheet.set_column(column_range, None, None, {"hidden": True})

    return output.getvalue()


def load_current_reference() -> tuple[pd.DataFrame, str, dict[str, object]]:
    reference_df = load_reference()
    ref_column = find_reference_column(reference_df)
    ref = build_reference_index(reference_df, ref_column)
    return reference_df, ref_column, ref


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption("Поиск НП выполняется только в тексте после слова 'адрес'. Справочник Google перечитывается перед каждым анализом.")

try:
    preview_reference_df, preview_ref_column, preview_ref = load_current_reference()
except Exception as error:
    st.error(f"Не удалось загрузить справочник: {error}")
    st.stop()

with st.expander("Справочник", expanded=False):
    st.write(f"Столбец справочника: {preview_ref_column}")
    st.write(f"Значений после фильтрации: {len(preview_ref['items'])}")
    st.dataframe(preview_reference_df.head(20), use_container_width=True)

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

if not st.button("Запустить анализ", type="primary"):
    st.info("Выберите исходный столбец и столбец результата, затем нажмите кнопку запуска.")
    st.stop()

try:
    _, ref_column, ref = load_current_reference()
except Exception as error:
    st.error(f"Не удалось обновить справочник перед анализом: {error}")
    st.stop()

parsed_df = analyze_dataframe(source_df, source_column, result_column, ref)
summary_df = build_summary(parsed_df, source_column, result_column, ref_column, len(ref["items"]))

left, middle, right = st.columns(3)
left.metric("Всего строк", len(parsed_df))
middle.metric("Найдено после адреса", int(summary_df.loc[summary_df["Показатель"] == "Найдено НП после слова адрес", "Значение"].iloc[0]))
right.metric("Не найдено", int(summary_df.loc[summary_df["Показатель"] == "Не найдено НП после слова адрес", "Значение"].iloc[0]))

st.subheader("Сводка")
st.dataframe(summary_df, use_container_width=True)
st.subheader("Исправленный файл")
st.dataframe(parsed_df, use_container_width=True, height=480)

st.download_button(
    "Скачать результат Excel",
    data=make_excel(parsed_df, summary_df),
    file_name="np_ai_parse_result.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
