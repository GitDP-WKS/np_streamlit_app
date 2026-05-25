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

ADDRESS_RE = re.compile(r"(?i)(?:^|[\s,;:])адрес(?:\s+[а-яa-z0-9_ -]+)?\s*[:\-–—]?\s*")
STOP_RE = re.compile(r"(?i)\b(фио|заявитель|потребитель|телефон|контакт|договор|комментарий|примечание|описание|вопрос|л\s*с|лицевой\s+счет)\b")
ADDRESS_MARKERS = ["ул", "улица", "пер", "переулок", "пр", "проспект", "д", "дом", "корп", "корпус", "кв", "квартира", "шоссе", "тракт", "набережная"]
REGION_WORDS = ["республика татарстан", "респ татарстан", "татарстан", "рт", "муниципальный район", "район", "р н", "мр"]
BLOCKED_KEYS = {"татарстан", "республика татарстан", "респ татарстан", "рт"}


def compact_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("ё", "е").strip()
    return re.sub(r"\s+", " ", text)


def normalize_for_match(value: object) -> str:
    text = compact_text(value).lower()
    text = text.replace("«", " ").replace("»", " ")
    text = re.sub(r"[\"'`.,;:()№/\\\-]+", " ", text)
    for word in REGION_WORDS:
        text = re.sub(rf"\b{re.escape(word)}\b", " ", text)
    text = re.sub(r"\bн\s*п\b", "нп", text)
    text = re.sub(r"\bп\s*г\s*т\b", "пгт", text)
    text = re.sub(r"\bж\s*д\s*ст\b", "жд ст", text)
    return re.sub(r"\s+", " ", text).strip()


def strip_type(value: str) -> str:
    text = normalize_for_match(value)
    for pattern in [r"^жд\s+ст\s+", r"^нп\s+", r"^пгт\s+", r"^снт\s+", r"^г\s+", r"^с\s+", r"^д\s+", r"^п\s+"]:
        text = re.sub(pattern, "", text)
    return text.strip()


def clean_np(value: object) -> str:
    text = compact_text(value)
    if not text:
        return ""
    text = text.replace("“", "«").replace("”", "»").replace('"', "«")
    m = re.match(r"(?i)^снт\s+[«\"]?(.+?)[»\"]?$", text)
    if m:
        name = m.group(1).strip(" «»\"")
        return f"СНТ «{name}»"
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


def is_blocked_key(key: str) -> bool:
    return not key or key in BLOCKED_KEYS or "татарстан" in key


def is_typed_key(key: str) -> bool:
    return bool(re.match(r"^(г|с|д|п|нп|пгт|снт)\s+", key or "") or re.match(r"^жд\s+ст\s+", key or ""))


def is_city_or_short(key: str) -> bool:
    parts = (key or "").split()
    return len(parts) == 1 or (len(parts) == 2 and parts[0] == "г")


def has_address_marker(text: str) -> bool:
    return any(re.search(rf"\b{re.escape(marker)}\b", text) for marker in ADDRESS_MARKERS)


def extract_address_text(value: object) -> str:
    text = compact_text(value)
    if not text:
        return ""
    matches = list(ADDRESS_RE.finditer(text))
    if not matches:
        return ""
    address = text[matches[-1].end():].strip(" :;-–—")
    stop = STOP_RE.search(address)
    if stop:
        address = address[:stop.start()].strip(" :;-–—")
    return address


def split_address_to_candidates(address: object) -> list[str]:
    text = compact_text(address)
    if not text:
        return []
    candidates = [normalize_for_match(text)]
    for piece in re.split(r"[,;\n\r]+", text):
        normalized = normalize_for_match(piece)
        if normalized:
            candidates.append(normalized)
        head = re.split(r"\b(ул|улица|пер|переулок|пр|проспект|д|дом|корп|корпус|кв|квартира|шоссе|тракт|набережная)\b", normalized, maxsplit=1)[0].strip()
        if head:
            candidates.append(head)
    result, seen = [], set()
    for item in candidates:
        if item and item not in seen:
            result.append(item)
            seen.add(item)
    return result


def load_reference() -> pd.DataFrame:
    return pd.read_excel(REFERENCE_URL)


def find_reference_column(df: pd.DataFrame) -> str:
    cols = {compact_text(col).lower(): col for col in df.columns}
    for candidate in REFERENCE_COLUMN_CANDIDATES:
        found = cols.get(compact_text(candidate).lower())
        if found:
            return found
    text_cols = [col for col in df.columns if df[col].dtype == object]
    return text_cols[0] if text_cols else df.columns[0]


def build_reference(df: pd.DataFrame, column: str) -> list[dict[str, object]]:
    values = df[column].dropna().map(compact_text)
    values = values[values != ""].drop_duplicates().tolist()
    reference = []
    for value in values:
        full_key = normalize_for_match(value)
        if is_blocked_key(full_key):
            continue
        reference.append({
            "clean": clean_np(value),
            "full_key": full_key,
            "name_key": strip_type(full_key),
            "is_typed": is_typed_key(full_key),
            "is_city_or_short": is_city_or_short(full_key),
        })
    return sorted(reference, key=lambda row: len(row["full_key"]), reverse=True)


def allowed(item: dict[str, object], address_key: str, address_like: bool) -> bool:
    full_key = str(item["full_key"])
    if is_blocked_key(full_key):
        return False
    if address_like and bool(item["is_city_or_short"]) and address_key != full_key:
        return False
    return True


def match_reference_cell(value: object, reference: list[dict[str, object]]) -> dict[str, object]:
    address = extract_address_text(value)
    if not address:
        return {"value": "", "score": 0, "status": "нет блока адреса"}
    address_key = normalize_for_match(address)
    if not address_key:
        return {"value": "", "score": 0, "status": "пустой блок адреса"}
    address_like = has_address_marker(address_key)
    candidates = split_address_to_candidates(address)

    for candidate in candidates:
        for item in reference:
            if not allowed(item, address_key, address_like):
                continue
            full_key = str(item["full_key"])
            name_key = str(item["name_key"])
            if candidate == full_key or re.search(rf"\b{re.escape(full_key)}\b", candidate):
                return {"value": item["clean"], "score": 100, "status": "найдено после адреса"}
            if bool(item["is_typed"]) and name_key and re.search(rf"\b{re.escape(name_key)}\b", candidate):
                if not (address_like and bool(item["is_city_or_short"])):
                    return {"value": item["clean"], "score": 99, "status": "найдено название после адреса"}

    choices = {}
    for item in reference:
        if not allowed(item, address_key, address_like):
            continue
        choices[str(item["full_key"])] = item
        if bool(item["is_typed"]) and str(item["name_key"]) and not bool(item["is_city_or_short"]):
            choices[str(item["name_key"])] = item

    best_result = None
    for candidate in candidates:
        best = process.extractOne(candidate, list(choices.keys()), scorer=fuzz.WRatio) if choices else None
        if best and best[1] >= MATCH_THRESHOLD:
            item = choices[best[0]]
            if best_result is None or best[1] > best_result["score"]:
                best_result = {"value": item["clean"], "score": round(best[1], 1), "status": "совпадение от 94% после адреса"}
    return best_result or {"value": "", "score": 0, "status": "НП не найден после адреса"}


def read_excel_sheets(uploaded_file) -> dict[str, pd.DataFrame]:
    return pd.read_excel(uploaded_file, sheet_name=None)


def choose_columns(df: pd.DataFrame) -> tuple[str, str]:
    columns = list(df.columns)
    st.subheader("Выбор столбцов")
    left, right = st.columns(2)
    with left:
        source_column = st.selectbox("Где находится текст для анализа", columns, key="source_column")
    with right:
        mode = st.radio("Куда записать найденный НП", ["Создать новую колонку", "Записать в существующую колонку"], horizontal=True)
        if mode == "Создать новую колонку":
            result_column = st.text_input("Название колонки результата", value=DEFAULT_RESULT_COLUMN)
        else:
            result_column = st.selectbox("Столбец для записи результата", columns, key="result_column")
    result_column = compact_text(result_column)
    if not result_column:
        st.error("Укажите колонку для результата.")
        st.stop()
    return source_column, result_column


def parse_dataframe(df: pd.DataFrame, source_column: str, result_column: str, reference: list[dict[str, object]]) -> pd.DataFrame:
    result = df.copy()
    matches = result[source_column].apply(lambda value: match_reference_cell(value, reference))
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
    cols = [result.columns[i] for i in DELETE_EXCEL_COLUMN_INDEXES if i < len(result.columns)]
    return result.drop(columns=cols) if cols else result


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
            for col_num, value in enumerate(sheet_df.columns):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, min(max(len(str(value)) + 4, 14), 42))
            worksheet.freeze_panes(1, 0)
        sheet = writer.sheets["Исправленный файл"]
        for col_range in HIDE_EXCEL_RANGES:
            sheet.set_column(col_range, None, None, {"hidden": True})
    return output.getvalue()


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption("Поиск НП выполняется только в тексте после слова 'адрес'. Справочник Google перечитывается при каждом запуске приложения.")

try:
    reference_df = load_reference()
except Exception as error:
    st.error(f"Не удалось загрузить справочник: {error}")
    st.stop()

ref_column = find_reference_column(reference_df)
reference = build_reference(reference_df, ref_column)

with st.expander("Справочник", expanded=False):
    st.write(f"Столбец справочника: {ref_column}")
    st.write(f"Значений после фильтрации: {len(reference)}")
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

if not st.button("Запустить анализ", type="primary"):
    st.info("Выберите исходный столбец и столбец результата, затем нажмите кнопку запуска.")
    st.stop()

reference_df = load_reference()
ref_column = find_reference_column(reference_df)
reference = build_reference(reference_df, ref_column)
parsed_df = parse_dataframe(source_df, source_column, result_column, reference)
summary_df = build_summary(parsed_df, source_column, result_column, ref_column, len(reference))

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
