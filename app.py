import io
import re

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz

# --- НАСТРОЙКИ ---

# Ссылка на гугл-таблицу со справочником (образец)
# Можно поменять на свою, только оставить export?format=xlsx
REFERENCE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1UhrxXABHp5yrtiAm7excLPicHhlgZrgN7-8R3Ada9ZY/export?format=xlsx&gid=0"
)

# Имя колонки в СПРАВОЧНИКЕ с эталонным НП
REF_COL_SHORT = "сокр.Населенный пункт"

# Имена колонок в РЕЕСТРЕ
REESTR_NP_COL = "Населенный пункт"
REESTR_FILIAL_COL = "Филиал"
REESTR_RES_COL = "РЭС"  # сейчас только учитываем, не группируем по нему

# Префиксы типов НП (важно — более длинные сначала)
TYPE_PREFIXES = [
    "ж/д ст",
    "пгт",
    "снт",
    "нп",
    "г",
    "с",
    "д",
    "п",
]

TYPE_MAP = {
    "г": "город",
    "с": "село",
    "д": "деревня",
    "п": "посёлок",
    "пгт": "посёлок городского типа",
    "нп": "населённый пункт",
    "ж/д ст": "железнодорожная станция",
    "снт": "садовое некоммерческое товарищество",
}


# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ РАЗБОРА НП ---

def normalize_text(text: str) -> str:
    """Привести текст к нормальной форме: lower, без лишних символов/пробелов."""
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = text.replace("ё", "е")
    text = re.sub(r"[\"'«».,;:()]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def split_type_and_name(text: str):
    """Разделить строку НП на тип и имя.

    Возвращает: (тип_нп, имя_нп, нормализованный_ключ)
    где нормализованный_ключ = "тип имя" или просто всё выражение без явного типа.
    """
    norm = normalize_text(text)
    if not norm:
        return None, None, ""

    for prefix in TYPE_PREFIXES:
        if norm.startswith(prefix + " "):
            name = norm[len(prefix):].strip()
            key = f"{prefix} {name}".strip()
            return prefix, name, key

    # Тип не распознан, всё считаем именем
    return None, norm, norm


def build_reference_dict(ref_df: pd.DataFrame) -> dict:
    """Построить словарь справочника:
    ключ = нормализованный "тип+имя", значение = данные по НП.
    """
    if REF_COL_SHORT not in ref_df.columns:
        raise ValueError(
            f"В справочнике нет колонки '{REF_COL_SHORT}'. "
            f"Колонки: {list(ref_df.columns)}"
        )

    df = ref_df.dropna(subset=[REF_COL_SHORT]).copy()

    df["np_type"], df["np_name"], df["np_key"] = zip(
        *df[REF_COL_SHORT].apply(split_type_and_name)
    )

    ref_dict = {}
    for _, row in df.iterrows():
        key = row["np_key"]
        if not key:
            continue
        ref_dict[key] = {
            "type": row["np_type"],
            "name": row["np_name"],
            "short": row[REF_COL_SHORT],
            "row": row,
        }

    return ref_dict


def match_np(raw_text: str, ref_dict: dict, score_threshold: int = 90) -> dict:
    """Сопоставить НП из реестра со справочником.

    Возвращает словарь с полями:
    - np_raw, np_type_raw, np_name_raw
    - np_canonical_type, np_canonical_name, np_canonical_short
    - np_match_score, np_match_status
    """
    np_type_raw, np_name_raw, key = split_type_and_name(raw_text)

    base_result = {
        "np_raw": raw_text,
        "np_type_raw": np_type_raw,
        "np_name_raw": np_name_raw,
        "np_canonical_type": None,
        "np_canonical_name": None,
        "np_canonical_short": None,
        "np_match_score": 0,
        "np_match_status": "",
    }

    if not key:
        base_result["np_match_status"] = "empty"
        return base_result

    # 1) Точное совпадение
    if key in ref_dict:
        ref = ref_dict[key]
        base_result.update(
            {
                "np_canonical_type": ref["type"],
                "np_canonical_name": ref["name"],
                "np_canonical_short": ref["short"],
                "np_match_score": 100,
                "np_match_status": "OK",
            }
        )
        return base_result

    # 2) Fuzzy-поиск
    choices = list(ref_dict.keys())
    best = process.extractOne(key, choices, scorer=fuzz.WRatio)

    if best is None:
        base_result["np_match_status"] = "not_found"
        return base_result

    best_key, score, _ = best
    ref = ref_dict[best_key]

    status = "fuzzy_ok" if score >= score_threshold else "suspicious"

    base_result.update(
        {
            "np_canonical_type": ref["type"],
            "np_canonical_name": ref["name"],
            "np_canonical_short": ref["short"],
            "np_match_score": score,
            "np_match_status": status,
        }
    )
    return base_result


def process_reestr(reestr_df: pd.DataFrame, ref_dict: dict) -> pd.DataFrame:
    """Обработать реестр:
    - распарсить НП,
    - найти в справочнике,
    - вернуть исходный DF + новые колонки.
    """
    if REESTR_NP_COL not in reestr_df.columns:
        raise ValueError(
            f"В реестре нет колонки '{REESTR_NP_COL}'. "
            f"Колонки: {list(reestr_df.columns)}"
        )

    results = reestr_df[REESTR_NP_COL].apply(
        lambda x: match_np(x, ref_dict)
    )
    res_df = pd.DataFrame(list(results))

    df_out = pd.concat([reestr_df, res_df], axis=1)
    return df_out


def compute_top5_by_filial(df_out: pd.DataFrame) -> pd.DataFrame:
    """Посчитать ТОП-5 адресов по каждому филиалу.
    Используем каноничную форму 'np_canonical_short'.
    """
    if REESTR_FILIAL_COL not in df_out.columns:
        raise ValueError(
            f"В реестре нет колонки '{REESTR_FILIAL_COL}' (Филиал). "
            f"Колонки: {list(df_out.columns)}"
        )

    tmp = df_out.copy()

    # Берём только те строки, где каноничный адрес определён
    tmp = tmp.dropna(subset=["np_canonical_short"])

    if tmp.empty:
        return pd.DataFrame()

    grp = (
        tmp.groupby([REESTR_FILIAL_COL, "np_canonical_short"])
        .size()
        .reset_index(name="count")
    )

    # ТОП-5 по каждому филиалу
    grp_sorted = grp.sort_values(
        by=[REESTR_FILIAL_COL, "count"],
        ascending=[True, False],
    )

    top5 = grp_sorted.groupby(REESTR_FILIAL_COL).head(5).reset_index(drop=True)
    return top5


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Сохранить DataFrame в память как Excel и вернуть bytes."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="реестр_анализ")
    return output.getvalue()


# --- STREAMLIT UI ---

st.set_page_config(page_title="Анализ населённых пунктов", layout="wide")
st.title("Анализ реестра по населённым пунктам")

st.markdown(
    """
1. Скрипт берёт **справочник НП** (образец) – по умолчанию из Google Sheets.
2. Ты загружаешь **Excel с реестром**.
3. Для каждой строки:
   - парсится столбец **Населенный пункт**,
   - находится соответствие в справочнике,
   - в конце строки появляется каноничный адрес из образца.
4. Дополнительно строится **ТОП-5 адресов по каждому филиалу**.
"""
)

st.sidebar.header("Настройки справочника")

use_custom_ref = st.sidebar.checkbox(
    "Загружать свой файл-справочник (вместо Google Sheets)",
    value=False,
)

ref_file = None
if use_custom_ref:
    ref_file = st.sidebar.file_uploader(
        "Файл со справочником НП (Excel)",
        type=["xlsx", "xls"],
        key="ref_file",
    )
    st.sidebar.caption(
        f"В файле должна быть колонка **{REF_COL_SHORT}** с образцовыми НП."
    )

st.subheader("1. Загрузка файла реестра")

reestr_file = st.file_uploader(
    "Загрузи Excel-файл с реестром",
    type=["xlsx", "xls"],
    key="reestr_file",
)

if reestr_file is not None:
    try:
        reestr_df = pd.read_excel(reestr_file)
    except Exception as e:
        st.error(f"Не получилось прочитать файл реестра: {e}")
        st.stop()

    st.write("Пример данных реестра (первые строки):")
    st.dataframe(reestr_df.head(20))

    # --- Загружаем/строим справочник ---

    st.subheader("2. Загрузка справочника населённых пунктов")

    try:
        if use_custom_ref:
            if ref_file is None:
                st.warning("Включён режим собственного справочника, но файл не загружен.")
                st.stop()
            ref_df = pd.read_excel(ref_file)
        else:
            ref_df = pd.read_excel(REFERENCE_URL)

        ref_dict = build_reference_dict(ref_df)
        st.success(
            f"Справочник загружен. Кол-во уникальных НП в справочнике: {len(ref_dict)}"
        )

    except Exception as e:
        st.error(f"Ошибка при загрузке справочника: {e}")
        st.stop()

    # --- Обработка реестра ---

    st.subheader("3. Анализ и парсинг реестра")

    try:
        df_out = process_reestr(reestr_df, ref_dict)
    except Exception as e:
        st.error(f"Ошибка при обработке реестра: {e}")
        st.stop()

    st.success("Анализ реестра завершён.")

    st.markdown("### Пример результата (первые 50 строк)")
    st.dataframe(df_out.head(50))

    st.markdown(
        """
**Пояснения по новым колонкам:**

- `np_raw` — как НП записан в реестре (исходная ячейка).
- `np_type_raw` — тип НП, выделенный из записи (г/с/д/п/пгт/…).
- `np_name_raw` — имя НП без типа.
- `np_canonical_short` — адрес НП из образца (сокращённая каноничная форма).
- `np_match_status` — статус сопоставления:
  - `OK` — точное совпадение,
  - `fuzzy_ok` — похоже, есть небольшие отличия,
  - `suspicious` — похоже слабо, нужно проверить,
  - `not_found` / `empty` — ничего не нашли.
"""
    )

    # --- ТОП-5 адресов по филиалам ---

    st.subheader("4. ТОП-5 адресов по каждому филиалу")

    try:
        top5_df = compute_top5_by_filial(df_out)
    except Exception as e:
        st.error(f"Ошибка при расчёте ТОП-5: {e}")
        top5_df = pd.DataFrame()

    if top5_df.empty:
        st.info(
            "Не удалось построить ТОП-5: либо нет колонки Филиал, "
            "либо не найдено ни одного каноничного адреса."
        )
    else:
        st.dataframe(top5_df)

    # --- Кнопка скачивания результата ---

    st.subheader("5. Выгрузка результата")

    excel_bytes = to_excel_bytes(df_out)

    st.download_button(
        label="Скачать обработанный реестр (Excel)",
        data=excel_bytes,
        file_name="reestr_processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Загрузи файл реестра, чтобы начать анализ.")
