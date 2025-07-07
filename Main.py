import os
import json
import pandas as pd
from datetime import datetime
import re

PRIORITY_COLS = [
    'SourceFile',
    'tournamentId',
    'employeeNumber',
    'lastName',
    'firstName',
    'divisionRatings_BANK_groupId',
    'terDivisionName',
    'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId',
    'employeeStatus',
    'businessBlock',
    'successValue',
    'indicatorValue',
    'divisionRatings_BANK_placeInRating',
    'divisionRatings_TB_placeInRating',
    'divisionRatings_GOSB_placeInRating',
    'divisionRatings_BANK_ratingCategoryName',
    'divisionRatings_TB_ratingCategoryName',
    'divisionRatings_GOSB_ratingCategoryName',
]

# Список полей для преобразования в int
INT_FIELDS = [
    'divisionRatings_BANK_groupId',
    'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId',
    'divisionRatings_BANK_placeInRating',
    'divisionRatings_TB_placeInRating',
    'divisionRatings_GOSB_placeInRating',
]

# Список полей для преобразования в float с 3 знаками после запятой
FLOAT_FIELDS = [
    'indicatorValue',
    'successValue',
]

def parse_float(val):
    """Преобразует строку к float с 3 знаками (убирает все пробелы, в т.ч. тонкие и неразрывные)."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return round(float(val), 3)
    s = str(val)
    s = re.sub(r'[\s\u00A0\u2009]', '', s)
    s = re.sub(r"[^\d.,\-]", "", s)
    if s.count(',') > 0 and s.count('.') > 0:
        if s.rfind('.') > s.rfind(','):
            s = s.replace(',', '')
        else:
            s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try:
        return round(float(s), 3)
    except Exception:
        return None

def parse_int(val):
    """Преобразует строку к int (убирает все пробелы и нечисловые символы)."""
    if val is None:
        return None
    if isinstance(val, int):
        return val
    s = str(val)
    s = re.sub(r'[\s\u00A0\u2009]', '', s)
    s = re.sub(r"[^\d\-]", "", s)
    try:
        return int(s)
    except Exception:
        return None

def flatten_leader(leader, tournament_id, source_file):
    """Плоская строка по участнику + tournamentId + имя файла. Исключаем photoData."""
    row = {
        'SourceFile': source_file,
        'tournamentId': tournament_id
    }
    # Прямые поля (без photoData и indicatorValue/successValue)
    for k, v in leader.items():
        if k in ("divisionRatings", "photoData"):
            continue
        if k in FLOAT_FIELDS:
            row[k] = parse_float(v)
        else:
            row[k] = v
    # divisionRatings по groupCode (BANK, TB, GOSB)
    if "divisionRatings" in leader and leader["divisionRatings"]:
        for div in leader["divisionRatings"]:
            group = div.get("groupCode")
            if not group:
                continue
            for field in ("groupId", "placeInRating", "ratingCategoryName"):
                colname = f"divisionRatings_{group}_{field}"
                if field in div:
                    value = div[field]
                    if colname in INT_FIELDS:
                        row[colname] = parse_int(value)
                    elif colname in FLOAT_FIELDS:
                        row[colname] = parse_float(value)
                    else:
                        row[colname] = value
    # Убеждаемся, что все FLOAT_FIELDS обработаны (на случай если не было в цикле выше)
    for f in FLOAT_FIELDS:
        if f not in row:
            row[f] = None
    return row

def process_json_file(filepath):
    """Парсит leaders, добавляет имя файла, исключает photoData."""
    filename = os.path.basename(filepath)
    with open(filepath, 'r', encoding='utf-8') as f:
        js = json.load(f)
    all_rows = []
    for block in js.values():
        blocks = block if isinstance(block, list) else [block]
        for subblock in blocks:
            tournament = subblock.get("body", {}).get("tournament", {})
            tournament_id = tournament.get("tournamentId")
            for leader in tournament.get("leaders", []):
                all_rows.append(flatten_leader(leader, tournament_id, filename))
    return pd.DataFrame(all_rows)

def align_and_sort(df, all_columns):
    """Выравнивает датафрейм под нужный порядок столбцов, добавляет пустые если их нет."""
    for col in all_columns:
        if col not in df.columns:
            df[col] = None
    rest = [c for c in df.columns if c not in all_columns]
    return df[all_columns + rest]

def main(
    source_dir: str,
    target_dir: str,
    before_filename: str,
    after_filename: str,
    result_excel: str
):
    print(f"[START] Экспорт лидеров по турнирам...")
    before_path = os.path.join(source_dir, before_filename)
    after_path = os.path.join(source_dir, after_filename)

    now = datetime.now()
    ts = now.strftime("%Y%m%d_%H%M%S")
    sheet_before = f"BEFORE_{ts}"
    sheet_after = f"AFTER_{ts}"

    df_before = process_json_file(before_path)
    print(f"[INFO] BEFORE: строк {len(df_before)}, колонок {len(df_before.columns)}")
    df_after = process_json_file(after_path)
    print(f"[INFO] AFTER: строк {len(df_after)}, колонок {len(df_after.columns)}")

    # Все уникальные колонки для обеих таблиц
    all_cols = PRIORITY_COLS.copy()
    all_cols += [c for c in set(df_before.columns).union(df_after.columns) if c not in all_cols]

    df_before = align_and_sort(df_before, all_cols)
    df_after = align_and_sort(df_after, all_cols)
    print(f"[INFO] Итоговое количество столбцов: {len(df_before.columns)}")

    # Имя Excel с таймштампом
    base, ext = os.path.splitext(result_excel)
    result_excel_ts = f"{base}_{ts}{ext}"
    out_excel = os.path.join(target_dir, result_excel_ts)

    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
        df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        print(f"[OK] BEFORE экспортирован: {len(df_before)} строк, {len(df_before.columns)} колонок.")
        df_after.to_excel(writer, index=False, sheet_name=sheet_after)
        print(f"[OK] AFTER экспортирован: {len(df_after)} строк, {len(df_after.columns)} колонок.")
    print(f"[SUCCESS] Файл {out_excel} создан.")

if __name__ == "__main__":
    main(
        source_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON",
        target_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX",
        before_filename="LFA_0.json",
        after_filename="LFA_2.json",
        result_excel="LFA_COMPARE.xlsx"
    )
