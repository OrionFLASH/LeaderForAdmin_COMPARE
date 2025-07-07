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
    'terDivisionName',
    'divisionRatings_BANK_groupId',
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

INT_FIELDS = [
    'divisionRatings_BANK_groupId',
    'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId',
    'divisionRatings_BANK_placeInRating',
    'divisionRatings_TB_placeInRating',
    'divisionRatings_GOSB_placeInRating',
]

FLOAT_FIELDS = [
    'indicatorValue',
    'successValue',
]

def parse_float(val):
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
    row = {
        'SourceFile': source_file,
        'tournamentId': tournament_id
    }
    for k, v in leader.items():
        if k in ("divisionRatings", "photoData"):
            continue
        if k in FLOAT_FIELDS:
            row[k] = parse_float(v)
        else:
            row[k] = v
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
    for f in FLOAT_FIELDS:
        if f not in row:
            row[f] = None
    return row

def process_json_file(filepath):
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
    for col in all_columns:
        if col not in df.columns:
            df[col] = None
    rest = [c for c in df.columns if c not in all_columns]
    return df[all_columns + rest]

def log_columns(df, label, final_columns):
    print(f"[LOG] {label} — исходно {len(df.columns)} колонок: {sorted(list(df.columns))}")
    print(f"[LOG] {label} — сохраняется {len(final_columns)} колонок: {final_columns}")
    added = [col for col in final_columns if col not in df.columns]
    removed = [col for col in df.columns if col not in final_columns]
    if added:
        print(f"[LOG] {label} — добавлены новые пустые колонки: {added}")
    if removed:
        print(f"[LOG] {label} — эти колонки будут удалены: {removed}")

def log_data_stats(df, label):
    print(f"[STAT] {label}: найдено уникальных tournamentId = {df['tournamentId'].nunique()}")
    for tid, group in df.groupby('tournamentId'):
        print(f"[STAT]   tournamentId = {tid}: людей = {len(group)}")

def make_compare_sheet(df_before, df_after, priority_cols, sheet_name):
    """
    Формирует датафрейм сравнения:
    - одна строка на уникальную пару (tournamentId, employeeNumber) из обоих файлов.
    - для каждого поля — пара колонок (BEFORE/AFTER).
    """
    join_keys = ['tournamentId', 'employeeNumber']

    # Оставим только уникальные строки по ключам (на случай дублирования)
    before_uniq = df_before.drop_duplicates(subset=join_keys, keep='last')
    after_uniq  = df_after.drop_duplicates(subset=join_keys, keep='last')

    # Список для compare — все ключи из обоих файлов
    all_keys = pd.concat([before_uniq[join_keys], after_uniq[join_keys]]).drop_duplicates()

    # Добавим суффиксы к столбцам
    before_cols = [c for c in priority_cols if c not in join_keys]
    after_cols = before_cols.copy()
    before_uniq = before_uniq.set_index(join_keys)
    after_uniq  = after_uniq.set_index(join_keys)

    before_uniq = before_uniq.add_suffix('_BEFORE')
    after_uniq  = after_uniq.add_suffix('_AFTER')

    # Объединяем
    compare_df = all_keys.set_index(join_keys) \
        .join(before_uniq, how='left') \
        .join(after_uniq, how='left') \
        .reset_index()

    # Итоговый порядок: ключи, все BEFORE, все AFTER (сохраняем ваш приоритет)
    cols_order = join_keys + \
                 [c + '_BEFORE' for c in before_cols] + \
                 [c + '_AFTER' for c in after_cols]
    compare_df = compare_df.reindex(columns=cols_order)

    print(f"[OK] Compare sheet готов: строк {len(compare_df)}, колонок {len(compare_df.columns)}")
    return compare_df, sheet_name

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
    sheet_compare = f"COMPARE_{ts}"

    df_before = process_json_file(before_path)
    print(f"[INFO] BEFORE: строк {len(df_before)}, колонок {len(df_before.columns)}")
    df_after = process_json_file(after_path)
    print(f"[INFO] AFTER: строк {len(df_after)}, колонок {len(df_after.columns)}")

    log_data_stats(df_before, "BEFORE")
    log_data_stats(df_after, "AFTER")

    all_cols = PRIORITY_COLS.copy()
    all_cols += [c for c in set(df_before.columns).union(df_after.columns) if c not in all_cols]

    log_columns(df_before, "BEFORE", all_cols)
    log_columns(df_after, "AFTER", all_cols)

    df_before = align_and_sort(df_before, all_cols)
    df_after = align_and_sort(df_after, all_cols)
    print(f"[INFO] Итоговое количество столбцов (финальная структура): {len(df_before.columns)}")

    # Формируем compare sheet
    compare_df, sheet_compare = make_compare_sheet(df_before, df_after, all_cols, sheet_compare)

    base, ext = os.path.splitext(result_excel)
    result_excel_ts = f"{base}_{ts}{ext}"
    out_excel = os.path.join(target_dir, result_excel_ts)

    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
        df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        print(f"[OK] BEFORE экспортирован: {len(df_before)} строк, {len(df_before.columns)} колонок.")
        df_after.to_excel(writer, index=False, sheet_name=sheet_after)
        print(f"[OK] AFTER экспортирован: {len(df_after)} строк, {len(df_after.columns)} колонок.")
        compare_df.to_excel(writer, index=False, sheet_name=sheet_compare)
        print(f"[OK] COMPARE экспортирован: {len(compare_df)} строк, {len(compare_df.columns)} колонок.")
    print(f"[SUCCESS] Файл {out_excel} создан.")

if __name__ == "__main__":
    main(
        source_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON",
        target_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX",
        before_filename="LFA_0.json",
        after_filename="LFA_4.json",
        result_excel="LFA_COMPARE.xlsx"
    )
