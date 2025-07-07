import os
import json
import pandas as pd
from datetime import datetime
import re
import logging

# === ПАРАМЕТРЫ ЛОГИРОВАНИЯ ===
LOG_BASENAME = "LOG"
LOG_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//LOGS"  # Папка для логов

def setup_logger(log_dir, basename):
    now = datetime.now()
    day_str = now.strftime("%Y%m%d")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, f"{basename}_{day_str}.log")
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    # Удаляем старые хэндлеры для повторных запусков
    if logger.hasHandlers():
        logger.handlers.clear()
    fh = logging.FileHandler(log_path, encoding='utf-8', mode='a')  # 'a' — append
    fh.setLevel(logging.DEBUG)
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    fmt = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s', "%Y-%m-%d %H:%M:%S")
    fh.setFormatter(fmt)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    logging.info(f"Лог-файл активен (append): {log_path}")
    return logger


# === Ваши списки полей и статусов ===
PRIORITY_COLS = [
    'SourceFile', 'tournamentId', 'employeeNumber', 'lastName', 'firstName',
    'terDivisionName', 'divisionRatings_BANK_groupId', 'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId', 'employeeStatus', 'businessBlock',
    'successValue', 'indicatorValue', 'divisionRatings_BANK_placeInRating',
    'divisionRatings_TB_placeInRating', 'divisionRatings_GOSB_placeInRating',
    'divisionRatings_BANK_ratingCategoryName', 'divisionRatings_TB_ratingCategoryName',
    'divisionRatings_GOSB_ratingCategoryName',
]

COMPARE_KEYS = [
    'tournamentId',
    'employeeNumber',
    'lastName',
    'firstName',
]

COMPARE_FIELDS = [
    'SourceFile',
    'terDivisionName',
    'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId',
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

STATUS_NEW_REMOVE = {
    "both":        "No Change",
    "before_only": "Remove",
    "after_only":  "New"
}
STATUS_INDICATOR = {
    "val_add":      "New ADD",
    "val_remove":   "Remove FROM",
    "val_nochange": "No Change",
    "val_down":     "Change DOWN",
    "val_up":       "Change UP"
}
STATUS_BANK_PLACE = {
    "val_add":      "Rang BANK NEW",
    "val_remove":   "Rang BANK REMOVE",
    "val_nochange": "Rang BANK NO CHANGE",
    "val_up":       "Rang BANK UP",
    "val_down":     "Rang BANK DOWN"
}
STATUS_TB_PLACE = {
    "val_add":      "Rang TB NEW",
    "val_remove":   "Rang TB REMOVE",
    "val_nochange": "Rang TB NO CHANGE",
    "val_up":       "Rang TB UP",
    "val_down":     "Rang TB DOWN"
}
STATUS_GOSB_PLACE = {
    "val_add":      "Rang GOSB NEW",
    "val_remove":   "Rang GOSB REMOVE",
    "val_nochange": "Rang GOSB NO CHANGE",
    "val_up":       "Rang GOSB UP",
    "val_down":     "Rang GOSB DOWN"
}

# === ФУНКЦИИ ПРЕОБРАЗОВАНИЯ ===
def parse_float(val):
    try:
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
        return round(float(s), 3)
    except Exception as ex:
        logging.error(f"Ошибка преобразования '{val}' в float: {ex}")
        return None

def parse_int(val):
    try:
        if val is None:
            return None
        if isinstance(val, int):
            return val
        s = str(val)
        s = re.sub(r'[\s\u00A0\u2009]', '', s)
        s = re.sub(r"[^\d\-]", "", s)
        return int(s)
    except Exception as ex:
        logging.error(f"Ошибка преобразования '{val}' в int: {ex}")
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
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            js = json.load(f)
    except Exception as ex:
        logging.error(f"Ошибка чтения файла {filepath}: {ex}")
        return pd.DataFrame()
    all_rows = []
    for block in js.values():
        blocks = block if isinstance(block, list) else [block]
        for subblock in blocks:
            tournament = subblock.get("body", {}).get("tournament", {})
            tournament_id = tournament.get("tournamentId")
            for leader in tournament.get("leaders", []):
                all_rows.append(flatten_leader(leader, tournament_id, filename))
    logging.info(f"Файл '{filename}': обработано лидеров {len(all_rows)}")
    return pd.DataFrame(all_rows)

def align_and_sort(df, all_columns):
    for col in all_columns:
        if col not in df.columns:
            df[col] = None
    rest = [c for c in df.columns if c not in all_columns]
    return df[all_columns + rest]

def log_columns(df, label, final_columns):
    logging.info(f"{label}: исходно {len(df.columns)} колонок: {sorted(list(df.columns))}")
    logging.info(f"{label}: сохраняется {len(final_columns)} колонок: {final_columns}")
    added = [col for col in final_columns if col not in df.columns]
    removed = [col for col in df.columns if col not in final_columns]
    if added:
        logging.info(f"{label}: добавлены новые пустые колонки: {added}")
    if removed:
        logging.info(f"{label}: эти колонки будут удалены: {removed}")

def log_data_stats(df, label):
    try:
        unique_tids = df['tournamentId'].nunique()
        logging.info(f"{label}: найдено уникальных tournamentId = {unique_tids}")
        for tid, group in df.groupby('tournamentId'):
            logging.info(f"{label}: tournamentId = {tid}: людей = {len(group)}")
    except Exception as ex:
        logging.error(f"{label}: Ошибка лога по tournamentId: {ex}")

def make_compare_sheet(df_before, df_after, compare_keys, compare_fields, sheet_name):
    try:
        join_keys = compare_keys
        before_uniq = df_before.drop_duplicates(subset=join_keys, keep='last')
        after_uniq  = df_after.drop_duplicates(subset=join_keys, keep='last')
        all_keys = pd.concat([before_uniq[join_keys], after_uniq[join_keys]]).drop_duplicates()
        before_uniq = before_uniq.set_index(join_keys)
        after_uniq  = after_uniq.set_index(join_keys)
        before_uniq = before_uniq[compare_fields] if len(before_uniq) else pd.DataFrame(columns=compare_fields)
        after_uniq  = after_uniq[compare_fields]  if len(after_uniq) else pd.DataFrame(columns=compare_fields)
        before_uniq = before_uniq.add_prefix('BEFORE_')
        after_uniq  = after_uniq.add_prefix('AFTER_')
        compare_df = all_keys.set_index(join_keys) \
            .join(before_uniq, how='left') \
            .join(after_uniq, how='left') \
            .reset_index()

        # New_Remove
        def new_remove_row(row):
            before_exist = not pd.isnull(row['BEFORE_indicatorValue']) or not pd.isnull(row['BEFORE_SourceFile'])
            after_exist  = not pd.isnull(row['AFTER_indicatorValue'])  or not pd.isnull(row['AFTER_SourceFile'])
            if before_exist and after_exist:
                return STATUS_NEW_REMOVE['both']
            elif before_exist:
                return STATUS_NEW_REMOVE['before_only']
            elif after_exist:
                return STATUS_NEW_REMOVE['after_only']
            else:
                return ""
        compare_df['New_Remove'] = compare_df.apply(new_remove_row, axis=1)

        # indicatorValue_Compare
        def value_compare(row):
            before = row.get('BEFORE_indicatorValue', None)
            after  = row.get('AFTER_indicatorValue', None)
            if pd.isnull(before) and not pd.isnull(after):
                return STATUS_INDICATOR['val_add']
            if not pd.isnull(before) and pd.isnull(after):
                return STATUS_INDICATOR['val_remove']
            if pd.isnull(before) and pd.isnull(after):
                return ""
            if before == after:
                return STATUS_INDICATOR['val_nochange']
            elif before > after:
                return STATUS_INDICATOR['val_down']
            else:
                return STATUS_INDICATOR['val_up']
        compare_df['indicatorValue_Compare'] = compare_df.apply(value_compare, axis=1)

        # Сравнения placeInRating — индивидуальные статусы
        def rang_compare(row, before_col, after_col, status_dict):
            before = row.get(f'BEFORE_{before_col}', None)
            after  = row.get(f'AFTER_{after_col}', None)
            if pd.isnull(before) and not pd.isnull(after):
                return status_dict['val_add']
            if not pd.isnull(before) and pd.isnull(after):
                return status_dict['val_remove']
            if pd.isnull(before) and pd.isnull(after):
                return ""
            if before == after:
                return status_dict['val_nochange']
            elif before > after:
                return status_dict['val_up']
            else:
                return status_dict['val_down']

        compare_df['divisionRatings_BANK_placeInRating_Compare'] = compare_df.apply(
            lambda row: rang_compare(row, 'divisionRatings_BANK_placeInRating', 'divisionRatings_BANK_placeInRating', STATUS_BANK_PLACE), axis=1)
        compare_df['divisionRatings_TB_placeInRating_Compare'] = compare_df.apply(
            lambda row: rang_compare(row, 'divisionRatings_TB_placeInRating', 'divisionRatings_TB_placeInRating', STATUS_TB_PLACE), axis=1)
        compare_df['divisionRatings_GOSB_placeInRating_Compare'] = compare_df.apply(
            lambda row: rang_compare(row, 'divisionRatings_GOSB_placeInRating', 'divisionRatings_GOSB_placeInRating', STATUS_GOSB_PLACE), axis=1)

        final_cols = compare_keys + [
            'New_Remove', 'indicatorValue_Compare',
            'divisionRatings_BANK_placeInRating_Compare',
            'divisionRatings_TB_placeInRating_Compare',
            'divisionRatings_GOSB_placeInRating_Compare'
        ] + ['BEFORE_' + c for c in compare_fields] + ['AFTER_' + c for c in compare_fields]
        compare_df = compare_df.reindex(columns=final_cols)
        logging.info(f"[OK] Compare sheet готов: строк {len(compare_df)}, колонок {len(compare_df.columns)}")
        return compare_df, sheet_name
    except Exception as ex:
        logging.error(f"Ошибка в make_compare_sheet: {ex}")
        return pd.DataFrame(), sheet_name

def add_smart_table(writer, df, sheet_name, table_name):
    worksheet = writer.sheets[sheet_name]
    (nrows, ncols) = df.shape
    if nrows == 0:
        logging.info(f"{sheet_name}: пустая таблица — не создаём Smart Table.")
        return
    col_letters = []
    for i in range(ncols):
        first = i // 26
        second = i % 26
        if first == 0:
            col_letters.append(chr(65 + second))
        else:
            col_letters.append(chr(65 + first - 1) + chr(65 + second))
    last_col = col_letters[-1]
    excel_range = f"A1:{last_col}{nrows+1}"
    worksheet.add_table(excel_range, {
        'name': table_name,
        'columns': [{'header': col} for col in df.columns],
        'style': 'TableStyleMedium9',
    })
    logging.info(f"{sheet_name}: Smart Table '{table_name}' оформлена.")

def main(
    source_dir: str,
    target_dir: str,
    log_dir: str,
    log_basename: str,
    before_filename: str,
    after_filename: str,
    result_excel: str
):
    logger = setup_logger(log_dir, log_basename)
    try:
        logging.info(f"[START] Экспорт лидеров по турнирам...")
        before_path = os.path.join(source_dir, before_filename)
        after_path = os.path.join(source_dir, after_filename)
        now = datetime.now()
        ts = now.strftime("%Y%m%d_%H%M%S")
        sheet_before = f"BEFORE_{ts}"
        sheet_after = f"AFTER_{ts}"
        sheet_compare = f"COMPARE_{ts}"

        df_before = process_json_file(before_path)
        logging.info(f"BEFORE: строк {len(df_before)}, колонок {len(df_before.columns)}")
        df_after = process_json_file(after_path)
        logging.info(f"AFTER: строк {len(df_after)}, колонок {len(df_after.columns)}")

        log_data_stats(df_before, "BEFORE")
        log_data_stats(df_after, "AFTER")

        all_cols = PRIORITY_COLS.copy()
        all_cols += [c for c in set(df_before.columns).union(df_after.columns) if c not in all_cols]
        log_columns(df_before, "BEFORE", all_cols)
        log_columns(df_after, "AFTER", all_cols)
        df_before = align_and_sort(df_before, all_cols)
        df_after = align_and_sort(df_after, all_cols)
        logging.info(f"Итоговое количество столбцов (финальная структура): {len(df_before.columns)}")

        compare_df, sheet_compare = make_compare_sheet(
            df_before, df_after, COMPARE_KEYS, COMPARE_FIELDS, sheet_compare
        )

        base, ext = os.path.splitext(result_excel)
        result_excel_ts = f"{base}_{ts}{ext}"
        out_excel = os.path.join(target_dir, result_excel_ts)

        with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
            df_before.to_excel(writer, index=False, sheet_name=sheet_before)
            add_smart_table(writer, df_before, sheet_before, "SMART_" + sheet_before)
            logging.info(f"BEFORE экспортирован: {len(df_before)} строк, {len(df_before.columns)} колонок.")

            df_after.to_excel(writer, index=False, sheet_name=sheet_after)
            add_smart_table(writer, df_after, sheet_after, "SMART_" + sheet_after)
            logging.info(f"AFTER экспортирован: {len(df_after)} строк, {len(df_after.columns)} колонок.")

            compare_df.to_excel(writer, index=False, sheet_name=sheet_compare)
            add_smart_table(writer, compare_df, sheet_compare, "SMART_" + sheet_compare)
            logging.info(f"COMPARE экспортирован: {len(compare_df)} строк, {len(compare_df.columns)} колонок.")

        logging.info(f"[SUCCESS] Файл {out_excel} создан.")
    except Exception as ex:
        logging.error(f"ГЛАВНАЯ ОШИБКА: {ex}")

if __name__ == "__main__":
    main(
        source_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON",
        target_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX",
        log_dir=LOG_DIR,
        log_basename=LOG_BASENAME,
        before_filename="LFA_5.json",
        after_filename="LFA_6.json",
        result_excel="LFA_COMPARE.xlsx"
    )
