import os
import json
import pandas as pd
from datetime import datetime
import re

# === ОПРЕДЕЛЕНИЕ ПРИОРИТЕТА КОЛОНОК ДЛЯ ОСНОВНЫХ ЛИСТОВ ===
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

# === СПИСОК КЛЮЧЕЙ ДЛЯ COMPARE-ЛИСТА (только эти поля определяют уникальность) ===
COMPARE_KEYS = [
    'tournamentId',
    'employeeNumber',
    'lastName',
    'firstName',
]

# === ПОЛЯ, КОТОРЫЕ СРАВНИВАЮТСЯ НА COMPARE-ЛИСТЕ (ТОЛЬКО ИХ) ===
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

# === СПИСКИ ДЛЯ ЧИСЛОВОГО ПРЕОБРАЗОВАНИЯ ===
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

# === КОНФИГУРАЦИЯ МЕТОК ДЛЯ СРАВНЕНИЯ (легко менять без правки логики) ===
COMPARE_STATUS = {
    "both":        "No Change",
    "before_only": "Remove",
    "after_only":  "New",
    "val_add":     "New ADD",
    "val_remove":  "Remove FROM",
    "val_nochange": "No Change",
    "val_down":     "Change DOWN",
    "val_up":       "Change UP"
}

# === ФУНКЦИИ ДЛЯ ПРЕОБРАЗОВАНИЯ ЧИСЛОВЫХ ПОЛЕЙ ===
def parse_float(val):
    """
    Преобразует значение к float с 3 знаками после запятой,
    удаляя все возможные пробелы и разделители.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return round(float(val), 3)
    s = str(val)
    s = re.sub(r'[\s\u00A0\u2009]', '', s)   # Удаляем любые пробелы, включая неразрывные и тонкие
    s = re.sub(r"[^\d.,\-]", "", s)          # Оставляем только цифры, точки, запятые, минус
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
    """
    Преобразует значение к int (целое число), убирая пробелы и нечисловые символы.
    """
    if val is None:
        return None
    if isinstance(val, int):
        return val
    s = str(val)
    s = re.sub(r'[\s\u00A0\u2009]', '', s)  # Удаляем любые пробелы
    s = re.sub(r"[^\d\-]", "", s)           # Оставляем только цифры и минус
    try:
        return int(s)
    except Exception:
        return None

# === ПЛОСКАЯ РАЗВЕРТКА ОДНОГО ЛИДЕРА ===
def flatten_leader(leader, tournament_id, source_file):
    """
    Преобразует структуру одного лидера в плоский словарь.
    Исключает photoData, разворачивает вложенные divisionRatings.
    Числовые поля преобразует согласно правилам.
    """
    row = {
        'SourceFile': source_file,
        'tournamentId': tournament_id
    }
    # Переносим все поля кроме специальных
    for k, v in leader.items():
        if k in ("divisionRatings", "photoData"):
            continue
        if k in FLOAT_FIELDS:
            row[k] = parse_float(v)
        else:
            row[k] = v
    # Разворачиваем вложенные divisionRatings с группировкой
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
    # Гарантируем наличие всех float-полей
    for f in FLOAT_FIELDS:
        if f not in row:
            row[f] = None
    return row

# === ЗАГРУЗКА И РАЗВОРАЧИВАНИЕ ВСЕГО JSON ФАЙЛА В СПИСОК ЛИДЕРОВ ===
def process_json_file(filepath):
    """
    Загружает файл, разворачивает всех лидеров.
    """
    filename = os.path.basename(filepath)
    with open(filepath, 'r', encoding='utf-8') as f:
        js = json.load(f)
    all_rows = []
    # Поддержка структуры файла (или списка, или словаря)
    for block in js.values():
        blocks = block if isinstance(block, list) else [block]
        for subblock in blocks:
            tournament = subblock.get("body", {}).get("tournament", {})
            tournament_id = tournament.get("tournamentId")
            for leader in tournament.get("leaders", []):
                all_rows.append(flatten_leader(leader, tournament_id, filename))
    return pd.DataFrame(all_rows)

# === ВЫРАВНИВАНИЕ ТАБЛИЦ ПОД ЖЕЛАЕМЫЙ ПОРЯДОК КОЛОНОК ===
def align_and_sort(df, all_columns):
    """
    Упорядочивает датафрейм согласно all_columns, добавляет пустые колонки если их не было.
    """
    for col in all_columns:
        if col not in df.columns:
            df[col] = None
    rest = [c for c in df.columns if c not in all_columns]
    return df[all_columns + rest]

# === ЛОГИРОВАНИЕ КОЛОНОК И ИЗМЕНЕНИЙ ===
def log_columns(df, label, final_columns):
    print(f"[LOG] {label} — исходно {len(df.columns)} колонок: {sorted(list(df.columns))}")
    print(f"[LOG] {label} — сохраняется {len(final_columns)} колонок: {final_columns}")
    added = [col for col in final_columns if col not in df.columns]
    removed = [col for col in df.columns if col not in final_columns]
    if added:
        print(f"[LOG] {label} — добавлены новые пустые колонки: {added}")
    if removed:
        print(f"[LOG] {label} — эти колонки будут удалены: {removed}")

# === ЛОГИРОВАНИЕ ПО КОЛИЧЕСТВУ ЛИДЕРОВ ПО ТУРНИРАМ ===
def log_data_stats(df, label):
    print(f"[STAT] {label}: найдено уникальных tournamentId = {df['tournamentId'].nunique()}")
    for tid, group in df.groupby('tournamentId'):
        print(f"[STAT]   tournamentId = {tid}: людей = {len(group)}")

# === ФОРМИРОВАНИЕ COMPARE-ЛИСТА С ПОЛЯМИ СРАВНЕНИЯ ===
def make_compare_sheet(df_before, df_after, compare_keys, compare_fields, sheet_name):
    """
    Формирует таблицу сравнения:
      - Уникальные ключи: только COMPARE_KEYS
      - Для каждого поля COMPARE_FIELDS — столбцы BEFORE_ и AFTER_
      - Добавляет столбцы сравнения наличия и изменений по indicatorValue
    """
    join_keys = compare_keys

    # 1. Выбираем только уникальные строки по ключам
    before_uniq = df_before.drop_duplicates(subset=join_keys, keep='last')
    after_uniq  = df_after.drop_duplicates(subset=join_keys, keep='last')

    # 2. Полный список всех уникальных ключей из обоих файлов
    all_keys = pd.concat([before_uniq[join_keys], after_uniq[join_keys]]).drop_duplicates()

    # 3. Индексация по ключам
    before_uniq = before_uniq.set_index(join_keys)
    after_uniq  = after_uniq.set_index(join_keys)

    # 4. Оставляем только нужные поля для сравнения
    before_uniq = before_uniq[compare_fields] if len(before_uniq) else pd.DataFrame(columns=compare_fields)
    after_uniq  = after_uniq[compare_fields]  if len(after_uniq) else pd.DataFrame(columns=compare_fields)

    # 5. Добавляем префиксы к столбцам
    before_uniq = before_uniq.add_prefix('BEFORE_')
    after_uniq  = after_uniq.add_prefix('AFTER_')

    # 6. Джойним оба датафрейма по всем ключам
    compare_df = all_keys.set_index(join_keys) \
        .join(before_uniq, how='left') \
        .join(after_uniq, how='left') \
        .reset_index()

    # 7. Итоговый порядок: ключи, сравнения, все BEFORE_, все AFTER_
    cols_order = compare_keys + \
        ['New_Remove', 'indicatorValue_Compare'] + \
        ['BEFORE_' + c for c in compare_fields] + \
        ['AFTER_' + c for c in compare_fields]
    compare_df = compare_df.reindex(columns=cols_order)

    # === Сравнение наличия записи: поле New_Remove ===
    def new_remove_row(row):
        before_exist = not pd.isnull(row['BEFORE_indicatorValue']) or not pd.isnull(row['BEFORE_SourceFile'])
        after_exist  = not pd.isnull(row['AFTER_indicatorValue'])  or not pd.isnull(row['AFTER_SourceFile'])
        if before_exist and after_exist:
            return COMPARE_STATUS['both']
        elif before_exist:
            return COMPARE_STATUS['before_only']
        elif after_exist:
            return COMPARE_STATUS['after_only']
        else:
            return ""
    compare_df['New_Remove'] = compare_df.apply(new_remove_row, axis=1)

    # === Сравнение значения indicatorValue: поле indicatorValue_Compare ===
    def value_compare(row):
        before = row.get('BEFORE_indicatorValue', None)
        after  = row.get('AFTER_indicatorValue', None)
        if pd.isnull(before) and not pd.isnull(after):
            return COMPARE_STATUS['val_add']
        if not pd.isnull(before) and pd.isnull(after):
            return COMPARE_STATUS['val_remove']
        if pd.isnull(before) and pd.isnull(after):
            return ""
        # оба есть
        if before == after:
            return COMPARE_STATUS['val_nochange']
        elif before > after:
            return COMPARE_STATUS['val_down']
        else:
            return COMPARE_STATUS['val_up']
    compare_df['indicatorValue_Compare'] = compare_df.apply(value_compare, axis=1)

    print(f"[OK] Compare sheet готов: строк {len(compare_df)}, колонок {len(compare_df.columns)}")
    return compare_df, sheet_name

# === СОЗДАНИЕ SMART-TABLE В EXCEL ===
def add_smart_table(writer, df, sheet_name, table_name):
    """
    Добавляет к листу Excel "умную таблицу" (smart table) с автофильтрами.
    """
    worksheet = writer.sheets[sheet_name]
    (nrows, ncols) = df.shape
    if nrows == 0:
        print(f"[SMART] {sheet_name}: пустая таблица — не создаём Smart Table.")
        return
    # Определяем Excel диапазон A1:...
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
        'style': 'TableStyleMedium9',  # Можно выбрать стиль
    })
    print(f"[SMART] {sheet_name}: Smart Table '{table_name}' оформлена.")

# === ГЛАВНАЯ ТОЧКА ВХОДА ===
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

    # === Загружаем и разворачиваем оба файла ===
    df_before = process_json_file(before_path)
    print(f"[INFO] BEFORE: строк {len(df_before)}, колонок {len(df_before.columns)}")
    df_after = process_json_file(after_path)
    print(f"[INFO] AFTER: строк {len(df_after)}, колонок {len(df_after.columns)}")

    # === Логируем количество записей по турнирам ===
    log_data_stats(df_before, "BEFORE")
    log_data_stats(df_after, "AFTER")

    # === Выравниваем структуру таблиц ===
    all_cols = PRIORITY_COLS.copy()
    all_cols += [c for c in set(df_before.columns).union(df_after.columns) if c not in all_cols]
    log_columns(df_before, "BEFORE", all_cols)
    log_columns(df_after, "AFTER", all_cols)
    df_before = align_and_sort(df_before, all_cols)
    df_after = align_and_sort(df_after, all_cols)
    print(f"[INFO] Итоговое количество столбцов (финальная структура): {len(df_before.columns)}")

    # === Формируем COMPARE-таблицу с полями сравнения ===
    compare_df, sheet_compare = make_compare_sheet(
        df_before, df_after, COMPARE_KEYS, COMPARE_FIELDS, sheet_compare
    )

    # === Сохраняем в Excel, создаем smart-tables ===
    base, ext = os.path.splitext(result_excel)
    result_excel_ts = f"{base}_{ts}{ext}"
    out_excel = os.path.join(target_dir, result_excel_ts)

    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
        # BEFORE
        df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        add_smart_table(writer, df_before, sheet_before, "SMART_" + sheet_before)
        print(f"[OK] BEFORE экспортирован: {len(df_before)} строк, {len(df_before.columns)} колонок.")

        # AFTER
        df_after.to_excel(writer, index=False, sheet_name=sheet_after)
        add_smart_table(writer, df_after, sheet_after, "SMART_" + sheet_after)
        print(f"[OK] AFTER экспортирован: {len(df_after)} строк, {len(df_after.columns)} колонок.")

        # COMPARE
        compare_df.to_excel(writer, index=False, sheet_name=sheet_compare)
        add_smart_table(writer, compare_df, sheet_compare, "SMART_" + sheet_compare)
        print(f"[OK] COMPARE экспортирован: {len(compare_df)} строк, {len(compare_df.columns)} колонок.")

    print(f"[SUCCESS] Файл {out_excel} создан.")

# === ЗАПУСК ===
if __name__ == "__main__":
    main(
        source_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON",
        target_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX",
        before_filename="LFA_5.json",
        after_filename="LFA_6.json",
        result_excel="LFA_COMPARE.xlsx"
    )
