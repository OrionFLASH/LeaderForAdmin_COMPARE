import os
import json
import pandas as pd
from datetime import datetime
import re

def parse_float(val):
    """Преобразует строку вида '3 168,89' или '3,168.890' в число с 3 знаками после запятой"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(f"{val:.3f}")
    # Удаляем все кроме цифр, запятой и точки
    s = str(val)
    s = re.sub(r"[^\d.,\-]", "", s)
    # Если оба разделителя, оставляем последний как основной
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

def flatten_leader(leader, tournament_id, source_file):
    """Возвращает плоскую строку по одному участнику + tournamentId + имя файла."""
    row = {
        'SourceFile': source_file,
        'tournamentId': tournament_id
    }
    for k, v in leader.items():
        if k not in ("divisionRatings", "indicatorValue", "successValue"):
            row[k] = v
    # Форматируем indicatorValue и successValue:
    row['indicatorValue'] = parse_float(leader.get('indicatorValue'))
    row['successValue'] = parse_float(leader.get('successValue'))
    # Разворачиваем divisionRatings:
    if "divisionRatings" in leader and leader["divisionRatings"]:
        for idx, div in enumerate(leader["divisionRatings"], 1):
            group = div.get("groupCode", f"DIV{idx:02d}")
            for field, value in div.items():
                if field == "groupCode":
                    continue
                colname = f"divisionRatings_{group}_{field}"
                row[colname] = value
    return row

def process_json_file(filepath):
    """Парсит leaders в плоский DataFrame, добавляет имя файла"""
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

def align_dataframes(df1, df2):
    """Обеспечивает одинаковую структуру обеих таблиц (одинаковый набор колонок)"""
    all_columns = sorted(set(df1.columns).union(df2.columns))
    return df1.reindex(columns=all_columns), df2.reindex(columns=all_columns)

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

    # Выравниваем столбцы по максимальному объединенному набору
    df_before, df_after = align_dataframes(df_before, df_after)
    print(f"[INFO] Итоговое количество столбцов (в обеих таблицах): {len(df_before.columns)}")

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
