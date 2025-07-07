from config import *
from logging_utils import setup_logger
from json_loader import process_json_file
from comparer import make_compare_sheet
import pandas as pd
from datetime import datetime

def add_smart_table(writer, df, sheet_name, table_name):
    worksheet = writer.sheets[sheet_name]
    (nrows, ncols) = df.shape
    if nrows == 0:
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

def apply_status_colors(writer, df, sheet_name, status_color_map, status_columns):
    worksheet = writer.sheets[sheet_name]
    for col_name in status_columns:
        if col_name not in df.columns:
            continue
        col_idx = df.columns.get_loc(col_name)
        col_letter = chr(65 + col_idx) if col_idx < 26 else chr(65 + col_idx // 26 - 1) + chr(65 + col_idx % 26)
        cell_range = f"{col_letter}2:{col_letter}{len(df)+1}"
        for status, color in status_color_map.items():
            worksheet.conditional_format(cell_range, {
                'type':     'text',
                'criteria': 'containing',
                'value':    status,
                'format':   writer.book.add_format({'bg_color': color})
            })

def main():
    logger = setup_logger()
    before_path = os.path.join(SOURCE_DIR, BEFORE_FILENAME)
    after_path = os.path.join(SOURCE_DIR, AFTER_FILENAME)
    now = datetime.now()
    ts = now.strftime("%Y%m%d_%H%M%S")
    sheet_before = f"BEFORE_{ts}"
    sheet_after = f"AFTER_{ts}"
    sheet_compare = f"COMPARE_{ts}"

    df_before = process_json_file(before_path)
    df_after = process_json_file(after_path)

    all_cols = PRIORITY_COLS.copy()
    all_cols += [c for c in set(df_before.columns).union(df_after.columns) if c not in all_cols]
    df_before = df_before.reindex(columns=all_cols)
    df_after = df_after.reindex(columns=all_cols)

    compare_df, sheet_compare = make_compare_sheet(
        df_before, df_after, sheet_compare
    )

    base, ext = os.path.splitext(RESULT_EXCEL)
    result_excel_ts = f"{base}_{ts}{ext}"
    out_excel = os.path.join(TARGET_DIR, result_excel_ts)

    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
        df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        add_smart_table(writer, df_before, sheet_before, "SMART_" + sheet_before)
        df_after.to_excel(writer, index=False, sheet_name=sheet_after)
        add_smart_table(writer, df_after, sheet_after, "SMART_" + sheet_after)
        compare_df.to_excel(writer, index=False, sheet_name=sheet_compare)
        add_smart_table(writer, compare_df, sheet_compare, "SMART_" + sheet_compare)
        apply_status_colors(
            writer,
            compare
