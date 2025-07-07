from config import (
    SOURCE_DIR, TARGET_DIR, BEFORE_FILENAME, AFTER_FILENAME,
    RESULT_EXCEL, PRIORITY_COLS, LOG_BASENAME, LOG_DIR,
    STATUS_COLORS_DICT, STATUS_COLOR_COLUMNS, STATUS_RU_DICT, STATUS_RATING_CATEGORY
)
from logging_utils import setup_logger
from json_loader import process_json_file
from comparer import make_compare_sheet
import pandas as pd
import os
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
    dark_bg = {"#383838", "#222222"}
    # Собираем статусы, которым нужен белый шрифт
    statuses_with_white_font = set()
    for status, color in status_color_map.items():
        if color.lower() in dark_bg:
            statuses_with_white_font.add(status)
    for col_name in status_columns:
        if col_name not in df.columns:
            continue
        col_idx = df.columns.get_loc(col_name)
        if col_idx < 26:
            col_letter = chr(65 + col_idx)
        else:
            col_letter = chr(65 + col_idx // 26 - 1) + chr(65 + col_idx % 26)
        cell_range = f"{col_letter}2:{col_letter}{len(df)+1}"
        for status, color in status_color_map.items():
            fmt = writer.book.add_format({'bg_color': color})
            if status in statuses_with_white_font:
                fmt.set_font_color("#FFFFFF")
            worksheet.conditional_format(cell_range, {
                'type':     'text',
                'criteria': 'containing',
                'value':    status,
                'format':   fmt
            })

def log_data_stats(df, label):
    if df.empty:
        import logging
        logging.info(f"[{label}] DataFrame пустой.")
        return
    n_rows = len(df)
    n_cols = len(df.columns)
    tournament_counts = df['tournamentId'].value_counts().to_dict()
    unique_tids = list(df['tournamentId'].unique())
    import logging
    logging.info(f"[{label}] строк: {n_rows}, колонок: {n_cols}")
    logging.info(f"[{label}] tournamentId всего: {len(unique_tids)} -> {unique_tids}")
    for tid in unique_tids:
        count = tournament_counts.get(tid, 0)
        logging.info(f"[{label}] tournamentId={tid}: людей={count}")
    logging.info(f"[{label}] Все поля: {list(df.columns)}")

def log_compare_stats(compare_df):
    import logging
    n_rows = len(compare_df)
    logging.info(f"[COMPARE] Строк всего: {n_rows}")
    for col in STATUS_COLOR_COLUMNS:
        if col in compare_df.columns:
            counts = compare_df[col].value_counts(dropna=False).to_dict()
            logging.info(f"[COMPARE] {col}: {counts}")

def add_status_legend(writer, status_colors, status_ru_dict, status_rating_category, sheet_name="STATUS_LEGEND"):
    """
    Добавляет лист Excel с легендой по статусам.
    """
    rows = []
    for key, eng_status in status_rating_category.items():
        ru, comment = status_ru_dict.get(eng_status, ("", ""))
        color = status_colors.get(eng_status, "#FFFFFF")
        rows.append({
            "Status code": eng_status,
            "Статус (рус)": ru,
            "Excel fill color": color,
            "Комментарий": comment
        })
    legend_df = pd.DataFrame(rows)
    legend_df = legend_df[["Status code", "Статус (рус)", "Excel fill color", "Комментарий"]]
    legend_df.to_excel(writer, index=False, sheet_name=sheet_name)
    worksheet = writer.sheets[sheet_name]
    for i, row in legend_df.iterrows():
        fmt = writer.book.add_format({'bg_color': row["Excel fill color"]})
        # Белый шрифт для темного
        if row["Excel fill color"].lower() == "#383838":
            fmt.set_font_color("#FFFFFF")
        worksheet.write(i+1, 0, row["Status code"])
        worksheet.write(i+1, 1, row["Статус (рус)"])
        worksheet.write(i+1, 2, row["Excel fill color"], fmt)
        worksheet.write(i+1, 3, row["Комментарий"])

def main():
    logger = setup_logger(LOG_DIR, LOG_BASENAME)

    before_path = os.path.join(SOURCE_DIR, BEFORE_FILENAME)
    after_path = os.path.join(SOURCE_DIR, AFTER_FILENAME)
    now = datetime.now()
    ts = now.strftime("%Y%m%d_%H%M%S")
    sheet_before = f"BEFORE_{ts}"
    sheet_after = f"AFTER_{ts}"
    sheet_compare = f"COMPARE_{ts}"

    logger.info(f"[MAIN] Читаем BEFORE: {before_path}")
    rows_before = process_json_file(before_path)
    df_before = pd.DataFrame(rows_before)
    log_data_stats(df_before, "BEFORE")
    logger.info(f"[MAIN] Читаем AFTER: {after_path}")
    rows_after = process_json_file(after_path)
    df_after = pd.DataFrame(rows_after)
    log_data_stats(df_after, "AFTER")

    before_tids = set(df_before['tournamentId'].unique())
    after_tids = set(df_after['tournamentId'].unique())
    added_tids = after_tids - before_tids
    removed_tids = before_tids - after_tids
    common_tids = before_tids & after_tids

    logger.info(f"[MAIN] Турниров в BEFORE: {len(before_tids)}, в AFTER: {len(after_tids)}")
    logger.info(f"[MAIN] Новые турниры (только в AFTER): {len(added_tids)} -> {list(added_tids)}")
    logger.info(f"[MAIN] Удалённые турниры (только в BEFORE): {len(removed_tids)} -> {list(removed_tids)}")
    logger.info(f"[MAIN] Общие турниры: {len(common_tids)} -> {list(common_tids)}")

    all_cols = PRIORITY_COLS.copy()
    all_cols += [c for c in set(df_before.columns).union(df_after.columns) if c not in all_cols]
    df_before = df_before.reindex(columns=all_cols)
    df_after = df_after.reindex(columns=all_cols)

    logger.info(f"[MAIN] Формируем COMPARE")
    compare_df, sheet_compare = make_compare_sheet(
        df_before, df_after, sheet_compare
    )

    log_compare_stats(compare_df)

    base, ext = os.path.splitext(RESULT_EXCEL)
    result_excel_ts = f"{base}_{ts}{ext}"
    out_excel = os.path.join(TARGET_DIR, result_excel_ts)

    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
        logger.info(f"[MAIN] Экспортируем BEFORE лист {sheet_before}")
        df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        add_smart_table(writer, df_before, sheet_before, "SMART_" + sheet_before)
        logger.info(f"[MAIN] Экспортируем AFTER лист {sheet_after}")
        df_after.to_excel(writer, index=False, sheet_name=sheet_after)
        add_smart_table(writer, df_after, sheet_after, "SMART_" + sheet_after)
        logger.info(f"[MAIN] Экспортируем COMPARE лист {sheet_compare}")
        compare_df.to_excel(writer, index=False, sheet_name=sheet_compare)
        add_smart_table(writer, compare_df, sheet_compare, "SMART_" + sheet_compare)
        apply_status_colors(
            writer,
            compare_df,
            sheet_compare,
            STATUS_COLORS_DICT,
            STATUS_COLOR_COLUMNS
        )
        add_status_legend(writer, STATUS_COLORS_DICT, STATUS_RU_DICT, STATUS_RATING_CATEGORY, sheet_name="STATUS_LEGEND")
        logger.info(f"[MAIN] Все данные выгружены в файл: {out_excel}")

if __name__ == "__main__":
    main()
