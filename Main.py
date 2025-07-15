import os
import json
import pandas as pd
import re
import logging
# import sys
from datetime import datetime
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# --- Параметры логирования ---
# LOG_LEVEL определяет глубину вывода в консоль (INFO или DEBUG)
LOG_LEVEL = logging.INFO

# --- Цветовые настройки и служебные тексты ---
# Набор цветов с тёмным фоном, требующих белый шрифт
APPLY_DARK_BG_COLORS = {"383838", "222222"}
# Аналогичный набор для листа легенды
LEGEND_DARK_BG_COLORS = {"383838", "222222", "000000"}
# Цвет, применяемый по умолчанию, если статус неизвестен
DEFAULT_STATUS_COLOR = "#FFFFFF"

# Текстовые обозначения этапов
FINAL_START_MESSAGE = "=== [FINAL] Построение итоговой сводной таблицы ==="
STATUS_LEGEND_SHEET = "STATUS_LEGEND"

# Шаблон итоговой строки
SUMMARY_TEMPLATE = (
    "[SUMMARY] турниров: {tourn}; сотрудников: {emps}; "
    "изменений: {changes}; load_before: {t1:.2f}s; "
    "load_after: {t2:.2f}s; compare: {t3:.2f}s; final: {t4:.2f}s; "
    "export: {t5:.2f}s; total: {tt:.2f}s"
)

# === Константы путей и имён файлов ===
# Здесь задаются пути к папкам с исходниками, результатами и логами,
# а также имена входных/выходных файлов. При необходимости их можно
# поменять под свои каталоги.
SOURCE_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON"
TARGET_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX"
LOG_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//LOGS"
LOG_BASENAME = "LOG1"
BEFORE_FILENAME = "leadersForAdmin_ALL_20250708-140508.json"
AFTER_FILENAME = "leadersForAdmin_ALL_20250714-093911.json"
RESULT_EXCEL = "LFA_COMPARE.xlsx"

# --- Список турниров, которые будут включены в анализ ---
# Если список пустой, сравниваются все турниры из исходных файлов.
ALLOWED_TOURNAMENT_IDS = [
        "t_01_2025-0_08-2_5_2021", "t_01_2025-2_08-2_6_2031", "TOURNAMENT_43_2025_Y01", "TOURNAMENT_66_2024_P02", "TOURNAMENT_67_2024_Y01",
        "t_01_2025-0_12-1_1_1001", "t_01_2025-0_12-1_1_1002", "t_01_2025-0_10-1_1_1001", "t_01_2025-0_10-1_2_1001", "t_01_2025-0_10-1_3_1001", "t_01_2025-0_10-1_4_1001",
        "t_01_2025-1_09-1_1_3061", "t_01_2025-2_09-1_1_3071", "t_01_2025-0_13-1_1_1001", "t_01_2025-0_13-1_2_1001", "t_01_2025-0_13-1_3_1001", "t_01_2025-1_14-1_1_3061",
        "t_01_2025-2_14-1_1_3071", "t_01_2025-1_16-1_1_2021", "t_01_2025-2_16-1_1_2031", "t_01_2025-0_08-2_3_2021", "t_01_2025-0_08-2_3_2031", "t_01_2025-1_05-1_1_3061",
        "t_01_2025-2_05-1_1_3071", "t_01_2025-1_08-1_1_3061", "t_01_2025-2_08-1_1_3071", "t_01_2025-0_08-6_1_4001", "t_01_2025-1_15-1_2_2021", "t_01_2025-2_15-1_2_2031",
        "t_01_2025-1_01-6_4001", "t_01_2025-0_18-7_1_1001", "t_01_2025-0_18-7_2_1001", "t_01_2025-0_18-7_3_1001", "t_01_2025-0_18-7_4_1001", "t_01_2025-1_07-5_1_4001",
        "t_01_2025-1_07-5_2_4001", "t_01_2025-1_07-5_3_4001", "t_01_2025-1_01-7_1_4001", "t_01_2025-2_01-7_1_4001", "t_01_2025-1_07-5_5_4001", "t_01_2025-1_07-5_6_4001",
        "t_01_2025-1_07-5_7_4001", "t_01_2025-1_07-5_8_4001", "t_01_2025-1_07-5_9_4001", "t_01_2025-1_07-5_10_4001", "t_01_2025-1_02-1_2_4001", "t_01_2025-0_10-3_1_1001",
        "t_01_2025-0_10-3_2_1001", "t_01_2025-0_10-3_3_1001", "t_01_2025-1_02-3_1t_3051", "t_01_2025-1_02-3_1t_3061", "t_01_2025-1_05-1_2t_3031", "t_01_2025-1_05-1_2t_3041",
        "t_01_2025-1_05-1_2t_3051", "t_01_2025-1_05-1_2t_3061", "t_01_2025-0_18-7_5_4001", "t_01_2025-1_17-2_1_4001", "t_01_2025-1_17-2_2_4001", "t_01_2025-1_17-2_3_4001",
        "t_01_2025-1_17-2_4_4001", "t_01_2025-0_13-1_4_1001", "t_01_2025-0_13-1_5_1001", "t_01_2025-0_13-1_6_1001", "t_01_2025-1_04-1_1_4001"

    # Если оставить пустым, то анализируются все турниры.
]

# --- Статусы, при которых считаем, что изменений не произошло ---
# Если встречается один из этих кодов, строка считается без изменений
NOCHANGE_STATUSES = [
    "", "No Change", "STAYED_OUT", "PRIZE_UNCHANGED", "Remove", "Remove FROM",
    "Rang BANK REMOVE", "Rang TB REMOVE", "Rang GOSB REMOVE",
    "Rang BANK NO CHANGE", "Rang TB NO CHANGE", "Rang GOSB NO CHANGE",
    "Rang NO CHANGE", "NO_RANK",
]

# --- Основные колонки в исходных данных ---
# Эти поля всегда присутствуют и выводятся в итоговую таблицу в первую очередь
PRIORITY_COLS = [
    'SourceFile', 'tournamentId', 'employeeNumber', 'lastName', 'firstName',
    'terDivisionName', 'divisionRatings_BANK_groupId', 'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId', 'employeeStatus', 'businessBlock',
    'successValue', 'indicatorValue', 'divisionRatings_BANK_placeInRating',
    'divisionRatings_TB_placeInRating', 'divisionRatings_GOSB_placeInRating',
    'divisionRatings_BANK_ratingCategoryName', 'divisionRatings_TB_ratingCategoryName',
    'divisionRatings_GOSB_ratingCategoryName',
]
# Ключевые поля, по которым склеиваются данные "до" и "после"
COMPARE_KEYS = [
    'tournamentId',
    'employeeNumber',
    'lastName',
    'firstName',
]
# Поля, которые сравниваются между выгрузками
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
# Поля, которые должны быть приведены к типу int
INT_FIELDS = [
    'divisionRatings_BANK_groupId',
    'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId',
    'divisionRatings_BANK_placeInRating',
    'divisionRatings_TB_placeInRating',
    'divisionRatings_GOSB_placeInRating',
]
# Поля, которые хранят вещественные значения
FLOAT_FIELDS = [
    'indicatorValue',
    'successValue',
]

# --- Все цвета статусов здесь ---
STATUS_COLORS_DICT = {
    'No Change':        '#BFBFBF',  # Серый
    'Rang BANK NO CHANGE': '#BFBFBF',
    'Rang TB NO CHANGE':   '#BFBFBF',
    'Rang GOSB NO CHANGE': '#BFBFBF',
    'Change UP':        '#C6EFCE',  # Светло-зелёный
    'Rang BANK UP':     '#C6EFCE',
    'Rang TB UP':       '#C6EFCE',
    'Rang GOSB UP':     '#C6EFCE',
    'Change DOWN':      '#FFC7CE',  # Светло-красный
    'Rang BANK DOWN':   '#FFC7CE',
    'Rang TB DOWN':     '#FFC7CE',
    'Rang GOSB DOWN':   '#FFC7CE',
    'New ADD':          '#E2EFDA',  # Бледно-зелёный
    'Rang BANK NEW':    '#E2EFDA',
    'Rang TB NEW':      '#E2EFDA',
    'Rang GOSB NEW':    '#E2EFDA',
    'Remove FROM':      '#383838',  # Темно-серый, ещё темнее
    'Rang BANK REMOVE': '#383838',
    'Rang TB REMOVE':   '#383838',
    'Rang GOSB REMOVE': '#383838',
    'Remove':           '#383838',
    'New':              '#E2EFDA',
    "NO_RANK":          "#EDEDED",  # Светло-серый цвет (можно взять любой RGB/HEX)
    "CONT":             "#C9DAF8",    # Светло-голубой или другой на ваш вкус
    "Not_used":         "#F5F5F5", # Как было

    # Для призовых (категорий)
    "ENTERED_PRIZE":    '#00B050',   # Зеленый (попал в призёры)
    "STAYED_OUT":       '#BFBFBF',   # Серый (остался вне призёров)
    "DROPPED_OUT_PRIZE":'#FF0000',   # Красный (выбыл из призёров)
    "LOST_VIEW":        '#383838',   # Темно-серый (пропал из вида)
    "PRIZE_UNCHANGED":  '#C6EFCE',   # Светло-зелёный (призёр без изменений)
    "PRIZE_UP":         '#00B050',   # Зеленый (улучшил место)
    "PRIZE_DOWN":       '#FFC7CE',   # Светло-красный (понизился)
}

# Какие колонки раскрашивать (для передачи в apply_status_colors)
STATUS_COLOR_COLUMNS = [
    'indicatorValue_Compare',
    'divisionRatings_BANK_placeInRating_Compare',
    'divisionRatings_TB_placeInRating_Compare',
    'divisionRatings_GOSB_placeInRating_Compare',
    'divisionRatings_BANK_ratingCategoryName_Compare',
    'divisionRatings_TB_ratingCategoryName_Compare',
    'divisionRatings_GOSB_ratingCategoryName_Compare'
]

# --- Справочник по статусам (Excel-код: (рус, комментарий)) ---
STATUS_RU_DICT = {
    "ENTERED_PRIZE":      ("ПОПАЛ В ПРИЗЁРЫ", "Был вне призёров, стал призёром. Это хорошо."),
    "STAYED_OUT":         ("ОСТАЛСЯ ВНЕ ПРИЗЁРОВ", "Не был призёром и не стал. Без изменений, но не лучший результат."),
    "DROPPED_OUT_PRIZE":  ("ВЫБЫЛ ИЗ ПРИЗЁРОВ", "Был призёром, стал вне призёров. Это плохо."),
    "LOST_VIEW":          ("ПРОПАЛ ИЗ ВИДА", "Был призёром, теперь отсутствует в итоговом файле."),
    "PRIZE_UNCHANGED":    ("ПРИЗЁР БЕЗ ИЗМЕНЕНИЙ", "Был призёром, остался на том же месте."),
    "PRIZE_UP":           ("УЛУЧШИЛ ПРИЗОВОЕ МЕСТО", "Был призёром и стал лучше (например, с бронзы на золото)."),
    "PRIZE_DOWN":         ("ПОНИЗИЛСЯ В РЕЙТИНГЕ ПРИЗЁРОВ", "Был призёром и опустился на худшее призовое место."),
}

# Статусы для сравнения (логика и сокращения)
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
    "val_down":     "Rang BANK DOWN",
    "val_norank":   "NO_RANK"  # добавлено!
}
STATUS_TB_PLACE = {
    "val_add":      "Rang TB NEW",
    "val_remove":   "Rang TB REMOVE",
    "val_nochange": "Rang TB NO CHANGE",
    "val_up":       "Rang TB UP",
    "val_down":     "Rang TB DOWN",
    "val_norank":   "NO_RANK"  # добавлено!
}
STATUS_GOSB_PLACE = {
    "val_add":      "Rang GOSB NEW",
    "val_remove":   "Rang GOSB REMOVE",
    "val_nochange": "Rang GOSB NO CHANGE",
    "val_up":       "Rang GOSB UP",
    "val_down":     "Rang GOSB DOWN",
    "val_norank":   "NO_RANK"  # добавлено!
}
CATEGORY_RANK_MAP = {
    "Вы в лидерах": 1,
    "Серебро": 2,
    "Бронза": 3,
    "Нужно поднажать": 4,
    "": 4,
    None: 4
}
STATUS_RATING_CATEGORY = {
    "in2prize": "ENTERED_PRIZE",
    "stay_out": "STAYED_OUT",
    "from2out": "DROPPED_OUT_PRIZE",
    "lost":     "LOST_VIEW",
    "same":     "PRIZE_UNCHANGED",
    "up":       "PRIZE_UP",
    "down":     "PRIZE_DOWN",
}


def setup_logger(log_dir, basename):
    """
    Создаёт логгер, который пишет и в файл (append, по дате), и в консоль.
    log_dir: путь к папке для логов
    basename: имя (без даты) для лог-файла
    """
    now = datetime.now()
    day_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H:%M:%S")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, f"{basename}_{day_str}.log")
    # Добавляем маркер начала новой сессии
    with open(log_path, "a", encoding="utf-8") as logf:
        logf.write(f"\n-------- NEW LOG START AT {day_str} ({time_str}) -------\n")
    # Стандартное подключение логгера
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    if logger.hasHandlers():
        logger.handlers.clear()
    fh = logging.FileHandler(log_path, encoding='utf-8', mode='a')
    fh.setLevel(logging.DEBUG)
    ch = logging.StreamHandler()
    ch.setLevel(LOG_LEVEL)
    fmt = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s', "%Y-%m-%d %H:%M:%S")
    fh.setFormatter(fmt)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    logging.info(f"Лог-файл активен (append): {log_path}")
    return logger

def log_data_stats(df, label):
    if df.empty:
        logging.info(f"[{label}] DataFrame пустой.")
        return
    n_rows = len(df)
    n_cols = len(df.columns)
    tournament_counts = df['tournamentId'].value_counts().to_dict()
    unique_tids = list(df['tournamentId'].unique())
    logging.info(f"[{label}] строк: {n_rows}, колонок: {n_cols}")
    logging.info(f"[{label}] tournamentId всего: {len(unique_tids)} -> {unique_tids}")
    for tid in unique_tids:
        count = tournament_counts.get(tid, 0)
        logging.info(f"[{label}] tournamentId={tid}: людей={count}")
    logging.info(f"[{label}] Все поля: {list(df.columns)}")

def log_compare_stats(compare_df):
    """Выводит сводную статистику по статусным колонкам таблицы сравнения."""
    n_rows = len(compare_df)
    logging.info(f"[COMPARE] Строк всего: {n_rows}")
    for col in STATUS_COLOR_COLUMNS:
        if col in compare_df.columns:
            counts = compare_df[col].value_counts(dropna=False).to_dict()
            logging.info(f"[COMPARE] {col}: {counts}")


def parse_float(val, context=None):
    """Преобразует значение в float, если возможно."""
    try:
        if val is None or (isinstance(val, str) and val.strip().lower() in {'', 'none', 'null'}):
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
        logging.error(
            f"[parse_float] Ошибка преобразования '{val}' в float: {ex} | Context: {context}"
        )
        return None

def parse_int(val, context=None):
    """Преобразует значение в int, если возможно."""
    try:
        if val is None or (isinstance(val, str) and val.strip().lower() in {'', 'none', 'null'}):
            return None
        if isinstance(val, int):
            return val
        s = str(val)
        s = re.sub(r'[\s\u00A0\u2009]', '', s)
        s = re.sub(r"[^\d\-]", "", s)
        return int(s)
    except Exception as ex:
        logging.error(
            f"[parse_int] Ошибка преобразования '{val}' в int: {ex} | Context: {context}"
        )
        return None

def flatten_leader(leader, tournament_id, source_file):
    """Разворачивает запись лидера в плоскую структуру для DataFrame."""
    context = f"файл={source_file}, турнир={tournament_id}, employee={leader.get('employeeNumber', 'N/A')}"
    row = {
        'SourceFile': source_file,
        'tournamentId': tournament_id
    }
    for k, v in leader.items():
        if k in ("divisionRatings", "photoData"):
            continue
        if k in FLOAT_FIELDS:
            row[k] = parse_float(v, context)
        else:
            row[k] = v
    # divisionRatings
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
                        row[colname] = parse_int(value, context)
                    elif colname in FLOAT_FIELDS:
                        row[colname] = parse_float(value, context)
                    else:
                        row[colname] = value
    for f in FLOAT_FIELDS:
        if f not in row:
            row[f] = None
    for f in INT_FIELDS:
        if f not in row:
            row[f] = None
    return row

def process_json_file(filepath):
    """Загружает JSON-файл и превращает его в список словарей."""
    filename = os.path.basename(filepath)
    rows = []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            js = json.load(f)
    except Exception as ex:
        logging.error(f"Ошибка загрузки файла {filepath}: {ex}")
        return []
    # Перебор турниров
    for tournament_key, records in js.items():
        # Универсальная обработка: dict, list, str, None
        entries = []
        if isinstance(records, list):
            entries = records
        elif isinstance(records, dict):
            entries = [records]
        else:
            logging.warning(f"[process_json_file] Некорректная запись в турнире {tournament_key}: {repr(records)[:100]}")
            continue
        for record in entries:
            try:
                if not isinstance(record, dict):
                    logging.warning(f"[process_json_file] Некорректная запись в турнире {tournament_key}: {repr(record)[:100]}")
                    continue
                tournament = record.get("body", {}).get("tournament", {})
                tournament_id = tournament.get("tournamentId", tournament_key)
                leaders = tournament.get("leaders", [])
                if isinstance(leaders, dict):
                    leaders = list(leaders.values())
                elif not isinstance(leaders, list):
                    leaders = []
                # === Обработка случая с пустым leaders ===
                if not leaders:
                    stub = {
                        'SourceFile': filename,
                        'tournamentId': tournament_id,
                        'employeeNumber': '00000000',
                        'lastName': 'None',
                        'firstName': 'None'
                    }
                    for field in FLOAT_FIELDS + INT_FIELDS:
                        stub[field] = None
                    rows.append(stub)
                    logging.info(f'Турнир {tournament_id} из файла {filename}: leaders пуст, добавлена заглушка')
                    continue

                # === Обработка нормальных лидеров ===
                for leader in leaders:
                    try:
                        row = flatten_leader(leader, tournament_id, filename)
                        rows.append(row)
                    except Exception as ex:
                        logging.error(
                            f"[flatten_leader] Ошибка обработки лидера в файле {filename} "
                            f"турнир {tournament_id} employee {leader.get('employeeNumber', 'N/A')}: {ex}"
                        )
            except Exception as ex:
                logging.error(
                    f"[process_json_file] Ошибка обработки записи в файле {filename}, турнир {tournament_key}: {ex}"
                )
    return rows

def load_json_folder(folder):
    """Читает все JSON-файлы из папки и объединяет их."""
    all_rows = []
    for fname in os.listdir(folder):
        if fname.lower().endswith('.json'):
            path = os.path.join(folder, fname)
            all_rows.extend(process_json_file(path))
    if not all_rows:
        logging.warning(f'Нет данных для экспорта из папки {folder}')
        return pd.DataFrame()
    df = pd.DataFrame(all_rows)
    return df

def make_compare_sheet(df_before, df_after, sheet_name):
    """Формирует таблицу сравнения показателей между выгрузками."""
    try:
        join_keys = COMPARE_KEYS
        before_uniq = df_before.drop_duplicates(subset=join_keys, keep='last')
        after_uniq  = df_after.drop_duplicates(subset=join_keys, keep='last')
        all_keys = pd.concat([before_uniq[join_keys], after_uniq[join_keys]]).drop_duplicates()
        before_uniq = before_uniq.set_index(join_keys)
        after_uniq  = after_uniq.set_index(join_keys)
        before_uniq = before_uniq[COMPARE_FIELDS] if len(before_uniq) else pd.DataFrame(columns=COMPARE_FIELDS)
        after_uniq  = after_uniq[COMPARE_FIELDS]  if len(after_uniq) else pd.DataFrame(columns=COMPARE_FIELDS)
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

        def rang_compare(row, before_col, after_col, status_dict):
            before = row.get(f'BEFORE_{before_col}', None)
            after  = row.get(f'AFTER_{after_col}', None)
            if pd.isnull(before) and not pd.isnull(after):
                return status_dict['val_add']
            if not pd.isnull(before) and pd.isnull(after):
                return status_dict['val_remove']
            if pd.isnull(before) and pd.isnull(after):
                return status_dict.get('val_norank', 'NO_RANK')
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

        # ratingCategoryName сравнение с логикой "меньше — лучше"
        def category_rank(cat):
            return CATEGORY_RANK_MAP.get(cat, 4)  # если неизвестно — вне призовых

        def category_compare(before_cat, after_cat):
            b_rank = category_rank(before_cat)
            a_rank = category_rank(after_cat)
            # 1) Не был в призовых или не было вообще, но попал в призовые - "ENTERED_PRIZE"
            if b_rank == 4 and a_rank < 4:
                return STATUS_RATING_CATEGORY["in2prize"]
            # 2) Не был в призовых - остался вне призовых - "STAYED_OUT"
            if b_rank == 4 and a_rank == 4:
                return STATUS_RATING_CATEGORY["stay_out"]
            # 3) Был в призовых - попал в "нужно поднажать" - "DROPPED_OUT_PRIZE"
            if b_rank < 4 and a_rank == 4:
                return STATUS_RATING_CATEGORY["from2out"]
            # 4) Был в призовых - нет в новом файле - "LOST_VIEW"
            if b_rank < 4 and (after_cat is None or after_cat == ""):
                return STATUS_RATING_CATEGORY["lost"]
            # 5) Призовое место не изменилось - "PRIZE_UNCHANGED"
            if b_rank == a_rank and b_rank < 4:
                return STATUS_RATING_CATEGORY["same"]
            # 6) Призовое место улучшилось (меньше — лучше) - "PRIZE_UP"
            if b_rank > a_rank and a_rank < 4:
                return STATUS_RATING_CATEGORY["up"]
            # 7) Призовое место ухудшилось (меньше — лучше) - "PRIZE_DOWN"
            if b_rank < a_rank and a_rank < 4:
                return STATUS_RATING_CATEGORY["down"]
            # Если ничего не подошло
            return ""

        # Для каждой категории делаем отдельное поле
        for group, col in [
            ('BANK',   'divisionRatings_BANK_ratingCategoryName'),
            ('TB',     'divisionRatings_TB_ratingCategoryName'),
            ('GOSB',   'divisionRatings_GOSB_ratingCategoryName'),
        ]:
            def cmp_func(row, col=col):
                before_cat = row.get(f'BEFORE_{col}', None)
                after_cat  = row.get(f'AFTER_{col}', None)
                return category_compare(before_cat, after_cat)
            compare_df[f"{col}_Compare"] = compare_df.apply(cmp_func, axis=1)

        final_cols = COMPARE_KEYS + [
            'New_Remove', 'indicatorValue_Compare',
            'divisionRatings_BANK_placeInRating_Compare',
            'divisionRatings_TB_placeInRating_Compare',
            'divisionRatings_GOSB_placeInRating_Compare',
            'divisionRatings_BANK_ratingCategoryName_Compare',
            'divisionRatings_TB_ratingCategoryName_Compare',
            'divisionRatings_GOSB_ratingCategoryName_Compare'
        ] + ['BEFORE_' + c for c in COMPARE_FIELDS] + ['AFTER_' + c for c in COMPARE_FIELDS]
        compare_df = compare_df.reindex(columns=final_cols)
        logging.info(f"[OK] Compare sheet готов: строк {len(compare_df)}, колонок {len(compare_df.columns)}")
        # === PATCH: ЛОГИРОВАНИЕ перед фильтрацией ===
        logging.info(f"[COMPARE] Строк до фильтрации по турнирам: {len(compare_df)}")
        logging.info(f"Уникальных турниров: {compare_df['tournamentId'].nunique()}")
        logging.info(f"Список турниров: {compare_df['tournamentId'].unique()[:10]}")
        # Если файл большой, логируем только часть уникальных id

        # === PATCH: Фильтрация по списку турниров ===
        if ALLOWED_TOURNAMENT_IDS:
            compare_df = compare_df[compare_df['tournamentId'].isin(ALLOWED_TOURNAMENT_IDS)]
            logging.info(f"[COMPARE] После фильтрации по ALLOWED_TOURNAMENT_IDS осталось строк: {len(compare_df)}")

        # === PATCH: определяем статусные колонки (ВАЖНО!) ===
        # Ниже — ваши реальные имена статусных колонок!
        status_cols = [
            'indicatorValue_Compare',
            'divisionRatings_BANK_placeInRating_Compare',
            'divisionRatings_TB_placeInRating_Compare',
            'divisionRatings_GOSB_placeInRating_Compare',
            'divisionRatings_BANK_ratingCategoryName_Compare',
            'divisionRatings_TB_ratingCategoryName_Compare',
            'divisionRatings_GOSB_ratingCategoryName_Compare'
        ]
        # Проверим, что эти колонки есть
        for col in status_cols:
            if col not in compare_df.columns:
                logging.warning(f"[COMPARE] Нет колонки {col} в compare_df!")

        # === PATCH: фильтрация строк без изменений ===
        def is_any_change(row):
            for col in status_cols:
                val = str(row.get(col, "")).strip()
                if val not in NOCHANGE_STATUSES:
                    return True  # есть отличие!
            return False  # только статусы из списка — строку УБИРАЕМ

        mask = compare_df.apply(is_any_change, axis=1)
        compare_df = compare_df[mask].reset_index(drop=True)

        logging.info(f"[COMPARE] После удаления строк без изменений осталось: {len(compare_df)}")
        # Для отладки — выводим статистику по статусам:
        try:
            stats = {}
            for col in status_cols:
                stats[col] = dict(compare_df[col].value_counts())
            logging.info(f"[COMPARE] Статистика по статусным колонкам: {stats}")
        except Exception as ex:
            logging.warning(f"[COMPARE] Не удалось вывести статистику по статусам: {ex}")

        return compare_df, sheet_name
    except Exception as ex:
        logging.error(f"Ошибка в make_compare_sheet: {ex}")
        return pd.DataFrame(), sheet_name

def add_smart_table(writer, df, sheet_name, table_name):
    """
    Экспортирует DataFrame на лист Excel с автоформатированием:
    - Жирные заголовки
    - Автоширина столбцов
    """
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    try:
        worksheet = writer.sheets[sheet_name]

        # Сделать заголовок жирным
        for cell in next(worksheet.iter_rows(min_row=1, max_row=1)):
            cell.font = Font(bold=True)

        # Автоширина столбцов
        for i, column in enumerate(df.columns, 1):
            max_length = max(
                df[column].astype(str).map(len).max(),
                len(str(column))
            )
            worksheet.column_dimensions[get_column_letter(i)].width = max_length + 2

    except Exception:
        pass  # Для других движков просто пропустить

def apply_status_colors(writer, df, sheet_name, status_color_map, status_columns):
    """
    Закрашивает ячейки сравнения в Excel по статусам (openpyxl).
    """
    worksheet = writer.sheets[sheet_name]
    dark_bg = APPLY_DARK_BG_COLORS
    # Cтатусы с белым шрифтом
    statuses_with_white_font = set()
    for status, color in status_color_map.items():
        color_clean = color.lstrip('#').lower()
        if color_clean in dark_bg:
            statuses_with_white_font.add(status)
    for col_name in status_columns:
        if col_name not in df.columns:
            continue
        col_idx = df.columns.get_loc(col_name) + 1  # 1-based
        for row_idx, value in enumerate(df[col_name], 2):  # 2 = первая строка данных
            status = str(value)
            color = status_color_map.get(status)
            if color:
                color_clean = color.lstrip('#')
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.fill = PatternFill(fill_type='solid', fgColor=color_clean)
                if status in statuses_with_white_font:
                    cell.font = Font(color="FFFFFF")

def add_status_legend(writer, status_colors, status_ru_dict, status_rating_category, sheet_name=STATUS_LEGEND_SHEET):
    """
    Добавляет лист Excel с легендой по статусам (openpyxl-версия).
    """
    rows = []
    for key, eng_status in status_rating_category.items():
        ru, comment = status_ru_dict.get(eng_status, ("", ""))
        color = status_colors.get(eng_status, DEFAULT_STATUS_COLOR)
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
    # Цветное оформление
    for row_idx, row in enumerate(legend_df.itertuples(index=False), start=2):
        color = str(row[2])  # "Excel fill color"
        color_clean = color.lstrip("#")
        cell = worksheet.cell(row=row_idx, column=3)  # 3-я колонка — Excel fill color
        cell.fill = PatternFill(fill_type="solid", fgColor=color_clean)
        if color_clean.lower() in LEGEND_DARK_BG_COLORS:
            cell.font = Font(color="FFFFFF")


def build_final_sheet_fast(compare_df, allowed_ids, out_prefix, category_rank_map, df_before, df_after, log):
    """Строит итоговый лист по всем турнирам и сотрудникам.
    Оптимизированная версия: lookup-структуры вместо фильтрации.
    """
    log.info(FINAL_START_MESSAGE)
    if allowed_ids:
        tournaments = list(allowed_ids)
    else:
        tournaments = sorted(compare_df['tournamentId'].dropna().unique())

    emp_cols = ['employeeNumber', 'lastName', 'firstName']
    employees = compare_df[emp_cols].drop_duplicates().sort_values(emp_cols)
    log.info(f"[FINAL] Уникальных сотрудников: {len(employees)}")
    total_loops = len(employees) * len(tournaments)
    log.info(f"[FINAL] Всего итераций обработки: {total_loops}")

    # --- OPTIMIZATION 1: Быстрый доступ к compare_df по сотруднику+турнир (tuple index)
    indexed = compare_df.set_index(['employeeNumber', 'lastName', 'firstName', 'tournamentId'])
    # --- OPTIMIZATION 2: Быстрый доступ к признаку "был ли в before/after"
    # (employeeNumber, tournamentId) -> True/False
    before_pairs = set(zip(df_before['employeeNumber'], df_before['tournamentId']))
    after_pairs = set(zip(df_after['employeeNumber'], df_after['tournamentId']))

    status_cols = [
        'divisionRatings_BANK_ratingCategoryName_Compare',
        'divisionRatings_TB_ratingCategoryName_Compare',
        'divisionRatings_GOSB_ratingCategoryName_Compare'
    ]

    result_rows = []

    for emp_idx, (_, emp) in enumerate(employees.iterrows(), 1):
        emp_key = (emp['employeeNumber'], emp['lastName'], emp['firstName'])
        row = {col: emp[col] for col in emp_cols}

        for t_id in tournaments:
            idx = emp_key + (t_id,)
            best_val = None
            best_rank = float('inf')

            # --- OPTIMIZATION 3: Без try, через .get() для индекса
            subset = indexed.loc[idx] if idx in indexed.index else None
            if subset is not None:
                candidates = []
                for col in status_cols:
                    val = subset[col]
                    if pd.notnull(val) and str(val).strip().upper() not in ['NONE', 'NULL', '']:
                        candidates.append(val)
                for v in candidates:
                    rank = category_rank_map.get(v, 99)
                    if rank < best_rank:
                        best_val = v
                        best_rank = rank

            # --- OPTIMIZATION 4: Мгновенная проверка наличия в before/after без .any()
            was_in_before = (emp['employeeNumber'], t_id) in before_pairs
            was_in_after = (emp['employeeNumber'], t_id) in after_pairs

            if best_val is not None:
                final_value = best_val
            elif was_in_before or was_in_after:
                final_value = "CONT"
            else:
                final_value = "Not_used"

            row[t_id] = final_value
        result_rows.append(row)

    final_df = pd.DataFrame(result_rows)
    log.info(f"[FINAL] Итоговая таблица построена: {final_df.shape[0]} x {final_df.shape[1]}")
    return final_df, tournaments


def main():
    """Основная точка входа в программу."""
    logger = setup_logger(LOG_DIR, LOG_BASENAME)

    t_start = datetime.now()
    before_path = os.path.join(SOURCE_DIR, BEFORE_FILENAME)
    after_path = os.path.join(SOURCE_DIR, AFTER_FILENAME)
    now = datetime.now()
    ts = now.strftime("%Y%m%d_%H%M%S")
    sheet_before = f"BEFORE_{ts}"
    sheet_after = f"AFTER_{ts}"
    sheet_compare = f"COMPARE_{ts}"

    logger.info(f"[MAIN] Читаем BEFORE: {before_path}")
    t_beg_before = datetime.now()
    rows_before = process_json_file(before_path)
    df_before = pd.DataFrame(rows_before)
    t_end_before = datetime.now()
    log_data_stats(df_before, "BEFORE")

    logger.info(f"[MAIN] Читаем AFTER: {after_path}")
    t_beg_after = datetime.now()
    rows_after = process_json_file(after_path)
    df_after = pd.DataFrame(rows_after)
    t_end_after = datetime.now()
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
    t_beg_compare = datetime.now()
    compare_df, sheet_compare = make_compare_sheet(
        df_before, df_after, sheet_compare
    )
    t_end_compare = datetime.now()

    # Построение финального листа (FINAL)
    t_beg_final = datetime.now()
    final_df, tournaments = build_final_sheet_fast(compare_df, ALLOWED_TOURNAMENT_IDS, "FINAL_", CATEGORY_RANK_MAP, df_before, df_after, logger)
    t_end_final = datetime.now()

    final_status_cols = tournaments
    log_compare_stats(compare_df)

    base, ext = os.path.splitext(RESULT_EXCEL)
    result_excel_ts = f"{base}_{ts}{ext}"
    out_excel = os.path.join(TARGET_DIR, result_excel_ts)

    t_beg_export = datetime.now()
    with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
        logger.info(f"[MAIN] Экспортируем BEFORE лист {sheet_before}")
    #    df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        add_smart_table(writer, df_before, sheet_before, "SMART_" + sheet_before)
        logger.info(f"[MAIN] Экспортируем AFTER лист {sheet_after}")
     #   df_after.to_excel(writer, index=False, sheet_name=sheet_after)
        add_smart_table(writer, df_after, sheet_after, "SMART_" + sheet_after)
        logger.info(f"[MAIN] Экспортируем COMPARE лист {sheet_compare}")
     #   compare_df.to_excel(writer, index=False, sheet_name=sheet_compare)
        add_smart_table(writer, compare_df, sheet_compare, "SMART_" + sheet_compare)
        logger.info(f"[MAIN] Экспортируем FINAL лист FINAL_{ts}")
      #  final_df.to_excel(writer, index=False, sheet_name="FINAL_" + ts)
        add_smart_table(writer, final_df, "FINAL_" + ts, "SMART_" + "FINAL_" + ts)

        apply_status_colors(
            writer,
            final_df,
            "FINAL_" + ts,
            STATUS_COLORS_DICT,
            final_status_cols
        )
        logger.info(f"[MAIN] Применена цветовая раскраска к FINAL_{ts}")

        apply_status_colors(
            writer,
            compare_df,
            sheet_compare,
            STATUS_COLORS_DICT,
            STATUS_COLOR_COLUMNS
        )
        logger.info(f"[MAIN] Применена цветовая раскраска к COMPARE_{ts}")
        add_status_legend(writer, STATUS_COLORS_DICT, STATUS_RU_DICT, STATUS_RATING_CATEGORY, sheet_name=STATUS_LEGEND_SHEET)
        logger.info(f"[MAIN] Все данные выгружены в файл: {out_excel}")
    t_end_export = datetime.now()

    summary = SUMMARY_TEMPLATE.format(
        tourn=len(tournaments),
        emps=len(final_df),
        changes=len(compare_df),
        t1=(t_end_before - t_beg_before).total_seconds(),
        t2=(t_end_after - t_beg_after).total_seconds(),
        t3=(t_end_compare - t_beg_compare).total_seconds(),
        t4=(t_end_final - t_beg_final).total_seconds(),
        t5=(t_end_export - t_beg_export).total_seconds(),
        tt=(t_end_export - t_start).total_seconds(),
    )
    logger.info(summary)


if __name__ == "__main__":
    main()
