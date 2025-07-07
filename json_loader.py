import os
import json
import pandas as pd
import re
import logging
from config import FLOAT_FIELDS, INT_FIELDS

def parse_float(val, context=None):
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
        for record in records:
            try:
                tournament = record.get("body", {}).get("tournament", {})
                tournament_id = tournament.get("tournamentId", tournament_key)
                leaders = tournament.get("leaders", [])
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
