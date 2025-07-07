import os
import json
import pandas as pd
import re
import logging
from config import FLOAT_FIELDS, INT_FIELDS

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
