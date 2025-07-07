import os
import json
import pandas as pd
from datetime import datetime

def flatten_leader(leader, tournament_id):
    row = {'tournamentId': tournament_id}
    for k, v in leader.items():
        if k != "divisionRatings":
            row[k] = v
    if "divisionRatings" in leader and leader["divisionRatings"]:
        for div in leader["divisionRatings"]:
            group = div.get("groupCode", "DIV")
            for field, value in div.items():
                if field == "groupCode":
                    continue
                colname = f"divisionRatings_{group}_{field}"
                row[colname] = value
    return row

def process_json_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        js = json.load(f)
    all_rows = []
    for block in js.values():
        # Универсально: если block — список, обработаем каждый элемент, иначе один объект
        blocks = block if isinstance(block, list) else [block]
        for subblock in blocks:
            tournament = subblock.get("body", {}).get("tournament", {})
            tournament_id = tournament.get("tournamentId")
            for leader in tournament.get("leaders", []):
                all_rows.append(flatten_leader(leader, tournament_id))
    return pd.DataFrame(all_rows)

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

    base, ext = os.path.splitext(result_excel)
    result_excel_ts = f"{base}_{ts}{ext}"
    out_excel = os.path.join(target_dir, result_excel_ts)

    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
        df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        df_after.to_excel(writer, index=False, sheet_name=sheet_after)
    print(f"[SUCCESS] Файл {out_excel} создан.")

if __name__ == "__main__":
    main(
        source_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON",
        target_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX",
        before_filename="LFA_0.json",
        after_filename="LFA_1.json",
        result_excel="LFA_COMPARE.xlsx"
    )
