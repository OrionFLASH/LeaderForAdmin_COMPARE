import os
import json
import pandas as pd
from datetime import datetime
from typing import Any, Dict, List, Union
from collections import defaultdict


def flatten_json_v2(
        y: Union[Dict, List],
        parent_key: str = '',
        sep: str = '_',
        key_count: dict = None
) -> List[Dict]:
    """
    Разворачивает вложенный JSON в плоскую структуру.
    Добавляет суффиксы к повторяющимся ключам (_01, _02 и т.д.).
    """
    items = []
    if key_count is None:
        key_count = defaultdict(int)

    def _flatten(obj, key_prefix, acc, local_key_count):
        if isinstance(obj, dict):
            for k, v in obj.items():
                # Учитываем повторения ключей на одном уровне
                full_key = f"{key_prefix}{sep}{k}" if key_prefix else k
                local_key_count[full_key] += 1
                count = local_key_count[full_key]
                key_with_idx = f"{full_key}_{count:02d}" if count > 1 else full_key
                _flatten(v, key_with_idx, acc, defaultdict(int))
        elif isinstance(obj, list):
            for idx, v in enumerate(obj):
                _flatten(v, f"{key_prefix}{sep}{idx}", acc, defaultdict(int))
        else:
            acc[key_prefix] = obj

    # Верхний уровень — dict или list
    if isinstance(y, dict):
        for k, v in y.items():
            acc = {}
            _flatten(v, k, acc, defaultdict(int))
            if acc:
                items.append(acc)
        if not items:
            items.append({})
    elif isinstance(y, list):
        for elem in y:
            acc = {}
            _flatten(elem, '', acc, defaultdict(int))
            if acc:
                items.append(acc)
    return items


def process_json_file(filepath: str) -> pd.DataFrame:
    with open(filepath, 'r', encoding='utf-8') as f:
        js = json.load(f)
    rows = flatten_json_v2(js)
    return pd.DataFrame(rows)


def main(source_dir: str, target_dir: str, before_filename: str, after_filename: str, result_excel: str):
    before_path = os.path.join(source_dir, before_filename)
    after_path = os.path.join(source_dir, after_filename)
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    now = datetime.now()
    ts = now.strftime("%y%m%d_%H%M")
    sheet_before = f"BEFORE_{ts}"
    sheet_after = f"AFTER_{ts}"

    df_before = process_json_file(before_path)
    df_after = process_json_file(after_path)

    out_excel = os.path.join(target_dir, result_excel)
    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
        df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        df_after.to_excel(writer, index=False, sheet_name=sheet_after)

    print(f"Done. Excel saved to: {out_excel}")


if __name__ == "__main__":
    # Замените пути и имена файлов на ваши
    main(
        source_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON",
        target_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX",
        before_filename="LFA_0.json",
        after_filename="LFA_2.json",
        result_excel="LFA_COMPARE.xlsx"
    )
