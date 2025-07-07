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
    Разворачивает вложенный JSON в плоскую структуру с уникальными именами столбцов.
    При совпадении имен на одном уровне добавляет суффиксы (_01, _02, ...).
    """
    items = []
    if key_count is None:
        key_count = defaultdict(int)

    def _flatten(obj, key_prefix, acc, local_key_count):
        if isinstance(obj, dict):
            for k, v in obj.items():
                # Учет повторяющихся ключей на текущем уровне
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

    # Начальная точка обработки — dict или list
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


def process_json_file(filepath: str, label: str) -> pd.DataFrame:
    print(f"[INFO] Чтение файла: {filepath}")
    # Проверяем, существует ли файл
    if not os.path.exists(filepath):
        print(f"[ERROR] Файл не найден: {filepath}")
        return pd.DataFrame()
    # Загружаем JSON
    with open(filepath, 'r', encoding='utf-8') as f:
        js = json.load(f)
    print(f"[INFO] Загружено содержимое файла. Начинаем разворачивание вложенной структуры...")
    # Разворачиваем JSON
    rows = flatten_json_v2(js)
    print(f"[INFO] Файл '{label}': всего записей (строк для экспорта): {len(rows)}")
    return pd.DataFrame(rows)


def main(
        source_dir: str,
        target_dir: str,
        before_filename: str,
        after_filename: str,
        result_excel: str
):
    print(f"[START] Старт программы сравнения JSON -> Excel")
    print(f"[PATHS] Исходная папка: {source_dir}")
    print(f"[PATHS] Папка для Excel: {target_dir}")
    print(f"[FILES] BEFORE: {before_filename} | AFTER: {after_filename} | Excel: {result_excel}")

    before_path = os.path.join(source_dir, before_filename)
    after_path = os.path.join(source_dir, after_filename)
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
        print(f"[INFO] Папка {target_dir} создана.")

    # Формируем имена листов по дате/времени
    now = datetime.now()
    ts = now.strftime("%y%m%d_%H%M")
    sheet_before = f"BEFORE_{ts}"
    sheet_after = f"AFTER_{ts}"

    # Обработка "до"
    df_before = process_json_file(before_path, label="BEFORE")
    print(f"[INFO] Столбцов (полей) в BEFORE: {len(df_before.columns)}")
    # Обработка "после"
    df_after = process_json_file(after_path, label="AFTER")
    print(f"[INFO] Столбцов (полей) в AFTER: {len(df_after.columns)}")

    # Путь к Excel
    out_excel = os.path.join(target_dir, result_excel)
    print(f"[INFO] Экспорт в файл: {out_excel}")

    # Сохраняем оба датафрейма на разные листы Excel
    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
        df_before.to_excel(writer, index=False, sheet_name=sheet_before)
        print(f"[OK] Записано в лист '{sheet_before}': {len(df_before)} строк.")
        df_after.to_excel(writer, index=False, sheet_name=sheet_after)
        print(f"[OK] Записано в лист '{sheet_after}': {len(df_after)} строк.")

    print(f"[SUCCESS] Готово. Excel файл сохранен по адресу: {out_excel}")


if __name__ == "__main__":
    main(
        source_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON",
        target_dir="//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX",
        before_filename="LFA_0.json",
        after_filename="LFA_2.json",
        result_excel="LFA_COMPARE.xlsx"
    )
