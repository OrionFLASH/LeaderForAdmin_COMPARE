import pandas as pd
import logging
from config import (
    COMPARE_KEYS, COMPARE_FIELDS,
    STATUS_NEW_REMOVE, STATUS_INDICATOR,
    STATUS_BANK_PLACE, STATUS_TB_PLACE, STATUS_GOSB_PLACE
)

def make_compare_sheet(df_before, df_after, sheet_name):
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
                return ""
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

        final_cols = COMPARE_KEYS + [
            'New_Remove', 'indicatorValue_Compare',
            'divisionRatings_BANK_placeInRating_Compare',
            'divisionRatings_TB_placeInRating_Compare',
            'divisionRatings_GOSB_placeInRating_Compare'
        ] + ['BEFORE_' + c for c in COMPARE_FIELDS] + ['AFTER_' + c for c in COMPARE_FIELDS]
        compare_df = compare_df.reindex(columns=final_cols)
        logging.info(f"[OK] Compare sheet готов: строк {len(compare_df)}, колонок {len(compare_df.columns)}")
        return compare_df, sheet_name
    except Exception as ex:
        logging.error(f"Ошибка в make_compare_sheet: {ex}")
        return pd.DataFrame(), sheet_name
