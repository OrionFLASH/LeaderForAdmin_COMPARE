from datetime import datetime

# Пути
SOURCE_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON"
TARGET_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX"
LOG_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//LOGS"
LOG_BASENAME = "LOG"
BEFORE_FILENAME = "LFA_0.json"
AFTER_FILENAME = "LFA_4.json"
RESULT_EXCEL = "LFA_COMPARE.xlsx"

# Структура колонок
PRIORITY_COLS = [
    'SourceFile', 'tournamentId', 'employeeNumber', 'lastName', 'firstName',
    'terDivisionName', 'divisionRatings_BANK_groupId', 'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId', 'employeeStatus', 'businessBlock',
    'successValue', 'indicatorValue', 'divisionRatings_BANK_placeInRating',
    'divisionRatings_TB_placeInRating', 'divisionRatings_GOSB_placeInRating',
    'divisionRatings_BANK_ratingCategoryName', 'divisionRatings_TB_ratingCategoryName',
    'divisionRatings_GOSB_ratingCategoryName',
]
COMPARE_KEYS = [
    'tournamentId',
    'employeeNumber',
    'lastName',
    'firstName',
]
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
INT_FIELDS = [
    'divisionRatings_BANK_groupId',
    'divisionRatings_TB_groupId',
    'divisionRatings_GOSB_groupId',
    'divisionRatings_BANK_placeInRating',
    'divisionRatings_TB_placeInRating',
    'divisionRatings_GOSB_placeInRating',
]
FLOAT_FIELDS = [
    'indicatorValue',
    'successValue',
]

# Статусы для сравнения
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
    "val_down":     "Rang BANK DOWN"
}
STATUS_TB_PLACE = {
    "val_add":      "Rang TB NEW",
    "val_remove":   "Rang TB REMOVE",
    "val_nochange": "Rang TB NO CHANGE",
    "val_up":       "Rang TB UP",
    "val_down":     "Rang TB DOWN"
}
STATUS_GOSB_PLACE = {
    "val_add":      "Rang GOSB NEW",
    "val_remove":   "Rang GOSB REMOVE",
    "val_nochange": "Rang GOSB NO CHANGE",
    "val_up":       "Rang GOSB UP",
    "val_down":     "Rang GOSB DOWN"
}

# Цвета для Excel-раскраски статусов
COMPARE_STATUS_COLORS = {
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
    'Remove FROM':      '#A6A6A6',  # Темно-серый
    'Rang BANK REMOVE': '#A6A6A6',
    'Rang TB REMOVE':   '#A6A6A6',
    'Rang GOSB REMOVE': '#A6A6A6',
    'Remove':           '#A6A6A6',
    'New':              '#E2EFDA',
}
