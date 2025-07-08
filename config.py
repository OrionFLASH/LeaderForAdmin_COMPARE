from datetime import datetime

# Пути
SOURCE_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//JSON"
TARGET_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//XLSX"
LOG_DIR = "//Users//orionflash//Desktop//MyProject//LeaderForAdmin_skript//LOGS"
LOG_BASENAME = "LOG"
BEFORE_FILENAME = "leadersForAdmin_ALL_20250708-140508.json"
AFTER_FILENAME = "leadersForAdmin_ALL_20250708-140508.json"
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

    # можно добавить аналогично другие статусы, если требуется
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
