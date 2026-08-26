# -*- coding: utf-8 -*-
"""
Menu configuration for the coffee shop POS app.
Each item: key (used internally / in Excel), display name (Persian),
icon file name (inside assets/), row and column position in the GUI grid,
and whether it belongs to the highlighted "Macchiato" group.
"""

MENU_ITEMS = [
    # --- Row 1 : Hot classics ---
    {"key": "aslashi", "name": "اسلاشی (یخ در بهشت)", "icon": "aslashi.png", "row": 0, "col": 0},
    {"key": "espresso", "name": "اسپرسو", "icon": "espresso.png", "row": 0, "col": 1},
    {"key": "espresso_single", "name": "اسپرسو سینگل", "icon": "espresso.png", "row": 0, "col": 2},
    {"key": "americano", "name": "امریکانو", "icon": "americano.png", "row": 0, "col": 3},
    {"key": "latte", "name": "لاته", "icon": "latte.png", "row": 0, "col": 4},
    {"key": "cappuccino", "name": "کاپوچینو", "icon": "cappuccino.png", "row": 0, "col": 5},
    {"key": "hot_chocolate", "name": "هات چاکلت", "icon": "hot_chocolate.png", "row": 0, "col": 6},

    # --- Row 2 : Chocolates / Macchiato group / Mocha / Cortado ---
    {"key": "pink_chocolate", "name": "پینک چاکلت", "icon": "pink_chocolate.png", "row": 1, "col": 0},
    {"key": "blue_chocolate", "name": "بلو چاکلت", "icon": "blue_chocolate.png", "row": 1, "col": 1},
    {"key": "macchiato_caramel", "name": "ماکیاتو کارامل", "icon": "macchiato_caramel.png", "row": 1, "col": 2, "group": "macchiato"},
    {"key": "macchiato_hazelnut", "name": "ماکیاتو فندق", "icon": "macchiato_hazelnut.png", "row": 1, "col": 3, "group": "macchiato"},
    {"key": "macchiato_irish", "name": "ماکیاتو ایریش", "icon": "macchiato_irish.png", "row": 1, "col": 4, "group": "macchiato"},
    {"key": "mocha", "name": "موکا", "icon": "mocha.png", "row": 1, "col": 5, "group": "macchiato"},
    {"key": "cortado", "name": "کورتادو", "icon": "cappuccino.png", "row": 1, "col": 6},

    # --- Row 3 : Iced & other drinks ---
    {"key": "ice_americano", "name": "ایس امریکانو", "icon": "ice_americano.png", "row": 2, "col": 0},
    {"key": "ice_latte", "name": "ایس لاته", "icon": "ice_latte.png", "row": 2, "col": 1},
    {"key": "ice_cappuccino", "name": "ایس کاپوچینو", "icon": "ice_cappuccino.png", "row": 2, "col": 2},
    {"key": "ice_macchiato", "name": "ایس ماکیاتو", "icon": "ice_macchiato.png", "row": 2, "col": 3},
    {"key": "ice_chocolate", "name": "ایس چاکلت", "icon": "ice_chocolate.png", "row": 2, "col": 4},
    {"key": "black_tea", "name": "چای سیاه", "icon": "black_tea.png", "row": 2, "col": 5},
    {"key": "turkish_coffee", "name": "قهوه ترک", "icon": "turkish_coffee.png", "row": 2, "col": 6},
]

MENU_BY_KEY = {item["key"]: item for item in MENU_ITEMS}
