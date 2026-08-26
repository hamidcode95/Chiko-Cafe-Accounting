# -*- coding: utf-8 -*-
"""
Loads the bundled Vazirmatn (Persian/UI) font files at startup and exposes
the resulting family name so every window in the app uses the same font,
regardless of what's installed on the end user's Windows PC.

To change the app's font later: drop different .ttf files into fonts/,
update TTF_FILES below, and everything that imports APP_FONT_FAMILY from
this module will pick up the new font automatically.
"""

import os

from PyQt5.QtGui import QFontDatabase

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(BASE_DIR, "fonts")

TTF_FILES = [
    "Vazirmatn-UI-Regular.ttf",
    "Vazirmatn-UI-Medium.ttf",
    "Vazirmatn-UI-Bold.ttf",
]

# Fallback if the bundled files are ever missing - Tahoma ships with every
# Windows install and supports Persian reasonably well.
APP_FONT_FAMILY = "Tahoma"


def load_app_fonts():
    """Call this once, right after creating the QApplication, before
    building any windows."""
    global APP_FONT_FAMILY
    loaded_family = None
    for filename in TTF_FILES:
        path = os.path.join(FONTS_DIR, filename)
        if not os.path.exists(path):
            continue
        font_id = QFontDatabase.addApplicationFont(path)
        if font_id != -1:
            families = QFontDatabase.applicationFontFamilies(font_id)
            if families:
                loaded_family = families[0]
    if loaded_family:
        APP_FONT_FAMILY = loaded_family
    return APP_FONT_FAMILY
