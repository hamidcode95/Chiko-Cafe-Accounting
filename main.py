# -*- coding: utf-8 -*-
"""
Coffee Shop POS - a simple Windows desktop app (PyQt5) that mirrors the
shop's drink menu as a grid of buttons. Clicking a drink records one sale
for "today" (Jalali/Shamsi calendar) and writes it into an Excel workbook
(sales_data/sales.xlsx), one sheet per day. Includes a sales report screen
(single day or date range) and support for one-off discounted prices.

Run:
    pip install -r requirements.txt
    python main.py
"""

import os
import sys

from PyQt5.QtCore import Qt, QLocale
from PyQt5.QtGui import QFont, QIcon, QPixmap, QDoubleValidator
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QGridLayout, QVBoxLayout,
    QHBoxLayout, QLabel, QPushButton, QLineEdit, QFrame, QScrollArea,
    QMessageBox, QSizePolicy, QAction, QFileDialog, QInputDialog
)

from menu_data import MENU_ITEMS
from price_store import load_prices, save_prices
import excel_manager
import date_utils
from report_dialog import ReportDialog
from backdate_dialog import BackdateSaleDialog
import app_font
from app_font import load_app_fonts

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

APP_BG = "#FFF3E0"
CARD_BG = "#FFFFFF"
ACCENT = "#4E342E"
ACCENT_LIGHT = "#8D6E63"
MACCHIATO_BORDER = "#EC407A"


class ItemCard(QFrame):
    """One drink: icon button + name + editable price field + today's count."""

    def __init__(self, item, price, on_sale, on_price_change, on_undo,
                 on_custom_sale, parent=None):
        super().__init__(parent)
        self.item = item
        self.on_sale = on_sale
        self.on_price_change = on_price_change
        self.on_undo = on_undo
        self.on_custom_sale = on_custom_sale

        self.setObjectName("card")
        self.setFrameShape(QFrame.NoFrame)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(4)

        # --- Icon button (click => +1 sale) ---
        self.icon_button = QPushButton()
        self.icon_button.setCursor(Qt.PointingHandCursor)
        icon_path = os.path.join(ASSETS_DIR, item["icon"])
        if os.path.exists(icon_path):
            self.icon_button.setIcon(QIcon(QPixmap(icon_path)))
            self.icon_button.setIconSize(QPixmap(icon_path).size().boundedTo(
                QPixmap(140, 130).size()))
        self.icon_button.setFixedSize(150, 130)
        self.icon_button.setToolTip("برای ثبت فروش کلیک کنید")
        self.icon_button.clicked.connect(self._handle_click)
        layout.addWidget(self.icon_button, alignment=Qt.AlignHCenter)

        # --- Name label ---
        name_label = QLabel(item["name"])
        name_label.setAlignment(Qt.AlignCenter)
        name_label.setWordWrap(True)
        name_label.setFont(QFont(app_font.APP_FONT_FAMILY, 10, QFont.Bold))
        layout.addWidget(name_label)

        # --- Price field (editable "label box" under each drink) ---
        self.price_edit = QLineEdit()
        self.price_edit.setAlignment(Qt.AlignCenter)
        self.price_edit.setPlaceholderText("قیمت (تومان)")
        validator = QDoubleValidator(0, 100000000, 0)
        validator.setLocale(QLocale(QLocale.English))
        self.price_edit.setValidator(validator)
        self.price_edit.setText(self._fmt_price(price))
        self.price_edit.editingFinished.connect(self._handle_price_change)
        layout.addWidget(self.price_edit)

        # --- Today count + undo ---
        bottom_row = QHBoxLayout()
        self.count_label = QLabel("امروز: 0")
        self.count_label.setAlignment(Qt.AlignCenter)
        self.count_label.setFont(QFont(app_font.APP_FONT_FAMILY, 9))
        bottom_row.addWidget(self.count_label)

        self.undo_button = QPushButton("↺")
        self.undo_button.setFixedWidth(28)
        self.undo_button.setToolTip("لغو آخرین فروش این مورد")
        self.undo_button.setCursor(Qt.PointingHandCursor)
        self.undo_button.clicked.connect(self._handle_undo)
        bottom_row.addWidget(self.undo_button)
        layout.addLayout(bottom_row)

        # --- Discount / custom price sale button ---
        self.discount_button = QPushButton("٪ فروش با قیمت دیگر")
        self.discount_button.setCursor(Qt.PointingHandCursor)
        self.discount_button.setToolTip(
            "برای ثبت فروش این نوشیدنی با قیمت متفاوت (تخفیف/مهمان ویژه) از "
            "قیمت پیش‌فرض بالا"
        )
        self.discount_button.clicked.connect(self._handle_custom_sale)
        layout.addWidget(self.discount_button)

        self.setStyleSheet(f"""
            QFrame#card {{
                background-color: {CARD_BG};
                border-radius: 12px;
                border: 1px solid #E0D6CC;
            }}
            QPushButton {{
                background-color: #FAF3EC;
                border: 1px solid #D7CCC8;
                border-radius: 10px;
            }}
            QPushButton:hover {{
                background-color: #F0E4D8;
                border: 1px solid {ACCENT_LIGHT};
            }}
            QPushButton:pressed {{
                background-color: #E4D2C3;
            }}
            QLineEdit {{
                border: 1px solid #C9B8A8;
                border-radius: 8px;
                padding: 3px;
            }}
            QPushButton#discount {{
                font-size: 10px;
                padding: 3px;
            }}
        """)
        self.discount_button.setObjectName("discount")

        if item.get("group") == "macchiato":
            self.setStyleSheet(self.styleSheet() + f"""
                QFrame#card {{
                    border: 2px solid {MACCHIATO_BORDER};
                }}
            """)

    @staticmethod
    def _fmt_price(value):
        try:
            return f"{int(float(value)):,}"
        except (ValueError, TypeError):
            return "0"

    def _current_price(self):
        text = self.price_edit.text().replace(",", "").strip()
        try:
            return float(text) if text else 0
        except ValueError:
            return 0

    def _handle_click(self):
        price = self._current_price()
        new_count = self.on_sale(self.item["key"], price)
        self.set_count(new_count)

    def _handle_undo(self):
        new_count = self.on_undo(self.item["key"])
        self.set_count(new_count)

    def _handle_custom_sale(self):
        default_price = self._current_price()
        price, ok = QInputDialog.getDouble(
            self,
            f"فروش {self.item['name']} با قیمت دیگر",
            "قیمت این فروش (تومان):",
            value=default_price,
            min=0,
            max=100000000,
            decimals=0,
        )
        if not ok:
            return
        new_count = self.on_custom_sale(self.item["key"], price)
        self.set_count(new_count)

    def _handle_price_change(self):
        price = self._current_price()
        self.price_edit.setText(self._fmt_price(price))
        self.on_price_change(self.item["key"], price)

    def set_count(self, count):
        self.count_label.setText(f"امروز: {count}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("صندوق فروش کافه - مدیریت فروش روزانه")
        self.resize(1300, 780)
        self.setLayoutDirection(Qt.RightToLeft)

        self.item_keys = [item["key"] for item in MENU_ITEMS]
        self.prices = load_prices(self.item_keys)

        self._build_ui()
        self._restore_today_counts()

    # ---------------------------------------------------------- UI build
    def _build_ui(self):
        central = QWidget()
        central.setStyleSheet(f"background-color: {APP_BG};")
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(16, 16, 16, 16)

        # --- Header ---
        header = QHBoxLayout()
        title = QLabel("☕ صندوق فروش کافه")
        title.setFont(QFont(app_font.APP_FONT_FAMILY, 18, QFont.Bold))
        title.setStyleSheet(f"color: {ACCENT};")
        header.addWidget(title)

        header.addStretch()

        today_label = QLabel(f"تاریخ امروز: {date_utils.today_display()}")
        today_label.setFont(QFont(app_font.APP_FONT_FAMILY, 11))
        header.addWidget(today_label)
        outer.addLayout(header)

        subtitle = QLabel(
            "برای ثبت فروش، روی تصویر هر نوشیدنی کلیک کنید. قیمت هر نوشیدنی را "
            "در کادر زیر آن وارد/ویرایش کنید."
        )
        subtitle.setFont(QFont(app_font.APP_FONT_FAMILY, 10))
        subtitle.setStyleSheet("color: #6D4C41;")
        outer.addWidget(subtitle)

        # --- Scrollable menu grid ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        grid_holder = QWidget()
        self.grid = QGridLayout(grid_holder)
        self.grid.setSpacing(14)
        scroll.setWidget(grid_holder)
        outer.addWidget(scroll)

        self.cards = {}
        for item in MENU_ITEMS:
            card = ItemCard(
                item,
                self.prices.get(item["key"], 0),
                on_sale=self._record_sale,
                on_price_change=self._record_price_change,
                on_undo=self._undo_sale,
                on_custom_sale=self._record_custom_sale,
            )
            self.grid.addWidget(card, item["row"], item["col"])
            self.cards[item["key"]] = card

        # --- Footer / status bar ---
        footer = QHBoxLayout()
        open_folder_btn = QPushButton("باز کردن پوشه فایل اکسل")
        open_folder_btn.setCursor(Qt.PointingHandCursor)
        open_folder_btn.clicked.connect(self._open_excel_folder)
        footer.addWidget(open_folder_btn)

        report_btn = QPushButton("📊 گزارش فروش")
        report_btn.setCursor(Qt.PointingHandCursor)
        report_btn.setStyleSheet(
            f"QPushButton {{ background-color: {ACCENT}; color: white; "
            f"padding: 6px 14px; border-radius: 8px; font-weight: bold; }}"
            f"QPushButton:hover {{ background-color: {ACCENT_LIGHT}; }}"
        )
        report_btn.clicked.connect(self._open_report)
        footer.addWidget(report_btn)

        backdate_btn = QPushButton("➕ ثبت فروش فراموش‌شده")
        backdate_btn.setCursor(Qt.PointingHandCursor)
        backdate_btn.setStyleSheet(
            f"QPushButton {{ background-color: {ACCENT_LIGHT}; color: white; "
            f"padding: 6px 14px; border-radius: 8px; font-weight: bold; }}"
        )
        backdate_btn.clicked.connect(self._open_backdate_dialog)
        footer.addWidget(backdate_btn)

        footer.addStretch()
        outer.addLayout(footer)

        self.statusBar().showMessage("آماده")

    # ------------------------------------------------------------ logic
    def _restore_today_counts(self):
        """If the app was closed and reopened the same day, reload counts
        already recorded in the Excel file so the on-screen numbers match."""
        try:
            counts = excel_manager.get_today_counts()
        except Exception:
            counts = {}
        for key, count in counts.items():
            if key in self.cards:
                self.cards[key].set_count(count)

    def _record_sale(self, item_key, price):
        """Normal click on the drink icon: sell at the price shown in its box."""
        try:
            new_count = excel_manager.record_sale(item_key, price)
            self.statusBar().showMessage(
                f"فروش ثبت شد: {MENU_ITEMS_BY_KEY[item_key]['name']} "
                f"به قیمت {int(price):,} تومان (تعداد امروز: {new_count})", 4000
            )
            return new_count
        except Exception as exc:
            QMessageBox.critical(self, "خطا در ثبت فروش", str(exc))
            return 0

    def _record_custom_sale(self, item_key, price):
        """'٪ فروش با قیمت دیگر' button: sell at a one-off price (discount /
        VIP price) without touching the item's saved default price."""
        try:
            new_count = excel_manager.record_sale(item_key, price)
            self.statusBar().showMessage(
                f"فروش با قیمت ویژه ثبت شد: {MENU_ITEMS_BY_KEY[item_key]['name']} "
                f"به قیمت {int(price):,} تومان (تعداد امروز: {new_count})", 5000
            )
            return new_count
        except Exception as exc:
            QMessageBox.critical(self, "خطا در ثبت فروش", str(exc))
            return 0

    def _undo_sale(self, item_key):
        try:
            new_count = excel_manager.decrement_sale(item_key)
            self.statusBar().showMessage(
                f"آخرین فروش لغو شد: {MENU_ITEMS_BY_KEY[item_key]['name']}", 4000
            )
            return new_count
        except Exception as exc:
            QMessageBox.critical(self, "خطا", str(exc))
            return 0

    def _record_price_change(self, item_key, price):
        """Just updates the saved 'default/suggested' price (prices.json).
        This is the price pre-filled for quick sales and for the discount
        dialog; it is NOT written into the Excel log by itself."""
        self.prices[item_key] = price
        save_prices(self.prices)

    def _open_excel_folder(self):
        folder = excel_manager.DATA_DIR
        os.makedirs(folder, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(folder)  # noqa: S606 - Windows only helper
        else:
            QFileDialog.getOpenFileName(self, "مسیر فایل اکسل", folder)

    def _open_report(self):
        dialog = ReportDialog(self)
        dialog.exec_()

    def _open_backdate_dialog(self):
        dialog = BackdateSaleDialog(self.prices, self._on_backdated_sale, self)
        dialog.exec_()

    def _on_backdated_sale(self, item_key, price, date_str):
        """If the backdated sale happens to be for *today*, refresh that
        card's on-screen counter too."""
        if date_str == date_utils.today_str() and item_key in self.cards:
            counts = excel_manager.get_today_counts()
            self.cards[item_key].set_count(counts.get(item_key, 0))
        self.statusBar().showMessage(
            f"فروش گذشته ثبت شد: {MENU_ITEMS_BY_KEY[item_key]['name']} "
            f"برای {date_utils.to_display(date_str)}", 5000
        )


MENU_ITEMS_BY_KEY = {item["key"]: item for item in MENU_ITEMS}


def main():
    app = QApplication(sys.argv)
    load_app_fonts()
    app.setLayoutDirection(Qt.RightToLeft)
    app.setFont(QFont(app_font.APP_FONT_FAMILY, 10))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
