# -*- coding: utf-8 -*-
"""
Dialog for recording a sale that happened on a *different* day than today
- e.g. "I sold an espresso on the 2nd of Shahrivar and forgot to enter it,
now it's the 4th." Lets the user pick any day (recent quick list, any
existing day with data, or type a date manually), pick the drink, set the
price, and add it straight into that day's sheet.
"""

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QDoubleValidator
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton,
    QLineEdit, QMessageBox,
)

from menu_data import MENU_ITEMS
import date_utils
import excel_manager
import app_font

ACCENT = "#4E342E"


class BackdateSaleDialog(QDialog):
    def __init__(self, prices, on_submit, parent=None):
        """
        prices: {item_key: price} - used to prefill the price field.
        on_submit(item_key, price, date_str): called after a successful save.
        """
        super().__init__(parent)
        self.prices = prices
        self.on_submit = on_submit

        self.setWindowTitle("ثبت فروش فراموش‌شده")
        self.setLayoutDirection(Qt.RightToLeft)
        self.resize(420, 260)

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        title = QLabel("➕ ثبت فروش برای یک روز دیگر")
        title.setFont(QFont(app_font.APP_FONT_FAMILY, 13, QFont.Bold))
        title.setStyleSheet(f"color: {ACCENT};")
        layout.addWidget(title)

        info = QLabel(
            "اگر یادتان رفته فروشی را همان روز ثبت کنید، از اینجا می‌توانید "
            "آن را برای روز درست ثبت کنید."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        # --- Date picker (editable combo: quick recent days + typing) ---
        date_row = QHBoxLayout()
        date_row.addWidget(QLabel("تاریخ فروش:"))
        self.date_combo = QComboBox()
        self.date_combo.setEditable(True)
        self._populate_dates()
        date_row.addWidget(self.date_combo)
        layout.addLayout(date_row)

        hint = QLabel("می‌توانید از لیست انتخاب کنید یا تاریخ را به فرمت ۱۴۰۵-۰۶-۰۲ تایپ کنید.")
        hint.setStyleSheet("color: #8D6E63; font-size: 10px;")
        layout.addWidget(hint)

        # --- Item picker ---
        item_row = QHBoxLayout()
        item_row.addWidget(QLabel("نوشیدنی:"))
        self.item_combo = QComboBox()
        for item in MENU_ITEMS:
            self.item_combo.addItem(item["name"], item["key"])
        self.item_combo.currentIndexChanged.connect(self._prefill_price)
        item_row.addWidget(self.item_combo)
        layout.addLayout(item_row)

        # --- Price ---
        price_row = QHBoxLayout()
        price_row.addWidget(QLabel("قیمت فروش (تومان):"))
        self.price_edit = QLineEdit()
        validator = QDoubleValidator(0, 100000000, 0)
        self.price_edit.setValidator(validator)
        price_row.addWidget(self.price_edit)
        layout.addLayout(price_row)
        self._prefill_price()

        layout.addStretch()

        # --- Buttons ---
        btn_row = QHBoxLayout()
        submit_btn = QPushButton("ثبت فروش")
        submit_btn.setCursor(Qt.PointingHandCursor)
        submit_btn.setStyleSheet(
            f"QPushButton {{ background-color: {ACCENT}; color: white; "
            f"padding: 6px 14px; border-radius: 8px; font-weight: bold; }}"
        )
        submit_btn.clicked.connect(self._submit)
        btn_row.addWidget(submit_btn)

        cancel_btn = QPushButton("انصراف")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    def _populate_dates(self):
        # merge recently-possible days with any day that already has data,
        # newest first, no duplicates
        candidates = list(dict.fromkeys(
            date_utils.recent_dates(60) + excel_manager.list_available_dates()[::-1]
        ))
        for raw in candidates:
            self.date_combo.addItem(date_utils.to_display(raw), raw)
        self.date_combo.setCurrentIndex(0)

    def _prefill_price(self):
        key = self.item_combo.currentData()
        price = self.prices.get(key, 0)
        self.price_edit.setText(str(int(price)) if price else "")

    def _resolve_date_str(self):
        """The combo is editable: its data-role only applies if the user
        picked an existing entry without editing the text. If they typed a
        custom value, fall back to parsing the visible text directly."""
        raw_data = self.date_combo.currentData()
        typed_text = self.date_combo.currentText().strip()

        if raw_data and date_utils.to_display(raw_data) == typed_text:
            return raw_data
        if date_utils.is_valid(typed_text):
            return typed_text
        return None

    def _submit(self):
        date_str = self._resolve_date_str()
        if not date_str:
            QMessageBox.warning(
                self, "تاریخ نامعتبر",
                "تاریخ را به فرمت صحیح وارد کنید، مثال: 1405-06-02"
            )
            return

        item_key = self.item_combo.currentData()
        price_text = self.price_edit.text().strip()
        try:
            price = float(price_text) if price_text else 0
        except ValueError:
            price = 0

        try:
            new_count = excel_manager.record_sale(item_key, price, date_str)
        except Exception as exc:
            QMessageBox.critical(self, "خطا در ثبت فروش", str(exc))
            return

        self.on_submit(item_key, price, date_str)
        item_name = self.item_combo.currentText()
        QMessageBox.information(
            self, "ثبت شد",
            f"فروش «{item_name}» برای تاریخ {date_utils.to_display(date_str)} "
            f"ثبت شد.\n(تعداد فروش این مورد در آن روز اکنون: {new_count})"
        )
        self.accept()
