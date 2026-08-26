# -*- coding: utf-8 -*-
"""
Sales report dialog: pick a single day or a date range (Jalali calendar)
from the days that actually have data, and view/export the aggregated
sales report (count + amount per drink, across that range).
"""

import os

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QFileDialog,
)

import date_utils
import excel_manager
import app_font

ACCENT = "#4E342E"


class ReportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("گزارش فروش")
        self.setLayoutDirection(Qt.RightToLeft)
        self.resize(700, 560)

        self.available_dates = excel_manager.list_available_dates()
        self._last_report = None
        self._last_range = None

        self._build_ui()
        self._populate_date_combos()
        if self.available_dates:
            self._show_report()

    # ------------------------------------------------------------ UI
    def _build_ui(self):
        layout = QVBoxLayout(self)

        title = QLabel("📊 گزارش فروش")
        title.setFont(QFont(app_font.APP_FONT_FAMILY, 14, QFont.Bold))
        title.setStyleSheet(f"color: {ACCENT};")
        layout.addWidget(title)

        if not self.available_dates:
            layout.addWidget(QLabel("هنوز هیچ فروشی ثبت نشده تا گزارشی نمایش داده شود."))
            close_btn = QPushButton("بستن")
            close_btn.clicked.connect(self.reject)
            layout.addWidget(close_btn)
            return

        # --- Range pickers ---
        picker_row = QHBoxLayout()
        picker_row.addWidget(QLabel("از تاریخ:"))
        self.from_combo = QComboBox()
        picker_row.addWidget(self.from_combo)

        picker_row.addWidget(QLabel("تا تاریخ:"))
        self.to_combo = QComboBox()
        picker_row.addWidget(self.to_combo)

        show_btn = QPushButton("نمایش گزارش")
        show_btn.setCursor(Qt.PointingHandCursor)
        show_btn.clicked.connect(self._show_report)
        picker_row.addWidget(show_btn)
        layout.addLayout(picker_row)

        # --- Quick shortcuts ---
        quick_row = QHBoxLayout()
        today_btn = QPushButton("امروز")
        today_btn.clicked.connect(self._select_today)
        quick_row.addWidget(today_btn)

        last7_btn = QPushButton("۷ روز اخیر")
        last7_btn.clicked.connect(lambda: self._select_last_n_days(7))
        quick_row.addWidget(last7_btn)

        last30_btn = QPushButton("۳۰ روز اخیر")
        last30_btn.clicked.connect(lambda: self._select_last_n_days(30))
        quick_row.addWidget(last30_btn)

        all_btn = QPushButton("همه‌ی روزها")
        all_btn.clicked.connect(self._select_all)
        quick_row.addWidget(all_btn)

        quick_row.addStretch()
        layout.addLayout(quick_row)

        # --- Summary line ---
        self.summary_label = QLabel("")
        self.summary_label.setFont(QFont(app_font.APP_FONT_FAMILY, 11, QFont.Bold))
        layout.addWidget(self.summary_label)

        # --- Table ---
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["نام نوشیدنی", "تعداد فروش", "مبلغ کل (تومان)"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)

        # --- Footer actions ---
        footer = QHBoxLayout()
        export_btn = QPushButton("خروجی اکسل از این گزارش")
        export_btn.setCursor(Qt.PointingHandCursor)
        export_btn.clicked.connect(self._export_report)
        footer.addWidget(export_btn)
        footer.addStretch()
        close_btn = QPushButton("بستن")
        close_btn.clicked.connect(self.accept)
        footer.addWidget(close_btn)
        layout.addLayout(footer)

    # ------------------------------------------------------------ data
    def _populate_date_combos(self):
        display_dates = [(d, date_utils.to_display(d)) for d in self.available_dates]
        for combo in (self.from_combo, self.to_combo):
            combo.clear()
            for raw, label in display_dates:
                combo.addItem(label, raw)
        # default: today if present, else the most recent day
        today = date_utils.today_str()
        default_index = len(display_dates) - 1
        for i, (raw, _label) in enumerate(display_dates):
            if raw == today:
                default_index = i
        self.from_combo.setCurrentIndex(default_index)
        self.to_combo.setCurrentIndex(default_index)

    def _select_today(self):
        today = date_utils.today_str()
        self._select_range(today, today)

    def _select_last_n_days(self, n):
        if not self.available_dates:
            return
        start = self.available_dates[max(0, len(self.available_dates) - n)]
        end = self.available_dates[-1]
        self._select_range(start, end)

    def _select_all(self):
        self._select_range(self.available_dates[0], self.available_dates[-1])

    def _select_range(self, start, end):
        for combo, value in ((self.from_combo, start), (self.to_combo, end)):
            idx = combo.findData(value)
            if idx >= 0:
                combo.setCurrentIndex(idx)
        self._show_report()

    def _show_report(self):
        date_from = self.from_combo.currentData()
        date_to = self.to_combo.currentData()
        if date_from is None or date_to is None:
            return
        if date_from > date_to:
            date_from, date_to = date_to, date_from

        report = excel_manager.get_report(date_from, date_to)
        self._last_report = report
        self._last_range = (date_from, date_to)

        n_days = len(report["days_included"])
        self.summary_label.setText(
            f"در بازه {date_utils.to_display(date_from)} تا {date_utils.to_display(date_to)} "
            f"({n_days} روز دارای فروش) — "
            f"مجموع فروش: {report['total_count']} عدد، "
            f"مجموع مبلغ: {int(report['total_amount']):,} تومان"
        )

        rows = [item for item in report["items"] if item["count"] > 0]
        self.table.setRowCount(len(rows))
        for r, item in enumerate(rows):
            self.table.setItem(r, 0, QTableWidgetItem(item["name"]))
            count_item = QTableWidgetItem(str(item["count"]))
            count_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 1, count_item)
            total_item = QTableWidgetItem(f"{int(item['total']):,}")
            total_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 2, total_item)

    def _export_report(self):
        if not self._last_range:
            QMessageBox.information(self, "گزارش", "ابتدا یک بازه تاریخی را نمایش دهید.")
            return
        date_from, date_to = self._last_range
        default_name = f"گزارش فروش {date_from} تا {date_to}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره گزارش", os.path.join(os.path.expanduser("~"), default_name),
            "Excel Files (*.xlsx)"
        )
        if not path:
            return
        try:
            excel_manager.export_report(date_from, date_to, path)
            QMessageBox.information(self, "گزارش", f"گزارش با موفقیت ذخیره شد:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "خطا در ذخیره گزارش", str(exc))
