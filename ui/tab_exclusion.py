# ui/tab_exclusion.py
"""
TabExclusion (단순화 + 엑셀 불러오기)
────────────────────────────────────
• exclude_strings = 전역 제외 문자열
• 수동 추가·삭제 + 엑셀 한꺼번에 추가(A열)
"""

from __future__ import annotations

from pathlib import Path
from typing import List

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QFileDialog
)
from PySide6.QtCore import Qt

from core.exclusion_manager import ExclusionManager


class TabExclusion(QWidget):
    def __init__(self, mgr: ExclusionManager):
        super().__init__()
        self.mgr = mgr

        # ─ 입력 UI ─
        self.edt_text = QLineEdit()
        self.btn_add  = QPushButton("추가")
        self.btn_del  = QPushButton("삭제")
        self.btn_load = QPushButton("엑셀 불러오기")   # NEW

        form = QHBoxLayout()
        form.addWidget(QLabel("제외 텍스트:"))
        form.addWidget(self.edt_text, 1)
        form.addWidget(self.btn_add)
        form.addWidget(self.btn_del)
        form.addWidget(self.btn_load)

        self.table = QTableWidget()
        self.table.setColumnCount(1)
        self.table.setHorizontalHeaderLabels(["저장된 제외 텍스트"])

        v = QVBoxLayout(self)
        v.addLayout(form)
        v.addWidget(self.table)

        # 시그널
        self.btn_add.clicked.connect(self.add_text)
        self.btn_del.clicked.connect(self.delete_selected)
        self.btn_load.clicked.connect(self.load_excel)

        self.populate_table()

    # ───────────────────────────
    def populate_table(self):
        items = self.mgr.list_all()
        self.table.setRowCount(len(items))
        for row, txt in enumerate(items):
            item = QTableWidgetItem(txt)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, item)

    # ───────────────────────────
    def add_text(self):
        txt = self.edt_text.text().strip()
        if not txt:
            return
        self.mgr.add_many([txt])
        self.edt_text.clear()
        self.populate_table()

    def delete_selected(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        for r in rows:
            txt = self.table.item(r, 0).text()
            self.mgr.remove(txt)
            self.table.removeRow(r)

    # ─ 엑셀 대량 추가 ─
    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 선택", "", "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if not path:
            return
        try:
            import pandas as pd
            df = pd.read_excel(Path(path), header=None, usecols=[0])
            texts = df.iloc[:, 0].dropna().astype(str).tolist()
            self.mgr.add_many(texts)
            QMessageBox.information(self, "완료", f"{len(texts)}개 텍스트를 추가했습니다.")
            self.populate_table()
        except Exception as e:
            QMessageBox.critical(self, "오류", f"엑셀 읽기 실패:\n{e}")
