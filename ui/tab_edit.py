# ui/tab_edit.py
"""
TabEdit  (탭 2: 좌표 세트 편집 – 이름도 수정 가능 버전)
─────────────────────────────────────────────────────
• 세트 선택 → ROI 목록 표시
• 모든 셀(이름 포함) 직접 수정 → ‘적용’ 저장
• 이름 중복 검사 · 행 삭제 지원
"""

from __future__ import annotations

from typing import List

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QTableWidget, QTableWidgetItem, QPushButton, QMessageBox
)
from PySide6.QtCore import Qt

from core.models import ROI, ROISet
from core.roi_manager import ROIManager


class TabEdit(QWidget):
    """좌표 세트 편집 탭"""

    HEADERS = ["이름", "X", "Y", "W", "H", "오차", "유형"]

    def __init__(self, mgr: ROIManager):
        super().__init__()
        self.mgr = mgr

        # ─ UI 구성 ─
        self.cmb_sets  = QComboBox()
        self.table     = QTableWidget()
        self.btn_apply = QPushButton("적용")
        self.btn_del   = QPushButton("삭제")

        h = QHBoxLayout()
        h.addWidget(QLabel("좌표 세트:"))
        h.addWidget(self.cmb_sets)
        h.addStretch()
        h.addWidget(self.btn_apply)
        h.addWidget(self.btn_del)

        v = QVBoxLayout(self)
        v.addLayout(h)
        v.addWidget(self.table)

        # 테이블 헤더
        self.table.setColumnCount(len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)

        # 초기화
        self.refresh_sets()

        # 시그널
        self.cmb_sets.currentTextChanged.connect(self.populate_table)
        self.btn_apply.clicked.connect(self.apply_changes)
        self.btn_del.clicked.connect(self.delete_selected)

    # ───────────────────────────────
    # 1. 세트 콤보 갱신
    # ───────────────────────────────
    def refresh_sets(self):
        current = self.cmb_sets.currentText()
        self.cmb_sets.blockSignals(True)
        self.cmb_sets.clear()
        self.cmb_sets.addItems(self.mgr.list_sets())
        self.cmb_sets.setCurrentText(current)
        self.cmb_sets.blockSignals(False)

        if not current and self.cmb_sets.count():
            self.populate_table(self.cmb_sets.currentText())

    # ───────────────────────────────
    # 2. 테이블 채우기
    # ───────────────────────────────
    def populate_table(self, set_name: str):
        rs: ROISet | None = self.mgr.get_set(set_name)
        if not rs:
            self.table.setRowCount(0)
            return

        self.table.setRowCount(len(rs.rois))
        for row, r in enumerate(rs.rois):
            values = [r.name, r.x, r.y, r.w, r.h, r.tolerance, r.field_type]
            for col, val in enumerate(values):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() | Qt.ItemIsEditable)  # 모든 셀 편집 허용
                self.table.setItem(row, col, item)

    # ───────────────────────────────
    # 3. 변경 적용
    # ───────────────────────────────
    def apply_changes(self):
        set_name = self.cmb_sets.currentText()
        if not set_name:
            return

        # 이름 중복 검사
        names = [self.table.item(r, 0).text().strip()
                 for r in range(self.table.rowCount())]
        if len(names) != len(set(names)):
            QMessageBox.warning(self, "중복 이름", "ROI 이름이 중복되었습니다.")
            return

        rois: List[ROI] = []
        for row in range(self.table.rowCount()):
            try:
                name = self.table.item(row, 0).text().strip()
                x    = int(self.table.item(row, 1).text())
                y    = int(self.table.item(row, 2).text())
                w    = int(self.table.item(row, 3).text())
                h    = int(self.table.item(row, 4).text())
                tol  = int(self.table.item(row, 5).text())
                ftyp = self.table.item(row, 6).text().strip()
                rois.append(ROI(name, x, y, w, h, tol, ftyp))
            except Exception:
                QMessageBox.warning(self, "입력 오류",
                                    f"{row+1}행 자료가 올바르지 않습니다.")
                return

        self.mgr.upsert_set(ROISet(set_name, rois))
        QMessageBox.information(self, "저장 완료", "변경 사항이 저장되었습니다.")

    # ───────────────────────────────
    # 4. ROI 삭제
    # ───────────────────────────────
    def delete_selected(self):
        rows = sorted(
            {idx.row() for idx in self.table.selectedIndexes()},
            reverse=True
        )
        if not rows:
            return
        for row in rows:
            self.table.removeRow(row)
