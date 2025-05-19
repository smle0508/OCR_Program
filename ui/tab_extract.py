# ui/tab_extract.py
"""
TabExtract  (PDF → OCR → Excel 탭, Drag&Drop 지원)
───────────────────────────────────────────────────
• ROI 세트 선택
• PDF 파일(여러 개) 선택  ─ 버튼·Drag&Drop 모두 지원
• 엑셀 파일 + 시트 선택  ─ 버튼·Drag&Drop 모두 지원
• 매핑 테이블에서 ‘엑셀 열’ 직접 입력
• 전역 제외 문자열 적용 후 원본 보존 + 새 파일로 저장
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List

from PySide6.QtCore import Slot
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QTableWidget, QTableWidgetItem, QMessageBox, QFileDialog
)

import fitz  # PyMuPDF
from openpyxl import load_workbook

from core.roi_manager import ROIManager
from core.exclusion_manager import ExclusionManager
from core.ocr_engine import OCREngine


class TabExtract(QWidget):
    def __init__(
        self,
        roi_mgr: ROIManager,
        ex_mgr: ExclusionManager,
        parent=None
    ):
        super().__init__(parent)
        self.roi_mgr = roi_mgr
        self.ex_mgr = ex_mgr
        self.ocr = OCREngine(dpi=600, lang="kor+eng")
        self.pdf_paths: List[Path] = []
        self.excel_path: Path | None = None

        # 위젯 초기화
        self.cmb_set = QComboBox()
        self.cmb_set.currentTextChanged.connect(self.populate_mapping)
        self.btn_pdf = QPushButton("PDF 선택")
        self.lbl_pdf = QLabel("선택 없음")
        self.btn_excel = QPushButton("엑셀 선택")
        self.cmb_sheet = QComboBox()
        self.table_map = QTableWidget(0, 2)
        self.table_map.setHorizontalHeaderLabels(["필드", "엑셀 열"])
        self.btn_run = QPushButton("변환 실행")

        # 버튼 시그널 연결
        self.btn_pdf.clicked.connect(self.select_pdfs)
        self.btn_excel.clicked.connect(self.select_excel)
        self.btn_run.clicked.connect(self.run_conversion)

        # 레이아웃 구성
        layout = QVBoxLayout(self)
        # ROI 세트
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("ROI 세트:"))
        row1.addWidget(self.cmb_set)
        layout.addLayout(row1)
        # PDF 선택
        row2 = QHBoxLayout()
        row2.addWidget(self.btn_pdf)
        row2.addWidget(self.lbl_pdf)
        layout.addLayout(row2)
        # 엑셀 선택
        row3 = QHBoxLayout()
        row3.addWidget(self.btn_excel)
        row3.addWidget(self.cmb_sheet)
        layout.addLayout(row3)
        # 매핑 테이블
        layout.addWidget(self.table_map)
        # 실행 버튼
        layout.addWidget(self.btn_run)

        # 초기 세트 로드
        self.refresh_sets()

    def select_pdfs(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "PDF 파일 선택", "", "PDF Files (*.pdf)"
        )
        if paths:
            self.pdf_paths = [Path(p) for p in paths]
            self.lbl_pdf.setText(f"{len(paths)}개 선택")

    def select_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if path:
            self.excel_path = Path(path)
            self.cmb_sheet.clear()
            wb = load_workbook(self.excel_path)
            self.cmb_sheet.addItems(wb.sheetnames)

    @Slot()
    def refresh_sets(self) -> None:
        names = self.roi_mgr.list_sets()
        self.cmb_set.blockSignals(True)
        self.cmb_set.clear()
        self.cmb_set.addItems(names)
        self.cmb_set.blockSignals(False)
        if names:
            self.populate_mapping(self.cmb_set.currentText())

    @Slot(str)
    def populate_mapping(self, set_name: str) -> None:
        self.table_map.setRowCount(0)
        rs = self.roi_mgr.get_set(set_name)
        if not rs:
            return
        for roi in rs.rois:
            row = self.table_map.rowCount()
            self.table_map.insertRow(row)
            self.table_map.setItem(row, 0, QTableWidgetItem(roi.name))
            self.table_map.setItem(row, 1, QTableWidgetItem(""))

    def run_conversion(self) -> None:
        set_name = self.cmb_set.currentText()
        rs = self.roi_mgr.get_set(set_name)
        if not rs:
            QMessageBox.warning(self, "실행 불가", "ROI 세트 선택 필요")
            return
        if not self.pdf_paths or not self.excel_path:
            QMessageBox.warning(self, "실행 불가", "PDF/엑셀 선택 필요")
            return

        mapping: Dict[str, str] = {
            self.table_map.item(r, 0).text(): self.table_map.item(r, 1).text()
            for r in range(self.table_map.rowCount())
        }

        wb = load_workbook(self.excel_path)
        ws = wb[self.cmb_sheet.currentText()]

        last = max(
            (next((i for i in range(ws.max_row, 0, -1)
                    if ws.cell(i, ws[f"{c}1"].column).value not in (None, "")), 0)
             for c in mapping.values()),
            default=0
        ) + 1

        for pdf in self.pdf_paths:
            doc = fitz.open(pdf)
            for p in range(doc.page_count):
                for roi in rs.rois:
                    # 디버그: 실제 크롭 좌표 및 페이지 정보 출력
                    print(f"[ExtractTab] Cropping PDF={pdf.name}, page={p}, "
                          f"x={roi.x}, y={roi.y}, w={roi.w}, h={roi.h}, tol={roi.tolerance}")
                    text = self.ocr.extract_roi(
                        pdf, p, (roi.x, roi.y, roi.w, roi.h), roi.tolerance
                    )
                    ws[f"{mapping[roi.name]}{last}"] = text
                last += 1

        wb.save(self.excel_path)
        QMessageBox.information(self, "완료", "변환이 완료되었습니다.")
