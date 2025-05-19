# ui/tab_extract.py
"""
TabExtract
──────────
• PDF 파일을 로드하여 OCR을 수행하고 새 엑셀 파일에 결과를 저장하는 탭
• ROI 이름 ↔ Excel 열 매핑 테이블 제공
"""
from __future__ import annotations
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QComboBox
)
from PySide6.QtCore import Qt

from core.ocr_engine import OCREngine
from core.excel_writer import ExcelWriter
from core.roi_manager import ROIManager
from core.exclusion_manager import ExclusionManager


class TabExtract(QWidget):
    def __init__(self, roi_mgr: ROIManager, ex_mgr: ExclusionManager):
        super().__init__()
        self.roi_mgr = roi_mgr
        self.ex_mgr = ex_mgr
        self.ocr = OCREngine()

        self.pdf_paths: list[Path] = []
        self.excel_path: Path | None = None

        self._init_ui()
        # ROI 세트 변경 시 매핑 재구성
        self.set_selector.currentIndexChanged.connect(self._populate_mapping)
        self._populate_mapping()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)

        # ROI 세트 선택
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("ROI 세트:"), alignment=Qt.AlignVCenter)
        self.set_selector = QComboBox()
        self.set_selector.addItems(self.roi_mgr.list_sets())
        h1.addWidget(self.set_selector)
        layout.addLayout(h1)

        # PDF 파일 선택
        h2 = QHBoxLayout()
        self.pdf_label = QLabel("선택된 PDF: 없음")
        btn_pdf = QPushButton("PDF 파일 선택")
        btn_pdf.clicked.connect(self.on_select_pdf)
        h2.addWidget(self.pdf_label)
        h2.addWidget(btn_pdf)
        layout.addLayout(h2)

        # 새 엑셀 파일 저장 경로 지정
        h3 = QHBoxLayout()
        self.excel_label = QLabel("새 엑셀 파일: 없음")
        btn_excel = QPushButton("새 엑셀 파일로 저장")
        btn_excel.clicked.connect(self.on_select_excel)
        h3.addWidget(self.excel_label)
        h3.addWidget(btn_excel)
        layout.addLayout(h3)

        # ROI ↔ Excel 열 매핑 테이블
        self.map_table = QTableWidget(0, 2)
        self.map_table.setHorizontalHeaderLabels(["ROI 이름", "열 주소 (예: A1)"])
        layout.addWidget(self.map_table)

        # 실행 버튼
        btn_run = QPushButton("OCR → Excel 실행")
        btn_run.clicked.connect(self.on_run)
        layout.addWidget(btn_run, alignment=Qt.AlignRight)

    def on_select_pdf(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "PDF 파일 선택", "", "PDF Files (*.pdf)"
        )
        if paths:
            self.pdf_paths = [Path(p) for p in paths]
            self.pdf_label.setText(f"선택된 PDF: {len(paths)}개")

    def on_select_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "새 엑셀 파일로 저장",
            "scanned_output.xlsx",
            "Excel Files (*.xlsx)"
        )
        if path:
            if not path.lower().endswith('.xlsx'):
                path += '.xlsx'
            self.excel_path = Path(path)
            self.excel_label.setText(f"새 엑셀 파일: {self.excel_path.name}")

    def _populate_mapping(self) -> None:
        """
        선택된 ROI 세트의 ROI 리스트로 매핑 테이블 초기화
        """
        set_name = self.set_selector.currentText()
        roi_set = self.roi_mgr.get_set(set_name)
        if not roi_set:
            return

        rois = getattr(roi_set, 'rois', [])
        self.map_table.setRowCount(len(rois))
        for i, roi in enumerate(rois):
            # ROI 이름
            name_item = QTableWidgetItem(roi.name)
            name_item.setFlags(Qt.ItemIsEnabled)
            self.map_table.setItem(i, 0, name_item)
            # 기본 열 주소 빈 값
            cell_item = QTableWidgetItem("")
            self.map_table.setItem(i, 1, cell_item)

    def on_run(self) -> None:
        # 입력 검증
        if not self.pdf_paths:
            QMessageBox.warning(self, "경고", "PDF 파일을 먼저 선택하세요.")
            return
        if not self.excel_path:
            QMessageBox.warning(self, "경고", "새 엑셀 파일을 지정하세요.")
            return

        # 새 워크북 생성
        writer = ExcelWriter(self.excel_path, sheet_name="Sheet1")

        # 매핑 정보 수집
        mapping: dict[str, str] = {}
        for row in range(self.map_table.rowCount()):
            roi_name = self.map_table.item(row, 0).text()
            cell_addr = self.map_table.item(row, 1).text().strip()
            if cell_addr:
                mapping[roi_name] = cell_addr

        set_name = self.set_selector.currentText()
        roi_set = self.roi_mgr.get_set(set_name)
        rois = getattr(roi_set, 'rois', [])

        # OCR 수행 및 엑셀 기록
        for pdf_path in self.pdf_paths:
            for roi in rois:
                if roi.name not in mapping:
                    continue
                # OCR 추출 (page 0)
                box = (roi.x, roi.y, roi.w, roi.h)
                text = self.ocr.extract_roi(pdf_path, 0, box, getattr(roi, 'tolerance', 0))
                # 기록
                writer.write_values({mapping[roi.name]: text})

        # 저장 및 알림
        writer.save()
        QMessageBox.information(
            self, "완료", f"새 엑셀 파일이 생성되었습니다:\n{self.excel_path}"
        )

    def refresh_sets(self) -> None:
        """
        외부에서 ROI 세트 변경 시 호출
        """
        self.set_selector.clear()
        self.set_selector.addItems(self.roi_mgr.list_sets())
        self._populate_mapping()
