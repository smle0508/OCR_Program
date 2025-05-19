# ui/tab_extract.py
# 이 파일은 데이터 추출 탭을 정의하며, 선택한 PDF의 모든 페이지를 스캔하여
# ROI에 매핑된 영역의 텍스트를 추출하고 엑셀로 저장하는 기능을 제공합니다.

from __future__ import annotations
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QComboBox
)
from PySide6.QtCore import Qt

from core.excel_writer import ExcelWriter
from core.models import ROI, ROISet
from openpyxl.utils.cell import column_index_from_string
from core.ocr_engine import OCREngine


class TabExtract(QWidget):
    def __init__(self, roi_mgr: ROIManager, ex_mgr: ExclusionManager):
        super().__init__()
        self.roi_mgr = roi_mgr
        self.ex_mgr = ex_mgr
        self.ocr = OCREngine()
        self.pdf_paths: list[Path] = []

        self._init_ui()
        # ROI 세트가 변경되면 매핑 테이블 갱신
        self.set_selector.currentIndexChanged.connect(self._populate_mapping)
        # 초기 매핑
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

        # ROI ↔ Excel 열 매핑 테이블
        self.map_table = QTableWidget()
        self.map_table.setColumnCount(2)
        self.map_table.setHorizontalHeaderLabels(["ROI 이름", "열 지정 (예: A, B, C)"])
        layout.addWidget(self.map_table)

        # 추출 실행 버튼
        btn_run = QPushButton("추출 시작")
        btn_run.clicked.connect(self.on_run)
        layout.addWidget(btn_run, alignment=Qt.AlignRight)

    def on_select_pdf(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "PDF 파일 선택", "", "PDF Files (*.pdf)"
        )
        if paths:
            self.pdf_paths = [Path(p) for p in paths]
            self.pdf_label.setText(f"선택된 PDF: {len(paths)}개")

    def on_run(self) -> None:
        """OCR 실행 후 엑셀에 결과 저장"""
        if not self.pdf_paths:
            QMessageBox.warning(self, "경고", "PDF 파일을 먼저 선택하세요.")
            return

        # 저장할 엑셀 파일 선택
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "저장할 엑셀 파일 선택",
            "scanned_output.xlsx",
            "Excel Files (*.xlsx)"
        )
        if not save_path:
            return
        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'
        excel_path = Path(save_path)

        # 새 워크북 생성
        writer = ExcelWriter(excel_path, sheet_name="Sheet1")

        # 셀 매핑 정보 수집 (열 지정만)
        mapping: dict[str, int] = {}
        for row in range(self.map_table.rowCount()):
            name_item = self.map_table.item(row, 0)
            col_item = self.map_table.item(row, 1)
            if not name_item or not col_item:
                continue
            name = name_item.text().strip()
            col_text = col_item.text().strip().upper()
            if not col_text:
                continue
            if not col_text.isalpha():
                QMessageBox.warning(
                    self, "경고",
                    f"잘못된 열 지정: {col_text}\n열은 A, B, C 등 알파벳만 입력해주세요."
                )
                return
            try:
                mapping[name] = column_index_from_string(col_text)
            except Exception:
                QMessageBox.warning(
                    self, "경고",
                    f"열 변환 오류: {col_text}"
                )
                return

        # OCR 수행 및 Excel에 쓰기
        set_name = self.set_selector.currentText()
        roi_set = self.roi_mgr.get_set(set_name)
        rois = getattr(roi_set, 'rois', [])

        import fitz  # PyMuPDF
        row_counter = 1
        for pdf_path in self.pdf_paths:
            doc = fitz.open(pdf_path)
            for page_num in range(doc.page_count):
                # 페이지 단위 결과 저장 dict
                page_values: dict[tuple[int, int], str] = {}
                # 단일 필드 OCR (첫 행에만)
                for roi in rois:
                    if getattr(roi, 'field_type', 'single') == 'single' and roi.name in mapping:
                        text = self.ocr.extract_roi(
                            pdf_path,
                            page_num,
                            (roi.x, roi.y, roi.w, roi.h),
                            getattr(roi, 'tolerance', 0)
                        )
                        page_values[(row_counter, mapping[roi.name])] = text
                # 표 형식 OCR
                max_rows = 0
                for roi in rois:
                    if getattr(roi, 'field_type', '') == 'table' and roi.name in mapping:
                        table = self.ocr.extract_table(
                            pdf_path,
                            page_num,
                            (roi.x, roi.y, roi.w, roi.h)
                        )
                        # 테이블 각 행, 열을 순차적으로 해당 열부터 배치
                        for i, row in enumerate(table):
                            for j, cell_text in enumerate(row):
                                page_values[(row_counter + i, mapping[roi.name] + j)] = cell_text
                        max_rows = max(max_rows, len(table))
                # 결과 워크북에 기록
                writer.write_values(page_values)
                # 다음 페이지 위치 이동
                row_counter += max(max_rows, 1)

        # 저장 및 완료 메시지
        writer.save()
        QMessageBox.information(
            self,
            "완료",
            f"엑셀 파일이 생성되었습니다:\n{excel_path}"
        )

    def _populate_mapping(self) -> None:
        """
        선택된 ROI 세트의 ROI 리스트로 매핑 테이블 초기화
        """
        set_name = self.set_selector.currentText()
        roi_set = self.roi_mgr.get_set(set_name)
        rois = getattr(roi_set, 'rois', [])

        self.map_table.setRowCount(len(rois))
        for i, roi in enumerate(rois):
            name_item = QTableWidgetItem(roi.name)
            name_item.setFlags(Qt.ItemIsEnabled)
            self.map_table.setItem(i, 0, name_item)
            cell_item = QTableWidgetItem("")
            self.map_table.setItem(i, 1, cell_item)

    def refresh_sets(self) -> None:
        """
        외부에서 ROI 세트 변경 시 호출
        """
        self.set_selector.clear()
        self.set_selector.addItems(self.roi_mgr.list_sets())
        self._populate_mapping()
