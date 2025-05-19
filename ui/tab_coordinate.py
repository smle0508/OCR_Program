# ui/tab_coordinate.py

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton,
    QInputDialog, QFileDialog
)

from ui.pdf_viewer import PDFViewer
from core.roi_manager import ROIManager
from core.models import ROI, ROISet


class TabCoordinate(QWidget):
    """
    PDF 좌표(ROI) 지정 및 관리 탭
    -------------------------------------------------
    • 상단: PDF 불러오기 버튼 + 좌표 세트 드롭다운 + 세트 삭제 버튼 + 세트 저장 버튼
    • 본문: PDFViewer (ROI 드래그·편집)

    좌표 세트 저장 시 ROIManager에 영구 보관되며, TabEdit/TabExtract에 실시간 반영됩니다.
    """

    set_changed = Signal(str)
    sets_updated = Signal()

    def __init__(self, mgr: ROIManager) -> None:
        super().__init__()
        self.roi_mgr = mgr

        # 위젯 생성
        self.btn_load_pdf   = QPushButton("PDF 불러오기")
        self.viewer         = PDFViewer(self)
        self.cmb_set        = QComboBox()
        self.btn_delete_set = QPushButton("삭제")
        self.btn_save_set   = QPushButton("저장")

        # 좌표 세트 드롭다운 초기화
        self._refresh_set_list()

        # 레이아웃 설정
        control_layout = QHBoxLayout()
        control_layout.addWidget(self.btn_load_pdf)
        control_layout.addWidget(QLabel("좌표 세트:"))
        control_layout.addWidget(self.cmb_set)
        control_layout.addWidget(self.btn_delete_set)
        control_layout.addWidget(self.btn_save_set)
        control_layout.addStretch()

        main_layout = QVBoxLayout(self)
        main_layout.addLayout(control_layout)
        main_layout.addWidget(self.viewer, 1)

        # 시그널 연결
        self.btn_load_pdf.clicked.connect(self.on_load_pdf)
        self.cmb_set.currentTextChanged.connect(self.on_set_changed)
        self.btn_delete_set.clicked.connect(self.on_delete_set)
        self.btn_save_set.clicked.connect(self.on_save_set)

    def _refresh_set_list(self) -> None:
        """ROIManager의 세트 리스트로 콤보박스 업데이트"""
        self.cmb_set.blockSignals(True)
        self.cmb_set.clear()
        for name in self.roi_mgr.list_sets():
            self.cmb_set.addItem(name)
        self.cmb_set.blockSignals(False)

    def on_load_pdf(self) -> None:
        """PDF 파일 선택 및 로드"""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "PDF 파일 선택",
            "",
            "PDF Files (*.pdf)"
        )
        if path:
            self.viewer.load_pdf(path)

    def on_set_changed(self, name: str) -> None:
        """세트 변경 시 호출: PDFViewer에 로드"""
        if name:
            self.roi_mgr.load_set(name)
            self.viewer.refresh()
            self.set_changed.emit(name)
            self.sets_updated.emit()

    def on_delete_set(self) -> None:
        """현재 선택된 세트 삭제"""
        name = self.cmb_set.currentText()
        if name:
            self.roi_mgr.delete_set(name)
            self._refresh_set_list()
            self.set_changed.emit("")
            self.sets_updated.emit()

    def on_save_set(self) -> None:
        """현재 ROI를 새로운 세트로 저장"""
        name, ok = QInputDialog.getText(
            self,
            "세트 저장",
            "새 세트 이름:"
        )
        if ok and name:
            # 현재 PDFViewer에 지정된 ROIItem들을 모델로 변환
            rois: list[ROI] = []
            for item in self.viewer.export_rois().values():
                rect = item.rect()
                x = int(rect.x())
                y = int(rect.y())
                w = int(rect.width())
                h = int(rect.height())
                rois.append(
                    ROI(
                        name=item.name,
                        x=x, y=y,
                        w=w, h=h,
                        tolerance=item.tolerance,
                        field_type=item.field_type
                    )
                )
            roi_set = ROISet(set_name=name, rois=rois)
            self.roi_mgr.upsert_set(roi_set)
            self._refresh_set_list()
            self.cmb_set.setCurrentText(name)
            self.set_changed.emit(name)
            self.sets_updated.emit()
