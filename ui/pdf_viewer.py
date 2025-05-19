# ui/pdf_viewer.py
from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional

import fitz  # PyMuPDF
from PySide6.QtCore import Qt, QRect, QRectF, QPoint, QSize
from PySide6.QtGui import (
    QImage, QPixmap, QMouseEvent, QWheelEvent, QTransform,
    QPen, QColor, QContextMenuEvent, QAction, QPainter
)
from PySide6.QtWidgets import (
    QGraphicsView, QGraphicsScene, QGraphicsRectItem, QRubberBand,
    QGraphicsItem, QMenu, QInputDialog, QDialog
)

from ui.roi_dialog import ROIDialog


# ──────────────────────────────────────────────────
# 1. ROIItem  (사각형 + 오차 점선 + 유형 편집)
# ──────────────────────────────────────────────────
class ROIItem(QGraphicsRectItem):
    def __init__(
        self,
        rect_local: QRectF,
        name: str,
        tolerance: int = 3,
        field_type: str = "single"
    ):
        super().__init__(rect_local)
        self.name = name
        self.tolerance = tolerance
        self.field_type = field_type

        pen = QPen(QColor("magenta"), 1)
        self.setPen(pen)
        self.setBrush(Qt.transparent)

        self.setFlags(
            QGraphicsItem.ItemIsSelectable |
            QGraphicsItem.ItemIsMovable |
            QGraphicsItem.ItemSendsGeometryChanges
        )

        dot_pen = QPen(QColor("magenta"))
        dot_pen.setStyle(Qt.DotLine)
        self._tol_rect = QGraphicsRectItem(self)
        self._tol_rect.setPen(dot_pen)
        self._update_tolerance()

    def _update_tolerance(self):
        base = self.rect()
        t = self.tolerance
        self._tol_rect.setRect(
            QRectF(
                base.left() - t,
                base.top() - t,
                base.width() + 2 * t,
                base.height() + 2 * t,
            )
        )

    def itemChange(self, change, value):
        if change in (
            QGraphicsItem.ItemPositionChange,
            QGraphicsItem.ItemScaleChange,
            QGraphicsItem.ItemTransformChange,
        ):
            self._update_tolerance()
        return super().itemChange(change, value)

    def contextMenuEvent(self, e: QContextMenuEvent):
        menu = QMenu()
        act_rename = QAction("이름 수정", menu)
        act_tol    = QAction("오차(px) 수정", menu)
        act_type   = QAction("유형 수정", menu)
        act_del    = QAction("삭제", menu)
        menu.addActions([act_rename, act_tol, act_type, act_del])

        sel = menu.exec(e.screenPos())
        if sel is act_del:
            self.scene().removeItem(self)
        elif sel is act_rename:
            new, ok = QInputDialog.getText(
                None, "ROI 이름", "새 이름:", text=self.name
            )
            if ok and new.strip():
                self.name = new.strip()
        elif sel is act_tol:
            val, ok = QInputDialog.getInt(
                None,
                "오차(px)",
                "새 오차:",
                self.tolerance,
                0,
                50
            )
            if ok:
                self.tolerance = val
                self._update_tolerance()
        elif sel is act_type:
            items = ["단일 필드", "표 형식"]
            current = 0 if self.field_type == "single" else 1
            choice, ok = QInputDialog.getItem(
                None,
                "필드 유형 선택",
                "유형:",
                items,
                current,
                False
            )
            if ok and choice:
                self.field_type = "single" if choice == "단일 필드" else "table"


# ──────────────────────────────────────────────────
# 2. PDFViewer  (PDF + ROI 편집)
# ──────────────────────────────────────────────────
class PDFViewer(QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScene(QGraphicsScene(self))
        self._rb = QRubberBand(QRubberBand.Rectangle, self)
        self._origin: Optional[QPoint] = None

        self.roi_items: Dict[str, ROIItem] = {}
        self.scale_factor = 1.0
        self.dpi = 600  # 기본 DPI 설정 (ROI 지정과 OCR 모두 일치)

        # 안티앨리어싱 활성화
        self.setRenderHint(QPainter.Antialiasing, True)

    def load_pdf(self, path: str | Path):
        """
        PDF 파일을 지정된 DPI로 렌더링하고 뷰어에 표시합니다.
        ROI 지정과 OCR 추출 모두 동일한 DPI를 사용하기 위함입니다.
        """
        doc = fitz.open(str(path))
        zoom = self.dpi / 72
        pix  = doc[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pm  = QPixmap.fromImage(img)
        doc.close()

        # 장면 초기화 및 이미지 추가
        self.scene().clear()
        self.roi_items.clear()
        self.scene().addPixmap(pm)
        self.setSceneRect(pm.rect())

        # 변환 및 확대/축소 비율 초기화
        self.resetTransform()
        self.scale_factor = 1.0

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._origin = e.pos()
            self._rb.setGeometry(QRect(self._origin, QSize()))
            self._rb.show()
        else:
            super().mousePressEvent(e)

    def mouseMoveEvent(self, e: QMouseEvent):
        if self._origin:
            self._rb.setGeometry(QRect(self._origin, e.pos()).normalized())
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton and self._origin:
            self._rb.hide()
            rect = self._rb.geometry()
            if rect.width() > 5 and rect.height() > 5:
                tl_scene = self.mapToScene(rect.topLeft())
                br_scene = self.mapToScene(rect.bottomRight())
                width  = br_scene.x() - tl_scene.x()
                height = br_scene.y() - tl_scene.y()

                idx = 1
                while f"ROI_{idx}" in self.roi_items:
                    idx += 1
                default_name = f"ROI_{idx}"

                dlg = ROIDialog(default_name, self)
                if dlg.exec() != QDialog.Accepted:
                    self._origin = None
                    return
                name, tol, ftype = dlg.values()
                if not name:
                    name = default_name
                if name in self.roi_items:
                    name += "_new"

                rect_local = QRectF(0, 0, width, height)
                item = ROIItem(rect_local, name, tol, ftype)
                item.setPos(tl_scene)
                self.scene().addItem(item)
                self.roi_items[name] = item
            self._origin = None
        else:
            super().mouseReleaseEvent(e)

    def wheelEvent(self, e: QWheelEvent):
        if e.modifiers() & Qt.ControlModifier:
            self.scale_factor *= 1.2 if e.angleDelta().y() > 0 else 1 / 1.2
            t = QTransform()
            t.scale(self.scale_factor, self.scale_factor)
            self.setTransform(t)
        else:
            super().wheelEvent(e)

    def keyPressEvent(self, e):
        if e.modifiers() & Qt.ControlModifier and e.key() in (
            Qt.Key_Plus, Qt.Key_Equal, Qt.Key_Minus
        ):
            self.scale_factor *= 1.2 if e.key() in (
                Qt.Key_Plus, Qt.Key_Equal
            ) else 1 / 1.2
            t = QTransform()
            t.scale(self.scale_factor, self.scale_factor)
            self.setTransform(t)
        else:
            super().keyPressEvent(e)

    def export_rois(self) -> Dict[str, ROIItem]:
        """현재 지정된 ROI 항목을 사전 형태로 반환합니다."""
        return self.roi_items

    def clear_rois(self) -> None:
        """현재 ROI를 모두 제거합니다."""
        for item in list(self.roi_items.values()):
            self.scene().removeItem(item)
        self.roi_items.clear()
