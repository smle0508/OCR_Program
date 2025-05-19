# ui/roi_dialog.py
"""
ROIDialog
─────────
• ROI 생성 시 이름 · 오차(px) · 필드 유형(single/table) 입력 팝업
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog, QFormLayout, QLineEdit, QSpinBox,
    QComboBox, QDialogButtonBox
)


class ROIDialog(QDialog):
    """ROI 속성 입력 대화상자"""

    def __init__(self, default_name: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("ROI 속성 입력")

        self.edt_name = QLineEdit(default_name)
        self.spn_tol = QSpinBox(); self.spn_tol.setRange(0, 50); self.spn_tol.setValue(3)
        self.cmb_type = QComboBox(); self.cmb_type.addItems(["single", "table"])

        form = QFormLayout(self)
        form.addRow("이름:", self.edt_name)
        form.addRow("오차(px):", self.spn_tol)
        form.addRow("필드 유형:", self.cmb_type)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        form.addWidget(btns)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

    # 결과 편의 메서드
    def values(self):
        return (
            self.edt_name.text().strip(),
            self.spn_tol.value(),
            self.cmb_type.currentText()
        )
