# ui/main_window.py
"""
MainWindow
──────────
• 4개 탭을 하나의 창으로 통합
    1) 좌표 지정      (TabCoordinate)
    2) 좌표 편집      (TabEdit)
    3) 제외 규칙 관리 (TabExclusion)
    4) 데이터 추출    (TabExtract)
• 좌표 세트 변경 시 TabEdit/TabExtract에 실시간 반영
• 프로그램 종료 시 확인 팝업
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QWidget, QVBoxLayout, QApplication, QMessageBox
)

from core.roi_manager import ROIManager
from core.exclusion_manager import ExclusionManager

from ui.tab_coordinate import TabCoordinate
from ui.tab_edit       import TabEdit
from ui.tab_exclusion  import TabExclusion
from ui.tab_extract    import TabExtract


def main():
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


class MainWindow(QMainWindow):
    """애플리케이션 메인 윈도우"""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("PDF OCR → Excel 변환기 by Moon")
        self.resize(1360, 900)

        # ─ 매니저 초기화 ─
        self.roi_mgr = ROIManager()
        self.ex_mgr  = ExclusionManager()

        # ─ 탭 인스턴스 생성 ─
        coord_tab     = TabCoordinate(self.roi_mgr)
        edit_tab      = TabEdit(self.roi_mgr)
        exclusion_tab = TabExclusion(self.ex_mgr)
        extract_tab   = TabExtract(self.roi_mgr, self.ex_mgr)

        # ─ 실시간 세트 변경 연동 ─
        coord_tab.sets_updated.connect(edit_tab.refresh_sets)
        coord_tab.sets_updated.connect(extract_tab.refresh_sets)

        # ─ 탭 위젯 구성 ─
        tabs = QTabWidget()
        tabs.addTab(coord_tab,     "좌표 지정")
        tabs.addTab(edit_tab,      "좌표 편집")
        tabs.addTab(exclusion_tab, "제외 규칙")
        tabs.addTab(extract_tab,   "데이터 추출")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.addWidget(tabs)
        self.setCentralWidget(container)

    # ───────────────────────────────
    # 종료 확인
    # ───────────────────────────────
    def closeEvent(self, event) -> None:
        reply = QMessageBox.question(
            self,
            "프로그램 종료",
            "프로그램을 종료하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    main()
