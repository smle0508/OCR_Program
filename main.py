# main.py
"""
엔트리 포인트
────────────
• 가상환경/전역 어디서든 실행 가능
• QApplication 생성 → MainWindow 호출
"""

import sys

from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)

    win = MainWindow()
    win.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
