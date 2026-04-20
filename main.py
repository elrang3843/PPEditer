"""
PPEditer - Planning Proposal Editor
Copyright (c) 2026 Noh JinMoon (HANDTECH - 상상공작소)
MIT License
"""
import sys
import os


def main():
    from PyQt5.QtWidgets import QApplication, QMessageBox
    from PyQt5.QtCore import Qt, QCoreApplication

    QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QCoreApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("PPEditer")
    app.setApplicationDisplayName("PPEditer")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("HANDTECH")
    app.setOrganizationDomain("handtech.co.kr")

    try:
        import pptx  # noqa: F401
    except ImportError:
        QMessageBox.critical(
            None,
            "오류 - PPEditer",
            "python-pptx 라이브러리가 설치되지 않았습니다.\n\n"
            "터미널에서 다음 명령을 실행해주세요:\n"
            "  pip install python-pptx",
        )
        return 1

    app.setStyle("Fusion")

    from src.ui.main_window import MainWindow

    window = MainWindow()

    if len(sys.argv) > 1:
        filepath = sys.argv[1]
        if os.path.exists(filepath) and filepath.lower().endswith((".ppt", ".pptx")):
            window.open_file(filepath)

    window.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
