"""
Application dialogs — About, New Presentation, etc.
"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QVBoxLayout,
)
from PyQt5.QtGui import QFont, QPixmap


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PPEditer 정보")
        self.setFixedSize(420, 280)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 24, 24, 20)

        # App name
        title = QLabel("PPEditer")
        font = QFont()
        font.setPointSize(20)
        font.setBold(True)
        title.setFont(font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Sub-title
        sub = QLabel("Planning Proposal Editor")
        sub_font = QFont()
        sub_font.setPointSize(10)
        sub.setFont(sub_font)
        sub.setAlignment(Qt.AlignCenter)
        layout.addWidget(sub)

        layout.addSpacerItem(QSpacerItem(0, 8, QSizePolicy.Minimum, QSizePolicy.Fixed))

        # Version
        ver = QLabel("버전 1.0.0")
        ver.setAlignment(Qt.AlignCenter)
        layout.addWidget(ver)

        # Company / author
        info_lines = [
            "회사: HANDTECH (핸텍) — 상상공작소",
            "저작권자: 노진문 (Noh JinMoon)",
            "라이선스: MIT License",
            "Copyright © 2026 Noh JinMoon",
        ]
        for line in info_lines:
            lbl = QLabel(line)
            lbl.setAlignment(Qt.AlignCenter)
            layout.addWidget(lbl)

        layout.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Close button
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok)
        btn_box.accepted.connect(self.accept)
        layout.addWidget(btn_box)
