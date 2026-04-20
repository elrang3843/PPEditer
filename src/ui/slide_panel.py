"""
Slide panel — left-side thumbnail list with drag-reorder and context menu.
"""
from PyQt5.QtCore import Qt, pyqtSignal, QSize, QMimeData, QPoint
from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QPen, QDrag
from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
    QMenu,
    QAction,
    QSizePolicy,
    QScrollArea,
)

THUMB_W = 160
THUMB_H = 90
ITEM_PAD = 12  # horizontal padding around thumbnail


class SlideThumbnailWidget(QWidget):
    """Custom widget drawn inside each list item."""

    def __init__(self, index: int, pixmap: QPixmap, parent=None):
        super().__init__(parent)
        self.index = index
        self.pixmap = pixmap
        self.setFixedSize(THUMB_W + ITEM_PAD * 2, THUMB_H + 28)

    def set_pixmap(self, pixmap: QPixmap):
        self.pixmap = pixmap
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)

        # Slide thumbnail area
        thumb_rect = self.rect().adjusted(ITEM_PAD, 4, -ITEM_PAD, -22)

        # Shadow
        shadow = thumb_rect.adjusted(2, 2, 2, 2)
        painter.fillRect(shadow, QColor(0, 0, 0, 60))

        # White background
        painter.fillRect(thumb_rect, Qt.white)

        # Thumbnail image
        if self.pixmap and not self.pixmap.isNull():
            painter.drawPixmap(thumb_rect, self.pixmap)

        # Slide border
        painter.setPen(QPen(QColor(180, 180, 180), 1))
        painter.drawRect(thumb_rect)

        # Slide number label
        num_rect = self.rect().adjusted(0, self.height() - 20, 0, 0)
        painter.setPen(Qt.black)
        font = QFont("맑은 고딕", 8)
        painter.setFont(font)
        painter.drawText(num_rect, Qt.AlignCenter, str(self.index + 1))


class SlidePanel(QWidget):
    """Left panel showing slide thumbnails."""

    slide_selected = pyqtSignal(int)           # emitted when user clicks a thumbnail
    slide_reorder_requested = pyqtSignal(int, int)  # (from_index, to_index)
    add_slide_requested = pyqtSignal(int)      # after_index
    delete_slide_requested = pyqtSignal(int)
    duplicate_slide_requested = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumWidth(THUMB_W + ITEM_PAD * 2 + 20)
        self.setMaximumWidth(220)
        self._model = None
        self._current_index = 0
        self._build_ui()

    # ── UI setup ───────────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header = QLabel("  슬라이드")
        header.setFixedHeight(28)
        header.setStyleSheet(
            "background:#f0f0f0; border-bottom:1px solid #ccc; font-size:11px; color:#444;"
        )
        layout.addWidget(header)

        self.list_widget = QListWidget()
        self.list_widget.setViewMode(QListWidget.ListMode)
        self.list_widget.setIconSize(QSize(THUMB_W, THUMB_H))
        self.list_widget.setSpacing(2)
        self.list_widget.setDragDropMode(QAbstractItemView.InternalMove)
        self.list_widget.setDefaultDropAction(Qt.MoveAction)
        self.list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.list_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_widget.setStyleSheet(
            "QListWidget { background:#e8e8e8; border:none; outline:none; }"
            "QListWidget::item { border:none; background:transparent; }"
            "QListWidget::item:selected { background:#cde6ff; border:1px solid #0078d4; }"
        )
        self.list_widget.currentRowChanged.connect(self._on_row_changed)
        self.list_widget.customContextMenuRequested.connect(self._show_context_menu)
        self.list_widget.model().rowsMoved.connect(self._on_rows_moved)
        layout.addWidget(self.list_widget)

    # ── Public methods ─────────────────────────────────────────────────

    def refresh(self, model, current_index: int = 0):
        """Rebuild thumbnails from *model*."""
        self._model = model
        self._current_index = current_index

        self.list_widget.blockSignals(True)
        self.list_widget.clear()

        if model and model.is_open:
            for i in range(model.slide_count):
                self._append_item(i, model)

        self.list_widget.blockSignals(False)
        self._set_current(current_index)

    def update_thumbnail(self, index: int):
        """Re-render a single thumbnail."""
        if self._model is None or not self._model.is_open:
            return
        item = self.list_widget.item(index)
        if item is None:
            return
        w = self.list_widget.itemWidget(item)
        if w is None:
            return
        pixmap = self._render_thumb(index)
        w.set_pixmap(pixmap)

    def set_current(self, index: int):
        self._current_index = index
        self._set_current(index)

    # ── Internal helpers ───────────────────────────────────────────────

    def _append_item(self, index: int, model):
        from src.core.slide_renderer import render_slide
        item = QListWidgetItem()
        thumb_widget = SlideThumbnailWidget(index, self._render_thumb(index, model))
        item.setSizeHint(thumb_widget.sizeHint())
        self.list_widget.addItem(item)
        self.list_widget.setItemWidget(item, thumb_widget)

    def _render_thumb(self, index: int, model=None) -> QPixmap:
        if model is None:
            model = self._model
        if model is None:
            return QPixmap()
        from src.core.slide_renderer import render_slide
        slide = model.get_slide(index)
        if slide is None:
            return QPixmap()
        try:
            return render_slide(slide, THUMB_W, THUMB_H,
                                model.slide_width, model.slide_height)
        except Exception:
            return QPixmap()

    def _set_current(self, index: int):
        if 0 <= index < self.list_widget.count():
            self.list_widget.blockSignals(True)
            self.list_widget.setCurrentRow(index)
            self.list_widget.blockSignals(False)

    # ── Signals ────────────────────────────────────────────────────────

    def _on_row_changed(self, row: int):
        if row >= 0:
            self._current_index = row
            self.slide_selected.emit(row)

    def _on_rows_moved(self, parent, start, end, dest_parent, dest_row):
        from_idx = start
        to_idx = dest_row if dest_row > start else dest_row
        # Clamp
        to_idx = max(0, min(to_idx, (self._model.slide_count - 1) if self._model else 0))
        if from_idx != to_idx:
            self.slide_reorder_requested.emit(from_idx, to_idx)
        # Re-number widgets
        self._renumber_widgets()

    def _renumber_widgets(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            w = self.list_widget.itemWidget(item)
            if w:
                w.index = i
                w.update()

    def _show_context_menu(self, pos: QPoint):
        item = self.list_widget.itemAt(pos)
        idx = self.list_widget.row(item) if item else self.list_widget.currentRow()

        menu = QMenu(self)
        act_add = menu.addAction("새 슬라이드 추가(&N)")
        act_dup = menu.addAction("슬라이드 복제(&D)")
        menu.addSeparator()
        act_del = menu.addAction("슬라이드 삭제(&X)")

        act_add.setShortcut("Ctrl+M")

        if self._model and self._model.slide_count <= 1:
            act_del.setEnabled(False)

        action = menu.exec_(self.list_widget.mapToGlobal(pos))
        if action == act_add:
            self.add_slide_requested.emit(idx)
        elif action == act_dup:
            self.duplicate_slide_requested.emit(idx)
        elif action == act_del:
            self.delete_slide_requested.emit(idx)
