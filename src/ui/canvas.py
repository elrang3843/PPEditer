"""
Slide canvas — main editing area with zoom, shape selection, and inline text editing.
"""
from PyQt5.QtCore import Qt, QRect, QPoint, pyqtSignal, QSize
from PyQt5.QtGui import (
    QColor,
    QFont,
    QPainter,
    QPen,
    QPixmap,
    QCursor,
    QKeySequence,
)
from PyQt5.QtWidgets import (
    QScrollArea,
    QWidget,
    QTextEdit,
    QSizePolicy,
    QApplication,
)

from src.core.slide_renderer import render_slide

MARGIN = 40          # pixels around the slide
MIN_ZOOM = 0.25
MAX_ZOOM = 3.0
ZOOM_STEP = 0.1


class SlideWidget(QWidget):
    """
    Renders one slide and handles:
      - shape selection (click)
      - inline text editing (double-click)
    """

    shape_selected = pyqtSignal(int)   # shape index (-1 = none)
    text_changed = pyqtSignal(int, str)  # (shape_index, new_text)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self._slide = None
        self._model = None
        self._slide_index = 0
        self._slide_w_emu = 9144000
        self._slide_h_emu = 5143500
        self._zoom = 1.0

        # Cached rendered pixmap
        self._pixmap: QPixmap | None = None
        self._dirty = True

        # Shape rects in *slide* coordinates (pixels at zoom=1 x 96dpi)
        self._shape_rects: list[QRect] = []
        self._selected_shape = -1

        # Inline text editor
        self._editor: QTextEdit | None = None
        self._editing_shape = -1

    # ── Public API ─────────────────────────────────────────────────────

    def set_slide(self, slide, model, slide_index: int):
        self._finish_editing(save=False)
        self._slide = slide
        self._model = model
        self._slide_index = slide_index
        if model:
            self._slide_w_emu = model.slide_width
            self._slide_h_emu = model.slide_height
        self._selected_shape = -1
        self._dirty = True
        self._update_size()
        self.update()

    def set_zoom(self, zoom: float):
        self._zoom = max(MIN_ZOOM, min(MAX_ZOOM, zoom))
        self._dirty = True
        self._update_size()
        self._reposition_editor()
        self.update()

    def zoom(self) -> float:
        return self._zoom

    def invalidate(self):
        """Force re-render on next paint."""
        self._dirty = True
        self.update()

    # ── Size helpers ───────────────────────────────────────────────────

    def _slide_display_size(self) -> tuple[int, int]:
        """Return (w, h) of the slide in widget pixels at current zoom."""
        aspect = self._slide_w_emu / max(1, self._slide_h_emu)
        parent = self.parent()
        avail_w = (parent.width() if parent else 800) - MARGIN * 2
        avail_h = (parent.height() if parent else 600) - MARGIN * 2
        avail_w = max(100, avail_w)
        avail_h = max(60, avail_h)

        if self._zoom != 1.0:
            # Manual zoom: base size derived from slide at 96 DPI
            base_w = int(self._slide_w_emu / 914400 * 96)
            base_h = int(self._slide_h_emu / 914400 * 96)
            return int(base_w * self._zoom), int(base_h * self._zoom)

        # Fit-to-window
        if avail_w / aspect <= avail_h:
            w = avail_w
            h = int(w / aspect)
        else:
            h = avail_h
            w = int(h * aspect)
        return w, h

    def _update_size(self):
        w, h = self._slide_display_size()
        self.setMinimumSize(w + MARGIN * 2, h + MARGIN * 2)

    def _slide_origin(self) -> QPoint:
        """Top-left corner of the slide rectangle in widget coordinates."""
        w, h = self._slide_display_size()
        ox = (self.width() - w) // 2
        oy = (self.height() - h) // 2
        return QPoint(max(MARGIN, ox), max(MARGIN, oy))

    def _slide_rect(self) -> QRect:
        w, h = self._slide_display_size()
        o = self._slide_origin()
        return QRect(o, QSize(w, h))

    # ── Shape hit-testing ──────────────────────────────────────────────

    def _build_shape_rects(self, slide_rect: QRect):
        """Compute widget-coordinate rects for all shapes."""
        self._shape_rects.clear()
        if self._slide is None:
            return
        sx = slide_rect.width() / max(1, self._slide_w_emu)
        sy = slide_rect.height() / max(1, self._slide_h_emu)
        ox, oy = slide_rect.left(), slide_rect.top()
        for shape in self._slide.shapes:
            try:
                x = ox + int((shape.left or 0) * sx)
                y = oy + int((shape.top or 0) * sy)
                w = int((shape.width or 0) * sx)
                h = int((shape.height or 0) * sy)
                self._shape_rects.append(QRect(x, y, w, h))
            except Exception:
                self._shape_rects.append(QRect())

    def _shape_at(self, pos: QPoint) -> int:
        """Return index of topmost shape under *pos*, or -1."""
        for i in range(len(self._shape_rects) - 1, -1, -1):
            if self._shape_rects[i].contains(pos):
                return i
        return -1

    # ── Paint ──────────────────────────────────────────────────────────

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)

        # Background
        painter.fillRect(self.rect(), QColor(100, 100, 100))

        if self._slide is None:
            painter.setPen(QColor(180, 180, 180))
            painter.drawText(self.rect(), Qt.AlignCenter, "열린 파일이 없습니다.")
            return

        sr = self._slide_rect()

        # Shadow
        shadow = sr.adjusted(4, 4, 4, 4)
        painter.fillRect(shadow, QColor(0, 0, 0, 80))

        # Re-render if dirty
        if self._dirty or self._pixmap is None or self._pixmap.size() != sr.size():
            self._pixmap = render_slide(
                self._slide, sr.width(), sr.height(),
                self._slide_w_emu, self._slide_h_emu
            )
            self._build_shape_rects(sr)
            self._dirty = False

        painter.drawPixmap(sr, self._pixmap)

        # Selection highlight
        if 0 <= self._selected_shape < len(self._shape_rects):
            r = self._shape_rects[self._selected_shape]
            pen = QPen(QColor(0, 120, 215), 2, Qt.SolidLine)
            painter.setPen(pen)
            painter.drawRect(r.adjusted(-1, -1, 1, 1))
            # Corner handles
            painter.setBrush(QColor(0, 120, 215))
            painter.setPen(Qt.NoPen)
            hs = 6
            for cx, cy in [(r.left(), r.top()), (r.right(), r.top()),
                           (r.left(), r.bottom()), (r.right(), r.bottom())]:
                painter.drawRect(cx - hs // 2, cy - hs // 2, hs, hs)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._dirty = True
        self._reposition_editor()
        self.update()

    # ── Mouse ──────────────────────────────────────────────────────────

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._finish_editing(save=True)
            idx = self._shape_at(event.pos())
            self._selected_shape = idx
            self.shape_selected.emit(idx)
            self.update()

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            idx = self._shape_at(event.pos())
            if idx >= 0:
                self._selected_shape = idx
                self.update()
                self._start_editing(idx)

    def mouseMoveEvent(self, event):
        idx = self._shape_at(event.pos())
        if idx >= 0:
            shapes = list(self._slide.shapes) if self._slide else []
            if idx < len(shapes) and shapes[idx].has_text_frame:
                self.setCursor(Qt.IBeamCursor)
            else:
                self.setCursor(Qt.SizeAllCursor)
        else:
            self.setCursor(Qt.ArrowCursor)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self._finish_editing(save=True)
            self._selected_shape = -1
            self.shape_selected.emit(-1)
            self.update()
        elif event.key() == Qt.Key_Delete:
            self._finish_editing(save=False)
            self._selected_shape = -1
            self.update()
        else:
            super().keyPressEvent(event)

    # ── Inline text editing ────────────────────────────────────────────

    def _start_editing(self, shape_index: int):
        if self._slide is None:
            return
        shapes = list(self._slide.shapes)
        if shape_index >= len(shapes):
            return
        shape = shapes[shape_index]
        if not shape.has_text_frame:
            return

        self._finish_editing(save=False)
        self._editing_shape = shape_index

        rect = self._shape_rects[shape_index] if shape_index < len(self._shape_rects) else QRect()
        if rect.isEmpty():
            return

        self._editor = QTextEdit(self)
        self._editor.setGeometry(rect)
        self._editor.setPlainText(shape.text_frame.text)
        self._editor.setStyleSheet(
            "background: rgba(255,255,255,230);"
            "border: 2px solid #0078d4;"
            "font-size: 13px;"
        )
        self._editor.installEventFilter(self)
        self._editor.show()
        self._editor.setFocus()
        self._editor.selectAll()

    def _finish_editing(self, save: bool = True):
        if self._editor is None:
            return
        if save and self._editing_shape >= 0 and self._model is not None:
            new_text = self._editor.toPlainText()
            self.text_changed.emit(self._editing_shape, new_text)
        self._editor.deleteLater()
        self._editor = None
        self._editing_shape = -1
        self._dirty = True
        self.update()

    def _reposition_editor(self):
        if self._editor is None or self._editing_shape < 0:
            return
        if self._editing_shape < len(self._shape_rects):
            self._editor.setGeometry(self._shape_rects[self._editing_shape])

    def eventFilter(self, obj, event):
        if obj is self._editor:
            from PyQt5.QtCore import QEvent
            if event.type() == QEvent.FocusOut:
                self._finish_editing(save=True)
                return False
            if event.type() == QEvent.KeyPress:
                if event.key() == Qt.Key_Escape:
                    self._finish_editing(save=False)
                    return True
        return super().eventFilter(obj, event)


class SlideCanvas(QScrollArea):
    """Scroll area wrapping SlideWidget with zoom controls."""

    shape_selected = pyqtSignal(int)
    text_changed = pyqtSignal(int, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setWidgetResizable(False)
        self.setStyleSheet("QScrollArea { background: #646464; border: none; }")

        self._slide_widget = SlideWidget()
        self._slide_widget.shape_selected.connect(self.shape_selected)
        self._slide_widget.text_changed.connect(self.text_changed)
        self.setWidget(self._slide_widget)

        self._fit_on_next_show = True

    # ── Public API ─────────────────────────────────────────────────────

    def set_slide(self, slide, model, slide_index: int):
        self._slide_widget.set_slide(slide, model, slide_index)
        if self._fit_on_next_show:
            self.fit_to_window()
            self._fit_on_next_show = False

    def invalidate(self):
        self._slide_widget.invalidate()

    def fit_to_window(self):
        self._slide_widget.set_zoom(1.0)
        self._slide_widget._zoom = 1.0
        self._slide_widget._dirty = True
        self._slide_widget._update_size()
        self._slide_widget.update()

    def zoom_in(self):
        z = round(self._slide_widget.zoom() + ZOOM_STEP, 2)
        self._slide_widget.set_zoom(z)

    def zoom_out(self):
        z = round(self._slide_widget.zoom() - ZOOM_STEP, 2)
        self._slide_widget.set_zoom(z)

    def zoom_value(self) -> float:
        return self._slide_widget.zoom()

    def set_zoom(self, zoom: float):
        self._slide_widget.set_zoom(zoom)

    def showEvent(self, event):
        super().showEvent(event)
        if self._fit_on_next_show:
            self.fit_to_window()
            self._fit_on_next_show = False

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # In fit mode (zoom≈1.0), re-fit on resize
        if abs(self._slide_widget.zoom() - 1.0) < 0.01:
            self._slide_widget._dirty = True
            self._slide_widget._update_size()
            self._slide_widget.update()
