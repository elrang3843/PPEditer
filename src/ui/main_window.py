"""
Main application window — menus, toolbar, layout, and all top-level actions.
"""
import os

from PyQt5.QtCore import Qt, QSettings, QSize, QPoint, QTimer
from PyQt5.QtGui import QKeySequence, QIcon, QFont, QColor, QPalette
from PyQt5.QtWidgets import (
    QMainWindow,
    QSplitter,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QStatusBar,
    QToolBar,
    QAction,
    QActionGroup,
    QMenuBar,
    QFileDialog,
    QMessageBox,
    QProgressDialog,
    QApplication,
    QSizePolicy,
    QSpinBox,
    QComboBox,
    QFrame,
)

from src.core.presentation_model import PresentationModel
from src.ui.slide_panel import SlidePanel
from src.ui.canvas import SlideCanvas

APP_NAME = "PPEditer"
RECENT_MAX = 10
ORG = "HANDTECH"
DOMAIN = "handtech.co.kr"


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self._model = PresentationModel()
        self._current_slide_index = 0
        self._settings = QSettings(ORG, APP_NAME)

        self._setup_ui()
        self._create_menu_bar()
        self._create_toolbar()
        self._restore_geometry()
        self._update_title()
        self._update_actions()

    # ── UI setup ───────────────────────────────────────────────────────

    def _setup_ui(self):
        self.setMinimumSize(960, 620)

        # Central splitter
        self._splitter = QSplitter(Qt.Horizontal)
        self._splitter.setHandleWidth(4)
        self.setCentralWidget(self._splitter)

        # Left: slide panel
        self._slide_panel = SlidePanel()
        self._slide_panel.slide_selected.connect(self._on_slide_selected)
        self._slide_panel.slide_reorder_requested.connect(self._on_slide_reorder)
        self._slide_panel.add_slide_requested.connect(self._on_add_slide_after)
        self._slide_panel.delete_slide_requested.connect(self._on_delete_slide)
        self._slide_panel.duplicate_slide_requested.connect(self._on_duplicate_slide)
        self._splitter.addWidget(self._slide_panel)

        # Right: canvas
        self._canvas = SlideCanvas()
        self._canvas.text_changed.connect(self._on_text_changed)
        self._canvas.shape_selected.connect(self._on_shape_selected)
        self._splitter.addWidget(self._canvas)

        self._splitter.setSizes([200, 760])
        self._splitter.setStretchFactor(0, 0)
        self._splitter.setStretchFactor(1, 1)

        # Status bar
        self._status_bar = self.statusBar()
        self._lbl_slide_info = QLabel("슬라이드 정보 없음")
        self._lbl_zoom = QLabel("100%")
        self._lbl_zoom.setFixedWidth(60)
        self._status_bar.addPermanentWidget(self._lbl_slide_info)
        self._status_bar.addPermanentWidget(self._lbl_zoom)

    # ── Menu bar ───────────────────────────────────────────────────────

    def _create_menu_bar(self):
        mb = self.menuBar()

        # ── File ──────────────────────────────────────────────────────
        file_menu = mb.addMenu("파일(&F)")

        self._act_new = self._action(
            "새로 만들기(&N)", "Ctrl+N", self._cmd_new, tip="새 프레젠테이션 만들기"
        )
        file_menu.addAction(self._act_new)

        self._act_open = self._action(
            "열기(&O)...", "Ctrl+O", self._cmd_open, tip="파일 열기"
        )
        file_menu.addAction(self._act_open)

        self._recent_menu = file_menu.addMenu("최근 파일(&R)")
        self._rebuild_recent_menu()

        file_menu.addSeparator()

        self._act_save = self._action(
            "저장(&S)", "Ctrl+S", self._cmd_save, tip="저장"
        )
        file_menu.addAction(self._act_save)

        self._act_save_as = self._action(
            "다른 이름으로 저장(&A)...", "Ctrl+Shift+S", self._cmd_save_as, tip="다른 이름으로 저장"
        )
        file_menu.addAction(self._act_save_as)

        file_menu.addSeparator()

        self._act_export_pdf = self._action(
            "PDF로 내보내기(&E)...", "Ctrl+Shift+E", self._cmd_export_pdf,
            tip="PDF 파일로 내보내기"
        )
        file_menu.addAction(self._act_export_pdf)

        file_menu.addSeparator()

        act_exit = self._action("끝내기(&X)", "Alt+F4", self.close, tip="프로그램 종료")
        file_menu.addAction(act_exit)

        # ── Edit ──────────────────────────────────────────────────────
        edit_menu = mb.addMenu("편집(&E)")

        self._act_undo = self._action("실행 취소(&Z)", "Ctrl+Z", self._cmd_undo)
        edit_menu.addAction(self._act_undo)

        self._act_redo = self._action("다시 실행(&Y)", "Ctrl+Y", self._cmd_redo)
        edit_menu.addAction(self._act_redo)

        edit_menu.addSeparator()

        # Standard clipboard actions delegate to focusWidget
        self._act_cut = self._action("잘라내기(&T)", "Ctrl+X",
                                      lambda: self._clipboard_action("cut"))
        edit_menu.addAction(self._act_cut)

        self._act_copy = self._action("복사(&C)", "Ctrl+C",
                                       lambda: self._clipboard_action("copy"))
        edit_menu.addAction(self._act_copy)

        self._act_paste = self._action("붙여넣기(&P)", "Ctrl+V",
                                        lambda: self._clipboard_action("paste"))
        edit_menu.addAction(self._act_paste)

        self._act_select_all = self._action("모두 선택(&A)", "Ctrl+A",
                                             lambda: self._clipboard_action("selectAll"))
        edit_menu.addAction(self._act_select_all)

        edit_menu.addSeparator()

        self._act_delete_shape = self._action(
            "선택 항목 삭제(&D)", "Delete", self._cmd_delete_selected_shape
        )
        edit_menu.addAction(self._act_delete_shape)

        # ── View ──────────────────────────────────────────────────────
        view_menu = mb.addMenu("보기(&V)")

        self._act_zoom_in = self._action("확대(&I)", "Ctrl+=", self._cmd_zoom_in)
        view_menu.addAction(self._act_zoom_in)

        self._act_zoom_out = self._action("축소(&O)", "Ctrl+-", self._cmd_zoom_out)
        view_menu.addAction(self._act_zoom_out)

        self._act_zoom_fit = self._action("창 맞춤(&F)", "Ctrl+0", self._cmd_zoom_fit)
        view_menu.addAction(self._act_zoom_fit)

        view_menu.addSeparator()

        self._act_panel = QAction("슬라이드 패널(&P)", self, checkable=True, checked=True)
        self._act_panel.triggered.connect(self._toggle_slide_panel)
        view_menu.addAction(self._act_panel)

        self._act_status_bar = QAction("상태 표시줄(&S)", self, checkable=True, checked=True)
        self._act_status_bar.triggered.connect(self._toggle_status_bar)
        view_menu.addAction(self._act_status_bar)

        # ── Slide ─────────────────────────────────────────────────────
        slide_menu = mb.addMenu("슬라이드(&L)")

        self._act_add_slide = self._action(
            "새 슬라이드(&N)", "Ctrl+M", self._cmd_add_slide
        )
        slide_menu.addAction(self._act_add_slide)

        self._act_dup_slide = self._action(
            "슬라이드 복제(&D)", "Ctrl+D", self._cmd_duplicate_slide
        )
        slide_menu.addAction(self._act_dup_slide)

        self._act_del_slide = self._action(
            "슬라이드 삭제(&X)", "Ctrl+Shift+Delete", self._cmd_delete_slide
        )
        slide_menu.addAction(self._act_del_slide)

        slide_menu.addSeparator()

        self._act_slide_up = self._action(
            "슬라이드 위로 이동(&U)", "Ctrl+Shift+Up", self._cmd_slide_up
        )
        slide_menu.addAction(self._act_slide_up)

        self._act_slide_down = self._action(
            "슬라이드 아래로 이동(&W)", "Ctrl+Shift+Down", self._cmd_slide_down
        )
        slide_menu.addAction(self._act_slide_down)

        # ── Help ──────────────────────────────────────────────────────
        help_menu = mb.addMenu("도움말(&H)")

        act_about = self._action("PPEditer 정보(&A)...", None, self._cmd_about)
        help_menu.addAction(act_about)

    # ── Toolbar ───────────────────────────────────────────────────────

    def _create_toolbar(self):
        tb = QToolBar("기본 도구 모음")
        tb.setObjectName("MainToolBar")
        tb.setMovable(False)
        tb.setIconSize(QSize(20, 20))
        tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.addToolBar(tb)

        tb.addAction(self._act_new)
        tb.addAction(self._act_open)
        tb.addAction(self._act_save)
        tb.addSeparator()
        tb.addAction(self._act_undo)
        tb.addAction(self._act_redo)
        tb.addSeparator()
        tb.addAction(self._act_add_slide)
        tb.addAction(self._act_del_slide)
        tb.addSeparator()
        tb.addAction(self._act_export_pdf)
        tb.addSeparator()

        # Zoom controls
        lbl = QLabel("  확대/축소: ")
        tb.addWidget(lbl)

        self._zoom_combo = QComboBox()
        self._zoom_combo.addItems(["50%", "75%", "100%", "125%", "150%", "200%"])
        self._zoom_combo.setCurrentText("100%")
        self._zoom_combo.setFixedWidth(72)
        self._zoom_combo.setEditable(True)
        self._zoom_combo.currentTextChanged.connect(self._on_zoom_combo_changed)
        tb.addWidget(self._zoom_combo)

        btn_fit = tb.addAction("창 맞춤")
        btn_fit.triggered.connect(self._cmd_zoom_fit)

    # ── Geometry persistence ───────────────────────────────────────────

    def _restore_geometry(self):
        geom = self._settings.value("geometry")
        state = self._settings.value("windowState")
        if geom:
            self.restoreGeometry(geom)
        else:
            self.resize(1280, 760)
            self._center_on_screen()
        if state:
            self.restoreState(state)

    def _center_on_screen(self):
        screen = QApplication.primaryScreen().geometry()
        self.move(
            (screen.width() - self.width()) // 2,
            (screen.height() - self.height()) // 2,
        )

    def closeEvent(self, event):
        if not self._confirm_discard():
            event.ignore()
            return
        self._settings.setValue("geometry", self.saveGeometry())
        self._settings.setValue("windowState", self.saveState())
        event.accept()

    # ── Action factory ─────────────────────────────────────────────────

    def _action(self, text, shortcut=None, slot=None, tip=None) -> QAction:
        act = QAction(text, self)
        if shortcut:
            act.setShortcut(QKeySequence(shortcut))
        if slot:
            act.triggered.connect(slot)
        if tip:
            act.setStatusTip(tip)
        return act

    # ── Commands ───────────────────────────────────────────────────────

    def _cmd_new(self):
        if not self._confirm_discard():
            return
        self._model.new()
        self._current_slide_index = 0
        self._refresh_all()

    def _cmd_open(self):
        if not self._confirm_discard():
            return
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "파일 열기",
            self._settings.value("lastDir", ""),
            "프레젠테이션 파일 (*.pptx *.ppt);;모든 파일 (*)",
        )
        if filepath:
            self.open_file(filepath)

    def open_file(self, filepath: str):
        try:
            self._model.open(filepath)
            self._current_slide_index = 0
            self._settings.setValue("lastDir", os.path.dirname(filepath))
            self._add_recent(filepath)
            self._refresh_all()
        except Exception as exc:
            QMessageBox.critical(self, "파일 열기 오류",
                                 f"파일을 열 수 없습니다:\n{exc}")

    def _cmd_save(self):
        if not self._model.is_open:
            return
        if not self._model.filepath:
            self._cmd_save_as()
            return
        try:
            self._model.save()
            self._update_title()
            self._status_bar.showMessage("저장되었습니다.", 3000)
        except Exception as exc:
            QMessageBox.critical(self, "저장 오류", f"저장 실패:\n{exc}")

    def _cmd_save_as(self):
        if not self._model.is_open:
            return
        filepath, _ = QFileDialog.getSaveFileName(
            self,
            "다른 이름으로 저장",
            self._settings.value("lastDir", ""),
            "PowerPoint 파일 (*.pptx);;모든 파일 (*)",
        )
        if filepath:
            if not filepath.lower().endswith((".pptx", ".ppt")):
                filepath += ".pptx"
            try:
                self._model.save(filepath)
                self._settings.setValue("lastDir", os.path.dirname(filepath))
                self._add_recent(filepath)
                self._update_title()
                self._status_bar.showMessage("저장되었습니다.", 3000)
            except Exception as exc:
                QMessageBox.critical(self, "저장 오류", f"저장 실패:\n{exc}")

    def _cmd_export_pdf(self):
        if not self._model.is_open:
            return
        default_name = os.path.splitext(self._model.filename)[0] + ".pdf"
        filepath, _ = QFileDialog.getSaveFileName(
            self,
            "PDF로 내보내기",
            os.path.join(self._settings.value("lastDir", ""), default_name),
            "PDF 파일 (*.pdf);;모든 파일 (*)",
        )
        if not filepath:
            return
        if not filepath.lower().endswith(".pdf"):
            filepath += ".pdf"

        progress = QProgressDialog("PDF 내보내기 중...", "취소", 0, self._model.slide_count, self)
        progress.setWindowTitle("내보내기")
        progress.setWindowModality(Qt.WindowModal)
        progress.show()

        def on_progress(current, total):
            progress.setValue(current)
            QApplication.processEvents()

        try:
            from src.core.exporter import export_to_pdf
            ok = export_to_pdf(self._model, filepath, on_progress)
            progress.setValue(self._model.slide_count)
            if ok:
                self._status_bar.showMessage(f"PDF 저장 완료: {filepath}", 5000)
            else:
                QMessageBox.warning(self, "내보내기 오류", "PDF 내보내기에 실패했습니다.")
        except Exception as exc:
            QMessageBox.critical(self, "내보내기 오류", f"PDF 내보내기 실패:\n{exc}")
        finally:
            progress.close()

    def _cmd_undo(self):
        if self._model.undo():
            # Try to keep current slide index valid
            self._current_slide_index = min(
                self._current_slide_index, self._model.slide_count - 1
            )
            self._refresh_all()

    def _cmd_redo(self):
        if self._model.redo():
            self._current_slide_index = min(
                self._current_slide_index, self._model.slide_count - 1
            )
            self._refresh_all()

    def _cmd_add_slide(self):
        if not self._model.is_open:
            return
        new_idx = self._model.add_slide(after_index=self._current_slide_index)
        self._current_slide_index = new_idx
        self._refresh_all()

    def _cmd_delete_slide(self):
        if not self._model.is_open:
            return
        if self._model.slide_count <= 1:
            QMessageBox.information(self, "슬라이드 삭제",
                                    "마지막 슬라이드는 삭제할 수 없습니다.")
            return
        reply = QMessageBox.question(
            self, "슬라이드 삭제",
            f"슬라이드 {self._current_slide_index + 1}을(를) 삭제하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self._model.delete_slide(self._current_slide_index)
            self._current_slide_index = min(
                self._current_slide_index, self._model.slide_count - 1
            )
            self._refresh_all()

    def _cmd_duplicate_slide(self):
        if not self._model.is_open:
            return
        new_idx = self._model.duplicate_slide(self._current_slide_index)
        self._current_slide_index = new_idx
        self._refresh_all()

    def _cmd_slide_up(self):
        idx = self._current_slide_index
        if idx > 0:
            self._model.move_slide(idx, idx - 1)
            self._current_slide_index = idx - 1
            self._refresh_all()

    def _cmd_slide_down(self):
        idx = self._current_slide_index
        if idx < self._model.slide_count - 1:
            self._model.move_slide(idx, idx + 1)
            self._current_slide_index = idx + 1
            self._refresh_all()

    def _cmd_zoom_in(self):
        self._canvas.zoom_in()
        self._update_zoom_display()

    def _cmd_zoom_out(self):
        self._canvas.zoom_out()
        self._update_zoom_display()

    def _cmd_zoom_fit(self):
        self._canvas.fit_to_window()
        self._update_zoom_display()

    def _cmd_delete_selected_shape(self):
        pass  # Shape deletion requires deeper PPTX XML manipulation — future feature

    def _cmd_about(self):
        from src.ui.dialogs import AboutDialog
        dlg = AboutDialog(self)
        dlg.exec_()

    # ── Slot handlers ──────────────────────────────────────────────────

    def _on_slide_selected(self, index: int):
        if index == self._current_slide_index:
            return
        self._current_slide_index = index
        self._show_current_slide()
        self._update_actions()
        self._update_slide_info()

    def _on_slide_reorder(self, from_idx: int, to_idx: int):
        self._model.move_slide(from_idx, to_idx)
        self._current_slide_index = to_idx
        self._refresh_all()

    def _on_add_slide_after(self, after_idx: int):
        new_idx = self._model.add_slide(after_index=after_idx)
        self._current_slide_index = new_idx
        self._refresh_all()

    def _on_delete_slide(self, index: int):
        if self._model.slide_count <= 1:
            QMessageBox.information(self, "슬라이드 삭제",
                                    "마지막 슬라이드는 삭제할 수 없습니다.")
            return
        self._model.delete_slide(index)
        self._current_slide_index = min(index, self._model.slide_count - 1)
        self._refresh_all()

    def _on_duplicate_slide(self, index: int):
        new_idx = self._model.duplicate_slide(index)
        self._current_slide_index = new_idx
        self._refresh_all()

    def _on_text_changed(self, shape_index: int, new_text: str):
        self._model.update_shape_text(
            self._current_slide_index, shape_index, new_text
        )
        self._slide_panel.update_thumbnail(self._current_slide_index)
        self._update_title()
        self._update_actions()

    def _on_shape_selected(self, shape_index: int):
        self._update_actions()

    def _on_zoom_combo_changed(self, text: str):
        try:
            value = float(text.strip().rstrip("%")) / 100.0
            self._canvas.set_zoom(value)
            self._update_zoom_display()
        except ValueError:
            pass

    # ── View toggles ───────────────────────────────────────────────────

    def _toggle_slide_panel(self, checked: bool):
        self._slide_panel.setVisible(checked)

    def _toggle_status_bar(self, checked: bool):
        self._status_bar.setVisible(checked)

    # ── Recent files ───────────────────────────────────────────────────

    def _add_recent(self, filepath: str):
        recents: list = self._settings.value("recentFiles", []) or []
        if filepath in recents:
            recents.remove(filepath)
        recents.insert(0, filepath)
        recents = recents[:RECENT_MAX]
        self._settings.setValue("recentFiles", recents)
        self._rebuild_recent_menu()

    def _rebuild_recent_menu(self):
        self._recent_menu.clear()
        recents: list = self._settings.value("recentFiles", []) or []
        for i, path in enumerate(recents):
            label = f"&{i + 1}. {os.path.basename(path)}"
            act = QAction(label, self)
            act.setStatusTip(path)
            act.triggered.connect(lambda checked, p=path: self._open_recent(p))
            self._recent_menu.addAction(act)
        if not recents:
            empty = QAction("(없음)", self)
            empty.setEnabled(False)
            self._recent_menu.addAction(empty)

    def _open_recent(self, filepath: str):
        if not os.path.exists(filepath):
            QMessageBox.warning(self, "파일 없음",
                                f"파일을 찾을 수 없습니다:\n{filepath}")
            return
        if not self._confirm_discard():
            return
        self.open_file(filepath)

    # ── Refresh helpers ────────────────────────────────────────────────

    def _refresh_all(self):
        """Rebuild thumbnail panel and redraw canvas."""
        self._slide_panel.refresh(self._model, self._current_slide_index)
        self._show_current_slide()
        self._update_title()
        self._update_actions()
        self._update_slide_info()

    def _show_current_slide(self):
        slide = self._model.get_slide(self._current_slide_index)
        self._canvas.set_slide(slide, self._model, self._current_slide_index)
        self._update_slide_info()

    def _update_title(self):
        if self._model.is_open:
            self.setWindowTitle(self._model.window_title)
        else:
            self.setWindowTitle(APP_NAME)

    def _update_slide_info(self):
        if self._model.is_open:
            self._lbl_slide_info.setText(
                f"슬라이드 {self._current_slide_index + 1} / {self._model.slide_count}"
            )
        else:
            self._lbl_slide_info.setText("")

    def _update_zoom_display(self):
        pct = int(self._canvas.zoom_value() * 100)
        self._lbl_zoom.setText(f"{pct}%")
        self._zoom_combo.blockSignals(True)
        self._zoom_combo.setCurrentText(f"{pct}%")
        self._zoom_combo.blockSignals(False)

    def _update_actions(self):
        has = self._model.is_open
        self._act_save.setEnabled(has and self._model.modified)
        self._act_save_as.setEnabled(has)
        self._act_export_pdf.setEnabled(has)
        self._act_add_slide.setEnabled(has)
        self._act_dup_slide.setEnabled(has)
        self._act_del_slide.setEnabled(has and self._model.slide_count > 1)
        self._act_slide_up.setEnabled(has and self._current_slide_index > 0)
        self._act_slide_down.setEnabled(
            has and self._current_slide_index < self._model.slide_count - 1
        )
        self._act_undo.setEnabled(self._model.can_undo)
        self._act_redo.setEnabled(self._model.can_redo)

    # ── Utility ────────────────────────────────────────────────────────

    def _confirm_discard(self) -> bool:
        """Return True if it's OK to discard unsaved changes."""
        if not self._model.is_open or not self._model.modified:
            return True
        reply = QMessageBox.question(
            self,
            "저장되지 않은 변경사항",
            f"'{self._model.filename}'의 변경사항을 저장하시겠습니까?",
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
        )
        if reply == QMessageBox.Save:
            self._cmd_save()
            return not self._model.modified  # False if save was cancelled/failed
        return reply == QMessageBox.Discard

    def _clipboard_action(self, action: str):
        fw = QApplication.focusWidget()
        if fw and hasattr(fw, action):
            getattr(fw, action)()

    # ── Keyboard shortcuts not covered by menu actions ─────────────────

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F5:
            self._show_current_slide()
        elif event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_Equal:
            self._cmd_zoom_in()
        elif event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_Minus:
            self._cmd_zoom_out()
        else:
            super().keyPressEvent(event)
