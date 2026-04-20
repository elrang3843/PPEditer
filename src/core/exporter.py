"""
Export functionality — PDF via Qt's QPrinter.
"""
from PyQt5.QtCore import QSizeF, Qt
from PyQt5.QtGui import QPainter
from PyQt5.QtPrintSupport import QPrinter

from .slide_renderer import render_slide


def export_to_pdf(model, filepath: str,
                  progress_callback=None) -> bool:
    """
    Render every slide and write them to a PDF file.

    *progress_callback(current, total)* is called before each page if provided.
    Returns True on success, False otherwise.
    """
    if not model.is_open:
        return False

    # Slide dimensions in mm (1 inch = 25.4 mm)
    emu_per_inch = 914400.0
    slide_w_mm = model.slide_width / emu_per_inch * 25.4
    slide_h_mm = model.slide_height / emu_per_inch * 25.4

    printer = QPrinter(QPrinter.HighResolution)
    printer.setOutputFormat(QPrinter.PdfFormat)
    printer.setOutputFileName(filepath)
    printer.setPageSize(QPrinter.Custom)
    printer.setPageSizeMM(QSizeF(slide_w_mm, slide_h_mm))
    printer.setPageMargins(0, 0, 0, 0, QPrinter.Millimeter)
    printer.setColorMode(QPrinter.Color)

    painter = QPainter()
    if not painter.begin(printer):
        return False

    try:
        page_rect = printer.pageRect(QPrinter.DevicePixel).toRect()
        pw = max(1, page_rect.width())
        ph = max(1, page_rect.height())

        for i in range(model.slide_count):
            if progress_callback:
                progress_callback(i, model.slide_count)
            if i > 0:
                printer.newPage()
            slide = model.get_slide(i)
            pixmap = render_slide(slide, pw, ph,
                                  model.slide_width, model.slide_height)
            painter.drawPixmap(page_rect, pixmap)
    finally:
        painter.end()

    return True
