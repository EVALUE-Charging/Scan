import sys
import os
import fitz  # PyMuPDF
import numpy as np
import cv2
from pyzbar.pyzbar import decode
import pandas as pd
from pathlib import Path
# openpyxl 僅作為預覽讀取備用，xlsx 輸出改為直接操作 XML
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel,
                             QVBoxLayout, QHBoxLayout, QWidget, QFileDialog,
                             QProgressBar, QMessageBox, QTextEdit, QFrame,
                             QSizePolicy, QGraphicsDropShadowEffect, QLineEdit,
                             QDateEdit, QPlainTextEdit, QTabWidget, QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate
from PyQt5.QtGui import QFont, QColor


# ── 固定參數 ─────────────────────────────────────────────────
DEFAULT_DPI         = 300
DEFAULT_REGION_ROWS = 2
DEFAULT_REGION_COLS = 2

COLUMNS = ["網址", "驗證碼", "序號", "有效起始日", "有效結束日", "商品名稱"]

# 品牌色
C_NAVY   = "#003366"
C_ORANGE = "#f28500"
C_NAVY_L = "#004a8f"
C_NAVY_D = "#002244"
C_BG     = "#f0f4f8"
C_CARD   = "#ffffff"
C_TEXT   = "#0d1b2a"
C_MUTED  = "#5a7099"
C_BORDER = "#c8d8ea"


# ── 影像處理 ─────────────────────────────────────────────────

def preprocess_variants(image_bgr):
    variants = []
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    variants.append(gray)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    variants.append(clahe.apply(gray))
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    _, otsu = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    variants.append(otsu)
    adapt_mean = cv2.adaptiveThreshold(
        blurred, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, 5)
    variants.append(adapt_mean)
    adapt_gauss = cv2.adaptiveThreshold(
        blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 15, 5)
    variants.append(adapt_gauss)
    kernel_sharpen = np.array([[-1, -1, -1], [-1, 9, -1], [-1, -1, -1]])
    sharpened = cv2.filter2D(gray, -1, kernel_sharpen)
    variants.append(sharpened)
    kernel = np.ones((2, 2), np.uint8)
    closed = cv2.morphologyEx(otsu, cv2.MORPH_CLOSE, kernel)
    variants.append(closed)
    variants.append(cv2.bitwise_not(otsu))
    h, w = gray.shape
    scaled = cv2.resize(gray, (int(w * 1.5), int(h * 1.5)), interpolation=cv2.INTER_CUBIC)
    variants.append(scaled)
    return variants


def decode_qr_from_image(image_bgr):
    found = {}
    for variant in preprocess_variants(image_bgr):
        results = decode(variant)
        for obj in results:
            try:
                text = obj.data.decode("utf-8")
            except UnicodeDecodeError:
                text = obj.data.decode("latin-1")
            if text not in found:
                found[text] = obj
    return list(found.keys())


def split_page_into_regions(image_bgr, rows=2, cols=2):
    h, w = image_bgr.shape[:2]
    rh, rw = h // rows, w // cols
    regions = []
    for r in range(rows):
        for c in range(cols):
            y1, y2 = r * rh, (r + 1) * rh if r < rows - 1 else h
            x1, x2 = c * rw, (c + 1) * rw if c < cols - 1 else w
            regions.append(image_bgr[y1:y2, x1:x2])
    return regions


def parse_manual_qr_input(text: str) -> list:
    return [line.strip() for line in text.splitlines() if line.strip()]


def save_to_excel(qr_codes: list, start_date: str, end_date: str,
                  product_name: str, output_path: str):
    """
    直接操作 XML 產生 xlsx，結構完全對齊系統參考檔：
    - 字串欄用 sharedStrings（t="s"），不用 inlineStr
    - A/B 欄（網址/驗證碼）完全不寫 cell（None）
    - D/E 欄（日期）存為整數，無 number_format 覆蓋
    """
    import zipfile, io, html

    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    # ── sharedStrings 索引 ──────────────────────────────────
    shared: list  = []
    shared_map: dict = {}

    def si(val: str) -> int:
        if val not in shared_map:
            shared_map[val] = len(shared)
            shared.append(val)
        return shared_map[val]

    # 標題先佔位
    for h in COLUMNS:
        si(h)
    # 資料
    for qr in qr_codes:
        si(str(qr))
    si(str(product_name))

    sd_int = int(start_date)
    ed_int = int(end_date)

    # ── sheet1.xml ──────────────────────────────────────────
    total_rows = 1 + len(qr_codes)
    sheet_parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{NS}">',
        f'<dimension ref="A1:F{total_rows}"/>',
        '<sheetData>',
    ]

    # 標題列
    hdr_cells = "".join(
        f'<c r="{chr(64+ci)}1" t="s"><v>{si(h)}</v></c>'
        for ci, h in enumerate(COLUMNS, 1)
    )
    sheet_parts.append(f'<row r="1">{hdr_cells}</row>')

    # 資料列（A/B 完全不寫）
    for ri, qr in enumerate(qr_codes, 2):
        sheet_parts.append(
            f'<row r="{ri}">'
            f'<c r="C{ri}" t="s"><v>{si(str(qr))}</v></c>'
            f'<c r="D{ri}"><v>{sd_int}</v></c>'
            f'<c r="E{ri}"><v>{ed_int}</v></c>'
            f'<c r="F{ri}" t="s"><v>{si(str(product_name))}</v></c>'
            f'</row>'
        )

    sheet_parts += ['</sheetData>', '</worksheet>']
    sheet_xml = "\n".join(sheet_parts)

    # ── sharedStrings.xml ───────────────────────────────────
    def xml_escape(s: str) -> str:
        return html.escape(str(s), quote=False)

    ss_items = "".join(f'<si><t>{xml_escape(v)}</t></si>' for v in shared)
    n = len(shared)
    ss_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="{NS}" count="{n}" uniqueCount="{n}">{ss_items}</sst>'
    )

    # ── 其他必要 XML ────────────────────────────────────────
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    wb_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="SheetJS" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{NS}">'
        '<numFmts count="0"/>'
        '<fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '</styleSheet>'
    )

    # ── 組裝 zip ────────────────────────────────────────────
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",   content_types)
        zf.writestr("_rels/.rels",           rels)
        zf.writestr("xl/workbook.xml",       wb_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/worksheets/sheet1.xml",   sheet_xml)
        zf.writestr("xl/sharedStrings.xml",  ss_xml)
        zf.writestr("xl/styles.xml",         styles_xml)

    with open(output_path, "wb") as f:
        f.write(buf.getvalue())


# ── 掃描執行緒 ───────────────────────────────────────────────

class QRScannerThread(QThread):
    update_progress = pyqtSignal(int)
    scan_complete   = pyqtSignal(list)
    error_occurred  = pyqtSignal(str)

    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path

    def run(self):
        try:
            doc   = fitz.open(self.pdf_path)
            total = len(doc)
            qrs   = []
            for pn in range(total):
                self.update_progress.emit(int((pn + 1) / total * 100))
                page = doc.load_page(pn)
                pix  = page.get_pixmap(dpi=DEFAULT_DPI)
                img  = cv2.imdecode(
                    np.frombuffer(pix.tobytes("png"), np.uint8), cv2.IMREAD_COLOR)
                seen = set()
                for t in decode_qr_from_image(img):
                    seen.add(t)
                for region in split_page_into_regions(img, DEFAULT_REGION_ROWS, DEFAULT_REGION_COLS):
                    for t in decode_qr_from_image(region):
                        seen.add(t)
                qrs.extend(seen)
            self.scan_complete.emit(qrs)
        except Exception as e:
            self.error_occurred.emit(str(e))


# ── 自訂元件 ─────────────────────────────────────────────────

class CardWidget(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        sh = QGraphicsDropShadowEffect(self)
        sh.setBlurRadius(20)
        sh.setOffset(0, 3)
        sh.setColor(QColor(0, 51, 102, 25))
        self.setGraphicsEffect(sh)


class PrimaryBtn(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setObjectName("primaryBtn")
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(52)


class SecondaryBtn(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setObjectName("secondaryBtn")
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(44)


class OrangeBtn(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setObjectName("orangeBtn")
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(56)


# ── 主視窗 ───────────────────────────────────────────────────

STYLESHEET = f"""
    QMainWindow, QWidget#root {{ background: {C_BG}; }}

    QFrame#card {{
        background: {C_CARD};
        border-radius: 16px;
        border: 1.5px solid {C_BORDER};
    }}

    QLabel#appTitle {{
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 30px; font-weight: 800;
        color: {C_NAVY_D}; letter-spacing: 1px;
    }}
    QLabel#appSubtitle {{
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 16px; color: {C_MUTED};
    }}
    QLabel#sectionLabel {{
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 12px; font-weight: 700;
        color: {C_ORANGE}; letter-spacing: 2.5px;
    }}
    QLabel#fieldLabel {{
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 16px; font-weight: 600; color: {C_NAVY};
    }}
    QLabel#pathLabel {{
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 15px; color: {C_MUTED};
        background: {C_BG}; border-radius: 10px;
        padding: 10px 14px;
        border: 1.5px solid {C_BORDER};
    }}
    QLabel#statusLabel {{
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 14px; color: {C_MUTED};
    }}
    QLabel#hintLabel {{
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 14px; color: {C_MUTED};
    }}

    QLineEdit, QDateEdit {{
        background: {C_BG};
        border: 2px solid {C_BORDER};
        border-radius: 10px;
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 16px; color: {C_TEXT};
        padding: 8px 14px;
        min-height: 40px;
    }}
    QLineEdit:focus, QDateEdit:focus {{
        border: 2px solid {C_NAVY};
        background: #ffffff;
    }}

    QPlainTextEdit {{
        background: {C_BG};
        border: 2px solid {C_BORDER};
        border-radius: 10px;
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 15px; color: {C_TEXT};
        padding: 10px 14px;
    }}
    QPlainTextEdit:focus {{ border: 2px solid {C_NAVY}; background: #ffffff; }}

    QPushButton#primaryBtn {{
        background: {C_NAVY};
        color: #ffffff; border: none; border-radius: 12px;
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 17px; font-weight: 700;
        letter-spacing: 1px; padding: 0 28px;
    }}
    QPushButton#primaryBtn:hover  {{ background: {C_NAVY_L}; }}
    QPushButton#primaryBtn:pressed {{ background: {C_NAVY_D}; }}
    QPushButton#primaryBtn:disabled {{ background: #b0bec9; color: #e0e8f0; }}

    QPushButton#orangeBtn {{
        background: {C_ORANGE};
        color: #ffffff; border: none; border-radius: 12px;
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 18px; font-weight: 700;
        letter-spacing: 1px; padding: 0 28px;
    }}
    QPushButton#orangeBtn:hover  {{ background: #d97500; }}
    QPushButton#orangeBtn:pressed {{ background: #bf6600; }}
    QPushButton#orangeBtn:disabled {{ background: #d4b07a; color: #fff8f0; }}

    QPushButton#secondaryBtn {{
        background: #e6edf5; color: {C_NAVY};
        border: 1.5px solid {C_BORDER};
        border-radius: 10px;
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 15px; font-weight: 600; padding: 0 18px;
    }}
    QPushButton#secondaryBtn:hover  {{ background: #d0dcea; }}
    QPushButton#secondaryBtn:pressed {{ background: #bccbdf; }}

    QProgressBar {{
        background: #dce6f0; border: none;
        border-radius: 8px; height: 14px; color: transparent;
    }}
    QProgressBar::chunk {{
        background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
            stop:0 {C_NAVY}, stop:1 {C_ORANGE});
        border-radius: 8px;
    }}

    QTextEdit {{
        background: #f7fafd;
        border: 2px solid {C_BORDER};
        border-radius: 12px;
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 15px; color: {C_TEXT};
        padding: 10px;
    }}

    QTabWidget::pane {{ border: none; background: transparent; }}
    QTabBar::tab {{
        font-family: 'Microsoft JhengHei UI', 'PingFang TC', sans-serif;
        font-size: 15px; font-weight: 600;
        background: #dce6f0; color: {C_MUTED};
        border-radius: 10px 10px 0 0;
        padding: 10px 30px; margin-right: 4px;
    }}
    QTabBar::tab:selected {{ background: {C_NAVY}; color: #ffffff; }}
    QTabBar::tab:hover:!selected {{ background: #c8d8ea; color: {C_NAVY}; }}

    QScrollArea {{ border: none; background: transparent; }}
    QScrollBar:vertical {{
        background: {C_BG}; width: 8px; border-radius: 4px; margin: 0;
    }}
    QScrollBar::handle:vertical {{
        background: #b0c4d8; border-radius: 4px; min-height: 30px;
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
"""


class QRScannerApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.qr_codes    = []
        self.pdf_path    = ""
        self.output_path = ""
        self.initUI()

    def initUI(self):
        self.setWindowTitle("PDF QR碼掃描器")
        self.setStyleSheet(STYLESHEET)
        self.showMaximized()

        # 捲動容器
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        container.setObjectName("root")
        scroll.setWidget(container)
        self.setCentralWidget(scroll)

        outer = QVBoxLayout(container)
        outer.setContentsMargins(60, 40, 60, 40)
        outer.setSpacing(28)

        # ── 標題 ──────────────────────────────────────────────
        hdr = QHBoxLayout()
        accent = QFrame()
        accent.setFixedWidth(7)
        accent.setStyleSheet(f"background:{C_ORANGE}; border-radius:3px;")
        hdr.addWidget(accent)
        hdr.addSpacing(18)

        tc = QVBoxLayout()
        tc.setSpacing(5)
        t = QLabel("PDF QR碼掃描器")
        t.setObjectName("appTitle")
        s = QLabel("自動辨識 PDF 中的 QR 碼，或手動輸入後匯出 Excel")
        s.setObjectName("appSubtitle")
        tc.addWidget(t)
        tc.addWidget(s)
        hdr.addLayout(tc)
        hdr.addStretch()
        outer.addLayout(hdr)

        # ── 參數卡片 ──────────────────────────────────────────
        param_card = CardWidget()
        pl = QVBoxLayout(param_card)
        pl.setContentsMargins(32, 24, 32, 24)
        pl.setSpacing(20)

        pl.addWidget(self._section("EXPORT PARAMETERS"))

        # 商品名稱
        r1 = QHBoxLayout()
        r1.setSpacing(16)
        l1 = QLabel("商品名稱")
        l1.setObjectName("fieldLabel")
        l1.setFixedWidth(110)
        self.product_input = QLineEdit()
        self.product_input.setPlaceholderText("請輸入商品名稱…")
        r1.addWidget(l1)
        r1.addWidget(self.product_input)
        pl.addLayout(r1)

        # 日期（yyyyMMdd，無斜線）
        r2 = QHBoxLayout()
        r2.setSpacing(16)
        l2 = QLabel("有效起始日")
        l2.setObjectName("fieldLabel")
        l2.setFixedWidth(110)
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setDisplayFormat("yyyyMMdd")

        l3 = QLabel("有效結束日")
        l3.setObjectName("fieldLabel")
        l3.setFixedWidth(110)
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate().addYears(1))
        self.end_date.setDisplayFormat("yyyyMMdd")

        r2.addWidget(l2)
        r2.addWidget(self.start_date)
        r2.addSpacing(24)
        r2.addWidget(l3)
        r2.addWidget(self.end_date)
        pl.addLayout(r2)

        outer.addWidget(param_card)

        # ── QR 來源卡片 ───────────────────────────────────────
        src_card = CardWidget()
        sl = QVBoxLayout(src_card)
        sl.setContentsMargins(32, 24, 32, 24)
        sl.setSpacing(16)
        sl.addWidget(self._section("QR CODE 來源"))

        self.tab = QTabWidget()
        sl.addWidget(self.tab)

        # Tab 1：手動輸入
        mt = QWidget()
        mtl = QVBoxLayout(mt)
        mtl.setContentsMargins(0, 12, 0, 0)
        mtl.setSpacing(10)
        hint = QLabel("每行貼上一筆 QR 碼（可一次貼入多筆）")
        hint.setObjectName("hintLabel")
        self.manual_input = QPlainTextEdit()
        self.manual_input.setPlaceholderText("ABC123\nDEF456\nGHI789")
        self.manual_input.setMinimumHeight(140)
        mtl.addWidget(hint)
        mtl.addWidget(self.manual_input)
        self.tab.addTab(mt, "✏️  手動輸入")

        # Tab 2：PDF 掃描
        pt = QWidget()
        ptl = QVBoxLayout(pt)
        ptl.setContentsMargins(0, 12, 0, 0)
        ptl.setSpacing(16)

        pdf_row = QHBoxLayout()
        pdf_row.setSpacing(12)
        self.pdf_path_label = QLabel("尚未選擇 PDF 檔案")
        self.pdf_path_label.setObjectName("pathLabel")
        self.pdf_path_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.pdf_pick_btn = SecondaryBtn("📂  選擇 PDF")
        self.pdf_pick_btn.setFixedWidth(160)
        self.pdf_pick_btn.clicked.connect(self.select_pdf_file)
        pdf_row.addWidget(self.pdf_path_label)
        pdf_row.addWidget(self.pdf_pick_btn)
        ptl.addLayout(pdf_row)

        prog_hdr = QHBoxLayout()
        ps2 = self._section("SCAN PROGRESS")
        self.status_label = QLabel("等待開始…")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignRight)
        prog_hdr.addWidget(ps2)
        prog_hdr.addStretch()
        prog_hdr.addWidget(self.status_label)
        ptl.addLayout(prog_hdr)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(14)
        ptl.addWidget(self.progress_bar)

        self.scan_btn = PrimaryBtn("▶  開始掃描 PDF")
        self.scan_btn.clicked.connect(self.start_scanning)
        ptl.addWidget(self.scan_btn)

        self.tab.addTab(pt, "📄  PDF 掃描")
        outer.addWidget(src_card)

        # ── 結果卡片 ──────────────────────────────────────────
        res_card = CardWidget()
        rl = QVBoxLayout(res_card)
        rl.setContentsMargins(32, 24, 32, 24)
        rl.setSpacing(12)
        rl.addWidget(self._section("SCAN RESULTS"))
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(160)
        self.result_text.setPlaceholderText("QR 碼預覽將顯示於此…")
        rl.addWidget(self.result_text)
        outer.addWidget(res_card)

        # ── 輸出卡片 ──────────────────────────────────────────
        out_card = CardWidget()
        ol = QVBoxLayout(out_card)
        ol.setContentsMargins(32, 24, 32, 24)
        ol.setSpacing(16)
        ol.addWidget(self._section("OUTPUT FILE"))

        out_row = QHBoxLayout()
        out_row.setSpacing(12)
        self.out_path_label = QLabel("尚未選擇輸出位置")
        self.out_path_label.setObjectName("pathLabel")
        self.out_path_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.out_pick_btn = OrangeBtn("📥  選擇位置並匯出")
        self.out_pick_btn.setFixedWidth(220)
        self.out_pick_btn.clicked.connect(self.select_output_file)
        out_row.addWidget(self.out_path_label)
        out_row.addWidget(self.out_pick_btn)
        ol.addLayout(out_row)

        outer.addWidget(out_card)
        outer.addStretch()

    # ── 工具 ─────────────────────────────────────────────────

    @staticmethod
    def _section(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("sectionLabel")
        return lbl

    # ── 槽函數 ───────────────────────────────────────────────

    def select_pdf_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "選擇 PDF 檔案", "", "PDF 檔案 (*.pdf)")
        if path:
            self.pdf_path = path
            self.pdf_path_label.setText(f"  {os.path.basename(path)}")
            self.status_label.setText("等待掃描…")
            self.progress_bar.setValue(0)

    def select_output_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "選擇輸出位置", "", "Excel 檔案 (*.xlsx)")
        if path:
            if not path.endswith(".xlsx"):
                path += ".xlsx"
            self.output_path = path
            self.out_path_label.setText(f"  {os.path.basename(path)}")
            self.export_excel()

    def start_scanning(self):
        if not self.pdf_path:
            QMessageBox.warning(self, "警告", "請先選擇 PDF 檔案")
            return
        self.result_text.clear()
        self.progress_bar.setValue(0)
        self.scan_btn.setEnabled(False)
        self.status_label.setText("掃描中…")
        self.scanner_thread = QRScannerThread(self.pdf_path)
        self.scanner_thread.update_progress.connect(self.progress_bar.setValue)
        self.scanner_thread.update_progress.connect(
            lambda v: self.status_label.setText(f"{v}%"))
        self.scanner_thread.scan_complete.connect(self.on_scan_complete)
        self.scanner_thread.error_occurred.connect(self.show_error)
        self.scanner_thread.start()

    def on_scan_complete(self, qr_list: list):
        self.qr_codes = qr_list
        self.scan_btn.setEnabled(True)
        self.status_label.setText(f"完成 ✓  共 {len(qr_list)} 筆")
        self._refresh_preview(qr_list)

    def _refresh_preview(self, codes: list):
        if not codes:
            self.result_text.setText("（未找到任何 QR 碼）")
        else:
            lines = [f"共 {len(codes)} 筆 QR 碼\n" + "─" * 50]
            for i, c in enumerate(codes, 1):
                lines.append(f"  {i:>4}.  {c}")
            self.result_text.setText("\n".join(lines))

    def export_excel(self):
        if self.tab.currentIndex() == 0:
            codes = parse_manual_qr_input(self.manual_input.toPlainText())
        else:
            codes = self.qr_codes

        if not codes:
            QMessageBox.warning(self, "警告", "沒有 QR 碼可匯出，請先輸入或掃描")
            return

        product = self.product_input.text().strip()
        start   = self.start_date.date().toString("yyyyMMdd")
        end     = self.end_date.date().toString("yyyyMMdd")

        if self.tab.currentIndex() == 0:
            self._refresh_preview(codes)

        try:
            save_to_excel(codes, start, end, product, self.output_path)
            QMessageBox.information(
                self, "匯出完成",
                f"共匯出 {len(codes)} 筆 QR 碼\n\n"
                f"商品名稱：{product or '（未填）'}\n"
                f"有效期間：{start} ～ {end}\n\n"
                f"已儲存至：\n{self.output_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"儲存時發生錯誤：\n{str(e)}")

    def show_error(self, msg):
        self.scan_btn.setEnabled(True)
        self.status_label.setText("發生錯誤")
        QMessageBox.critical(self, "錯誤", f"掃描過程中發生錯誤：\n{msg}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = QRScannerApp()
    sys.exit(app.exec_())