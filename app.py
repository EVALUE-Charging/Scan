import sys
import os
import io
import html
import zipfile
import tempfile

import fitz                     # PyMuPDF
import numpy as np
import cv2
from pyzbar.pyzbar import decode as pyzbar_decode
import streamlit as st
from pathlib import Path

# ── 固定參數 ─────────────────────────────────────────────────
DEFAULT_DPI         = 300
DEFAULT_REGION_ROWS = 2
DEFAULT_REGION_COLS = 2

COLUMNS = ["網址", "驗證碼", "序號", "有效起始日", "有效結束日", "商品名稱"]

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
    scaled = cv2.resize(gray, (int(w * 1.5), int(h * 1.5)),
                        interpolation=cv2.INTER_CUBIC)
    variants.append(scaled)

    return variants


def decode_qr_from_image(image_bgr):
    """用 pyzbar 解碼，對每種前處理變體嘗試"""
    found = {}
    for variant in preprocess_variants(image_bgr):
        results = pyzbar_decode(variant)
        for obj in results:
            try:
                text = obj.data.decode("utf-8")
            except UnicodeDecodeError:
                text = obj.data.decode("latin-1")
            if text and text not in found:
                found[text] = obj
    return list(found.keys())


def split_page_into_regions(image_bgr, rows=2, cols=2):
    h, w = image_bgr.shape[:2]
    rh, rw = h // rows, w // cols
    regions = []
    for r in range(rows):
        for c in range(cols):
            y1 = r * rh
            y2 = (r + 1) * rh if r < rows - 1 else h
            x1 = c * rw
            x2 = (c + 1) * rw if c < cols - 1 else w
            regions.append(image_bgr[y1:y2, x1:x2])
    return regions


def scan_pdf(pdf_bytes: bytes) -> list[str]:
    """掃描 PDF bytes，回傳所有找到的 QR 碼字串列表"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = len(doc)
    all_qrs = []

    progress = st.progress(0, text="準備掃描…")
    status   = st.empty()

    for pn in range(total):
        pct = int((pn + 1) / total * 100)
        progress.progress(pct, text=f"掃描第 {pn+1} / {total} 頁…")
        status.caption(f"{pct}%")

        page = doc.load_page(pn)
        pix  = page.get_pixmap(dpi=DEFAULT_DPI)
        img  = cv2.imdecode(
            np.frombuffer(pix.tobytes("png"), np.uint8),
            cv2.IMREAD_COLOR
        )

        seen = set()
        for t in decode_qr_from_image(img):
            seen.add(t)
        for region in split_page_into_regions(img, DEFAULT_REGION_ROWS, DEFAULT_REGION_COLS):
            for t in decode_qr_from_image(region):
                seen.add(t)

        all_qrs.extend(seen)

    progress.progress(100, text="掃描完成 ✓")
    status.empty()
    return all_qrs


# ── Excel 產生（直接操作 XML，對齊系統參考格式）────────────────

def save_to_excel_bytes(qr_codes: list, start_date: str,
                        end_date: str, product_name: str) -> bytes:
    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    shared: list = []
    shared_map: dict = {}

    def si(val: str) -> int:
        if val not in shared_map:
            shared_map[val] = len(shared)
            shared.append(val)
        return shared_map[val]

    for h in COLUMNS:
        si(h)
    for qr in qr_codes:
        si(str(qr))
    si(str(product_name))

    sd_int = int(start_date)
    ed_int = int(end_date)

    total_rows = 1 + len(qr_codes)
    sheet_parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{NS}">',
        f'<dimension ref="A1:F{total_rows}"/>',
        '<sheetData>',
    ]

    hdr_cells = "".join(
        f'<c r="{chr(64+ci)}1" t="s"><v>{si(h)}</v></c>'
        for ci, h in enumerate(COLUMNS, 1)
    )
    sheet_parts.append(f'<row r="1">{hdr_cells}</row>')

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

    def xml_escape(s: str) -> str:
        return html.escape(str(s), quote=False)

    ss_items = "".join(f'<si><t>{xml_escape(v)}</t></si>' for v in shared)
    n = len(shared)
    ss_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="{NS}" count="{n}" uniqueCount="{n}">{ss_items}</sst>'
    )

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

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",        content_types)
        zf.writestr("_rels/.rels",                rels)
        zf.writestr("xl/workbook.xml",            wb_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/worksheets/sheet1.xml",   sheet_xml)
        zf.writestr("xl/sharedStrings.xml",       ss_xml)
        zf.writestr("xl/styles.xml",              styles_xml)

    return buf.getvalue()


# ── Streamlit UI ─────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="PDF QR碼掃描器",
        page_icon="📄",
        layout="centered"
    )

    st.title("📄 PDF QR碼掃描器")
    st.caption("自動辨識 PDF 中的 QR 碼，或手動輸入後匯出 Excel")

    st.divider()

    # ── Session state 初始化 ──
    if "qr_codes" not in st.session_state:
        st.session_state.qr_codes = []

    # ── 匯出參數 ──
    st.subheader("匯出參數")
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        product_name = st.text_input("商品名稱", placeholder="請輸入商品名稱…")
    with col2:
        start_date = st.date_input("有效起始日")
    with col3:
        end_date = st.date_input("有效結束日",
                                  value=start_date.replace(year=start_date.year + 1)
                                  if hasattr(start_date, 'year') else None)

    st.divider()

    # ── QR 來源 ──
    st.subheader("QR Code 來源")
    tab_manual, tab_pdf = st.tabs(["✏️ 手動輸入", "📄 PDF 掃描"])

    # ── 手動輸入 ──
    with tab_manual:
        st.caption("每行貼上一筆 QR 碼（可一次貼入多筆）")
        manual_text = st.text_area(
            "QR 碼內容",
            placeholder="ABC123\nDEF456\nGHI789",
            height=180,
            label_visibility="collapsed"
        )
        if st.button("👁 預覽", key="preview_manual"):
            codes = [l.strip() for l in manual_text.splitlines() if l.strip()]
            st.session_state.qr_codes = codes
            if codes:
                st.success(f"共 {len(codes)} 筆 QR 碼")
            else:
                st.warning("尚未輸入任何內容")

    # ── PDF 掃描 ──
    with tab_pdf:
        uploaded = st.file_uploader(
            "上傳 PDF 檔案",
            type=["pdf"],
            label_visibility="collapsed"
        )
        if uploaded:
            st.info(f"已選擇：{uploaded.name}　（{uploaded.size / 1024:.1f} KB）")
            if st.button("▶ 開始掃描 PDF", type="primary"):
                with st.spinner("讀取 PDF…"):
                    pdf_bytes = uploaded.read()
                try:
                    codes = scan_pdf(pdf_bytes)
                    st.session_state.qr_codes = codes
                    if codes:
                        st.success(f"掃描完成，共找到 {len(codes)} 筆 QR 碼")
                    else:
                        st.warning("未找到任何 QR 碼，請確認 PDF 內容或調高 DPI")
                except Exception as e:
                    st.error(f"掃描失敗：{e}")

    st.divider()

    # ── 結果預覽 ──
    st.subheader("掃描結果")
    codes = st.session_state.qr_codes
    if codes:
        st.caption(f"共 {len(codes)} 筆")
        preview = "\n".join(f"{i+1:>4}.  {c}" for i, c in enumerate(codes))
        st.code(preview, language=None)
    else:
        st.caption("QR 碼預覽將顯示於此…")

    st.divider()

    # ── 匯出 ──
    st.subheader("匯出 Excel")
    if not codes:
        st.info("請先輸入或掃描 QR 碼")
    else:
        sd = start_date.strftime("%Y%m%d")
        ed = end_date.strftime("%Y%m%d")
        try:
            xlsx_bytes = save_to_excel_bytes(codes, sd, ed, product_name)
            filename = f"QR_Export_{sd}_{ed}.xlsx"
            st.download_button(
                label=f"📥 下載 Excel（{len(codes)} 筆）",
                data=xlsx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            st.caption(f"商品名稱：{product_name or '（未填）'}　有效期間：{sd} ～ {ed}")
        except Exception as e:
            st.error(f"產生 Excel 時發生錯誤：{e}")


if __name__ == "__main__":
    main()
