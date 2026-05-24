import streamlit as st
import openpyxl
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import zipfile
import io
import base64
import os
import requests
from datetime import datetime

# ─── Cấu hình trang ───────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Tạo Phiếu Tự động - Phòng Mua Hàng",
    page_icon="📦",
    layout="centered"
)
st.image("https://static.ybox.vn/2023/7/4/1689816973012-LOGO.jpg", width=120)
st.title("📦 TẠO PHIẾU XUẤT KHO TỰ ĐỘNG")
st.caption("Upload file Excel → Tự động tạo phiếu Word cho từng trung tâm")

# ─── Logo URL ─────────────────────────────────────────────────────────────────
LOGO_URL = "https://static.ybox.vn/2023/7/4/1689816973012-LOGO.jpg"

def load_logo_from_url(url):
    try:
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        return response.content
    except Exception:
        return None

# ─── Helpers ──────────────────────────────────────────────────────────────────
def add_header(doc, logo_bytes):
    """Logo góc trái header"""
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    para = header.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    if logo_bytes:
        try:
            run = para.add_run()
            run.add_picture(io.BytesIO(logo_bytes), height=Cm(1.2))
        except Exception:
            pass

def set_cell_border(cell, color="CCCCCC"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        tag = OxmlElement(f"w:{edge}")
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), "4")
        tag.set(qn("w:color"), color)
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def fmt_number(n):
    try:
        return f"{int(n):,}".replace(",", ".")
    except:
        return "0"

def make_ma_phieu(dt):
    return f"PK-{dt.strftime('%Y%m%d%H%M%S')}"

def add_footer(doc):
    """Thêm footer: Phòng Mua Hàng bên trái, số trang bên phải"""
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False

    # Xóa paragraph mặc định
    for p in footer.paragraphs:
        p.clear()

    # Lấy paragraph đầu tiên của footer
    para = footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)

    # Thêm tab stop ở giữa và cuối để căn phải số trang
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    pPr = para._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")

    tab_right = OxmlElement("w:tab")
    tab_right.set(qn("w:val"), "right")
    tab_right.set(qn("w:pos"), "9026")  # A4 content width ~17cm
    tabs.append(tab_right)
    pPr.append(tabs)

    # Vẽ đường kẻ trên footer
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "6")
    top.set(qn("w:color"), "CCCCCC")
    pBdr.append(top)
    pPr.append(pBdr)

    # Text "Phòng Mua Hàng" bên trái
    r1 = para.add_run("Phòng Mua Hàng")
    r1.font.size = Pt(9)
    r1.font.name = "Arial"
    r1.font.color.rgb = RGBColor(128, 128, 128)

    # Tab để đẩy sang phải
    para.add_run("\t")
    para.add_run("\t")

    # Số trang: "Trang X/Y"
    r2 = para.add_run("Trang ")
    r2.font.size = Pt(9)
    r2.font.name = "Arial"
    r2.font.color.rgb = RGBColor(128, 128, 128)

    # Field PAGE
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r_page = para.add_run()
    r_page._r.append(fld_begin)

    instr = OxmlElement("w:instrText")
    instr.text = " PAGE "
    r_instr = para.add_run()
    r_instr._r.append(instr)

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r_end = para.add_run()
    r_end._r.append(fld_end)

    for r in [r_page, r_instr, r_end]:
        r.font.size = Pt(9)
        r.font.name = "Arial"
        r.font.color.rgb = RGBColor(128, 128, 128)

    r3 = para.add_run("/")
    r3.font.size = Pt(9)
    r3.font.name = "Arial"
    r3.font.color.rgb = RGBColor(128, 128, 128)

    # Field NUMPAGES
    fld_begin2 = OxmlElement("w:fldChar")
    fld_begin2.set(qn("w:fldCharType"), "begin")
    r_np = para.add_run()
    r_np._r.append(fld_begin2)

    instr2 = OxmlElement("w:instrText")
    instr2.text = " NUMPAGES "
    r_instr2 = para.add_run()
    r_instr2._r.append(instr2)

    fld_end2 = OxmlElement("w:fldChar")
    fld_end2.set(qn("w:fldCharType"), "end")
    r_end2 = para.add_run()
    r_end2._r.append(fld_end2)

    for r in [r_np, r_instr2, r_end2]:
        r.font.size = Pt(9)
        r.font.name = "Arial"
        r.font.color.rgb = RGBColor(128, 128, 128)

# ─── Đọc Excel ────────────────────────────────────────────────────────────────
def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    all_data = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 3:
            continue

        header_row = rows[0]
        sub_row = rows[1]

        has_ncc = False
        col_ten_hang = 0
        col_gia_cost = 1
        col_dvt = None
        col_ncc = None
        col_branch_start = 2

        if sub_row[0] and str(sub_row[0]).strip() == "NCC":
            has_ncc = True
            col_ncc = 0
            col_ten_hang = 1
            col_dvt = 2
            col_gia_cost = 3
            col_branch_start = 4

        branches = {}
        for c, val in enumerate(header_row):
            if val and str(val).strip() not in ("", "Tổng order", "None"):
                branches[str(val).strip()] = c

        products = []
        for row in rows[3:]:
            if not row:
                continue
            ten_hang = row[col_ten_hang] if col_ten_hang < len(row) else None
            if not ten_hang:
                continue
            ten_hang = str(ten_hang).strip()
            if not ten_hang or "Tổng Cộng" in ten_hang:
                continue
            if "CHIẾT KHẤU" in ten_hang.upper() or "TT XÁC NHẬN" in ten_hang.upper():
                continue

            gia_cost = row[col_gia_cost] if col_gia_cost < len(row) else 0
            dvt = str(row[col_dvt]).strip() if col_dvt and col_dvt < len(row) and row[col_dvt] else ""
            ncc = str(row[col_ncc]).strip() if has_ncc and col_ncc < len(row) and row[col_ncc] else "Tiến Son"
            if not ncc or ncc.lower() == "none":
                ncc = "Tiến Son"

            branch_qty = {}
            for branch, col in branches.items():
                if col < len(row):
                    try:
                        qty = float(row[col] or 0)
                        if qty > 0:
                            branch_qty[branch] = int(qty)
                    except:
                        pass

            if branch_qty:
                products.append({
                    "ten_hang": ten_hang,
                    "gia_cost": int(gia_cost) if gia_cost else 0,
                    "dvt": dvt,
                    "ncc": ncc,
                    "branch_qty": branch_qty,
                })

        all_data[sheet_name] = {
            "branches": list(branches.keys()),
            "products": products,
        }

    return all_data

# ─── Tạo một phiếu Word ───────────────────────────────────────────────────────
def build_phieu(branch_name, items, now, logo_bytes=None):
    doc = Document()

    # Căn lề trang A4
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin    = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin   = Cm(2)
    section.right_margin  = Cm(1.5)

    # ── Header + Footer ──────────────────────────────────────────────────────
    add_header(doc, logo_bytes)
    add_footer(doc)

    def add_para(text, bold=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT,
                 color=None, space_before=0, space_after=4):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after  = Pt(space_after)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Arial"
        if color:
            run.font.color.rgb = RGBColor(*color)
        return p

    # Tiêu đề
    add_para("PHIẾU XUẤT KHO", bold=True, size=18,
             align=WD_ALIGN_PARAGRAPH.CENTER,
             color=(43, 84, 126), space_after=2)

    # Đường kẻ dưới tiêu đề
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_after = Pt(8)
    pPr = p_line._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:color"), "2B547E")
    pBdr.append(bottom)
    pPr.append(pBdr)

    ma_phieu = make_ma_phieu(now)
    ngay_tao = now.strftime("%d/%m/%Y")
    # Lấy NCC từ item đầu tiên, viết HOA toàn bộ
    ncc = items[0]["ncc"].upper() if items else "TIẾN SON"

    # Thông tin phiếu
    def info_line(label, value, value_color=None, value_bold=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        r1 = p.add_run(label)
        r1.bold = True
        r1.font.size = Pt(11)
        r1.font.name = "Arial"
        r2 = p.add_run(value)
        r2.font.size = Pt(11)
        r2.font.name = "Arial"
        if value_color:
            r2.font.color.rgb = RGBColor(*value_color)
        if value_bold:
            r2.bold = True

    info_line("Mã phiếu: ", f"{ma_phieu}          Ngày: {ngay_tao}")
    # "Chi nhánh" → "Trung Tâm"
    info_line("Trung Tâm nhận: ", branch_name, (43, 84, 126), value_bold=True)
    # NCC viết HOA, màu đỏ nổi bật
    info_line("Nhà cung cấp: ", ncc, (192, 0, 0), value_bold=True)

    # Ghi chú đỏ
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(2)
    p_note.paragraph_format.space_after  = Pt(10)
    r_label = p_note.add_run("Ghi chú: ")
    r_label.bold = True
    r_label.font.size = Pt(11)
    r_label.font.name = "Arial"
    r_val = p_note.add_run("Nhập hàng Link HPC (Không Kiot)")
    r_val.italic = True
    r_val.font.size = Pt(11)
    r_val.font.name = "Arial"
    r_val.font.color.rgb = RGBColor(192, 0, 0)

    # Bảng hàng hóa
    col_widths = [Cm(1.0), Cm(6.5), Cm(1.5), Cm(2.0), Cm(2.5), Cm(2.5)]
    headers    = ["STT", "Tên hàng", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền"]

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"

    hdr = table.rows[0]
    for i, (cell, text, w) in enumerate(zip(hdr.cells, headers, col_widths)):
        cell.width = w
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_bg(cell, "2B547E")
        set_cell_border(cell, "2B547E")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after  = Pt(3)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = "Arial"
        run.font.color.rgb = RGBColor(255, 255, 255)

    total_qty    = 0
    total_amount = 0

    for idx, item in enumerate(items):
        qty    = item["qty"]
        amount = qty * item["gia_cost"]
        total_qty    += qty
        total_amount += amount
        shade = "F5F8FC" if idx % 2 == 0 else "FFFFFF"

        row = table.add_row()
        values = [
            str(idx + 1),
            item["ten_hang"],
            item["dvt"],
            fmt_number(qty),
            fmt_number(item["gia_cost"]),
            fmt_number(amount),
        ]
        aligns = [
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ]
        for cell, text, w, align in zip(row.cells, values, col_widths, aligns):
            cell.width = w
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_bg(cell, shade)
            set_cell_border(cell, "CCCCCC")
            p = cell.paragraphs[0]
            p.alignment = align
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            run = p.add_run(text)
            run.font.size = Pt(10)
            run.font.name = "Arial"

    # Dòng tổng cộng
    row_total = table.add_row()
    cells = row_total.cells

    cells[0].merge(cells[1]).merge(cells[2])
    set_cell_bg(cells[0], "EFF4FA")
    set_cell_border(cells[0], "2B547E")
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(3)
    run = p.add_run("TỔNG CỘNG")
    run.bold = True
    run.font.size = Pt(10)
    run.font.name = "Arial"
    run.font.color.rgb = RGBColor(43, 84, 126)

    for cell, text in [(cells[3], fmt_number(total_qty)),
                       (cells[4], ""),
                       (cells[5], fmt_number(total_amount))]:
        set_cell_bg(cell, "2B547E")
        set_cell_border(cell, "2B547E")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after  = Pt(3)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = "Arial"
        run.font.color.rgb = RGBColor(255, 255, 255)

    # ── Ký tên: chỉ Người Gửi và Người Nhận ────────────────────────────────
    doc.add_paragraph().paragraph_format.space_after = Pt(16)

    sig_table = doc.add_table(rows=1, cols=2)
    sig_labels = ["Người Gửi", "Người Nhận"]
    for cell, label in zip(sig_table.rows[0].cells, sig_labels):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for edge in ("top", "left", "bottom", "right"):
            tag = OxmlElement(f"w:{edge}")
            tag.set(qn("w:val"), "none")
            tcBorders.append(tag)
        tcPr.append(tcBorders)

        p1 = cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(label)
        r1.bold = True
        r1.font.size = Pt(10)
        r1.font.name = "Arial"

        p2 = cell.add_paragraph("(Ký, ghi rõ họ tên)")
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.runs[0].italic = True
        p2.runs[0].font.size = Pt(9)
        p2.runs[0].font.name = "Arial"
        p2.runs[0].font.color.rgb = RGBColor(128, 128, 128)


    # ── Footer ──────────────────────────────────────────────────────────────
    add_footer(doc)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─── Tạo file Excel mẫu ──────────────────────────────────────────────────────
def create_sample_excel():
    wb = openpyxl.Workbook()

    # ── Sheet 1: Dạng thường (không có cột NCC) ──────────────────────────────
    ws1 = wb.active
    ws1.title = "Sheet_Thuong"

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill("solid", fgColor="2B547E")
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border  = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )

    # Hàng 1: Tên chi nhánh (từ cột C trở đi)
    ws1["A1"] = "Tên hàng"
    ws1["B1"] = "Đơn giá (cost)"
    ws1["C1"] = "Trung Tâm A"
    ws1["D1"] = "Trung Tâm B"
    ws1["E1"] = "Trung Tâm C"

    for cell in ws1[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    # Hàng 2: Sub-header
    ws1["A2"] = "Tên hàng"
    ws1["B2"] = "Giá cost"
    ws1["C2"] = "Số đặt"
    ws1["D2"] = "Số đặt"
    ws1["E2"] = "Số đặt"

    # Hàng 3: Tiêu đề cột (bỏ trống hoặc ghi chú)
    ws1["A3"] = "(Bỏ trống hoặc ghi chú)"

    # Dữ liệu mẫu từ hàng 4
    sample_data = [
        ("Sữa tươi Vinamilk 1L", 28000, 10, 5, 8),
        ("Bánh mì sandwich", 15000, 20, 15, 12),
        ("Nước suối Lavie 500ml", 5000, 50, 30, 40),
        ("Snack khoai tây", 12000, 15, 10, 20),
    ]
    for row_data in sample_data:
        ws1.append(list(row_data))

    # Điều chỉnh độ rộng cột
    ws1.column_dimensions["A"].width = 30
    ws1.column_dimensions["B"].width = 18
    for col in ["C", "D", "E"]:
        ws1.column_dimensions[col].width = 14

    # ── Sheet 2: Dạng có NCC (như sheet Kids) ────────────────────────────────
    ws2 = wb.create_sheet("Sheet_CoNCC")

    ws2["A1"] = "NCC"
    ws2["B1"] = "Tên hàng"
    ws2["C1"] = "ĐVT"
    ws2["D1"] = "Đơn giá (cost)"
    ws2["E1"] = "Trung Tâm A"
    ws2["F1"] = "Trung Tâm B"
    ws2["G1"] = "Trung Tâm C"

    for cell in ws2[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    ws2["A2"] = "NCC"
    ws2["B2"] = "Tên hàng"
    ws2["C2"] = "ĐVT"
    ws2["D2"] = "Giá cost"
    ws2["E2"] = "Số đặt"
    ws2["F2"] = "Số đặt"
    ws2["G2"] = "Số đặt"

    ws2["A3"] = "(Bỏ trống)"

    sample_data2 = [
        ("Vinamilk",   "Sữa bột Dielac 900g",    "Hộp",   285000, 5, 3, 4),
        ("TH True",    "Sữa tươi TH 180ml",       "Hộp",   7500,  30, 20, 25),
        ("Nestle",     "Milo hộp 180ml",           "Hộp",   9500,  20, 15, 18),
        ("Abbott",     "Ensure Gold 850g",         "Lon",   450000, 2, 1, 3),
    ]
    for row_data in sample_data2:
        ws2.append(list(row_data))

    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 28
    ws2.column_dimensions["C"].width = 10
    ws2.column_dimensions["D"].width = 18
    for col in ["E", "F", "G"]:
        ws2.column_dimensions[col].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─── Giao diện chính ──────────────────────────────────────────────────────────

# ── Hướng dẫn & file Excel mẫu ──────────────────────────────────────────────
with st.expander("📋 Hướng dẫn sử dụng & tải file Excel mẫu", expanded=False):
    st.markdown("""
### 📌 Cấu trúc file Excel

App hỗ trợ **2 dạng sheet**:

---

#### 🔹 Dạng 1 — Sheet thường (không có cột NCC)

| Cột A | Cột B | Cột C | Cột D | ... |
|---|---|---|---|---|
| **Tên hàng** | **Đơn giá** | **Trung Tâm A** | **Trung Tâm B** | ... |
| *(sub-header)* | *(sub-header)* | Số đặt | Số đặt | ... |
| *(bỏ trống)* | | | | |
| Sữa tươi 1L | 28000 | 10 | 5 | ... |

- **Hàng 1**: Tên cột + Tên từng Trung Tâm (từ cột C trở đi)
- **Hàng 2**: Sub-header (ví dụ: "Số đặt")
- **Hàng 3**: Bỏ trống hoặc ghi chú
- **Hàng 4 trở đi**: Dữ liệu hàng hóa

---

#### 🔹 Dạng 2 — Sheet có cột NCC (dành cho nhà cung cấp riêng)

| Cột A | Cột B | Cột C | Cột D | Cột E | Cột F | ... |
|---|---|---|---|---|---|---|
| **NCC** | **Tên hàng** | **ĐVT** | **Đơn giá** | **TT A** | **TT B** | ... |
| NCC | Tên hàng | ĐVT | Giá cost | Số đặt | Số đặt | ... |
| *(bỏ trống)* | | | | | | |
| Vinamilk | Sữa bột 900g | Hộp | 285000 | 5 | 3 | ... |

- **Hàng 1, ô A1 = "NCC"** → App tự nhận dạng dạng 2
- Mỗi hàng hóa ghi tên NCC ở cột A
- NCC sẽ được in **ĐỎ, VIẾT HOA** trên phiếu

---

### ⚠️ Lưu ý quan trọng

- Ô bỏ trống hoặc số lượng = 0 → **không tạo phiếu cho Trung Tâm đó**
- Dòng chứa "Tổng Cộng", "CHIẾT KHẤU", "TT XÁC NHẬN" → **tự động bỏ qua**
- Tên Trung Tâm **không được** đặt trùng nhau trong cùng 1 sheet
- Mỗi sheet = 1 nhóm hàng (ví dụ: sheet "Kids", sheet "Food")
    """)

    sample_bytes = create_sample_excel()
    st.download_button(
        label="⬇️ Tải file Excel mẫu",
        data=sample_bytes,
        file_name="Excel_Mau_Phieu_Xuat_Kho.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()

# ── Upload file ──────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Kéo thả hoặc chọn file Excel",
    type=["xlsx"],
    help="File Excel có các sheet chứa danh sách hàng theo trung tâm"
)

# Đọc logo một lần
logo_bytes = load_logo_from_url(LOGO_URL)

if uploaded:
    st.success(f"Đã tải file: **{uploaded.name}**")

    with st.spinner("Đang đọc file Excel..."):
        try:
            all_data = parse_excel(uploaded.read())
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")
            st.stop()

    total_branches = sum(
        len([b for b in data["branches"]
             if any(b in p["branch_qty"] for p in data["products"])])
        for data in all_data.values()
    )
    total_sheets = len(all_data)

    col1, col2 = st.columns(2)
    col1.metric("Số sheet", total_sheets)
    col2.metric("Tổng số phiếu sẽ tạo", total_branches)

    st.divider()

    sheet_names = list(all_data.keys())
    selected_sheets = st.multiselect(
        "Chọn sheet cần tạo phiếu",
        options=sheet_names,
        default=sheet_names,
    )

    if st.button("🖨️ Tạo phiếu xuất kho", type="primary", use_container_width=True):
        now = datetime.now()
        zip_buf = io.BytesIO()
        count = 0

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            progress = st.progress(0, text="Đang tạo phiếu...")

            tasks = []
            for sheet_name in selected_sheets:
                data = all_data.get(sheet_name, {})
                for branch in data.get("branches", []):
                    items = [
                        {**p, "qty": p["branch_qty"][branch]}
                        for p in data["products"]
                        if branch in p["branch_qty"]
                    ]
                    if items:
                        tasks.append((sheet_name, branch, items))

            for i, (sheet_name, branch, items) in enumerate(tasks):
                docx_bytes = build_phieu(branch, items, now, logo_bytes)

                # Tên file: NCC_Tên Trung Tâm_Mã phiếu
                ncc_name = items[0]["ncc"].upper() if items else "NCC"
                ma_phieu = make_ma_phieu(now)
                safe_ncc    = ncc_name.replace(" ", "_").replace("/", "-")
                safe_branch = branch.strip().replace("/", "-").replace("\\", "-")
                safe_ma     = ma_phieu.replace(":", "-")
                filename = f"{safe_ncc}_{safe_branch}_{safe_ma}.docx"

                zf.writestr(filename, docx_bytes)
                count += 1
                progress.progress((i + 1) / len(tasks),
                                   text=f"Đang tạo: {branch} ({i+1}/{len(tasks)})")

        progress.empty()
        zip_buf.seek(0)

        st.success(f"✅ Đã tạo xong **{count} phiếu xuất kho**!")

        zip_name = f"phieu_xuat_kho_{now.strftime('%Y%m%d_%H%M%S')}.zip"
        st.download_button(
            label=f"⬇️ Tải về {zip_name}",
            data=zip_buf,
            file_name=zip_name,
            mime="application/zip",
            use_container_width=True,
        )

else:
    st.info("👆 Upload file Excel để bắt đầu. Xem hướng dẫn và tải file mẫu ở trên ☝️")
