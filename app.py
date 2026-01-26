import io, os, zipfile, tempfile
import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont
import qrcode

# ===== A4 LABEL GRID (edit if needed) =====
LABEL_COLS = 3
LABEL_ROWS = 7
EXCEL_IMG_W = 240
EXCEL_IMG_H = 380
LABEL_COL_WIDTH = 22
LABEL_ROW_HEIGHT = 300
# =========================================

APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(APP_DIR, "logo.png")

def make_qr_block(qr_data: str, building_id: str, out_png: str):
    logo = Image.open(LOGO_PATH).convert("RGBA")

    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=2)
    qr.add_data(qr_data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    qr_img = qr_img.resize((260, 260), Image.NEAREST)

    scale = min(260 / logo.width, 90 / logo.height)
    logo = logo.resize((int(logo.width * scale), int(logo.height * scale)), Image.LANCZOS)

    canvas = Image.new("RGBA", (300, 420), "white")
    canvas.paste(logo, ((300 - logo.width) // 2, 10), logo)
    canvas.paste(qr_img, ((300 - 260) // 2, 10 + logo.height + 10), qr_img)

    draw = ImageDraw.Draw(canvas)
    orange = (255, 140, 0)

    font = None
    for fp in [r"/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
               r"C:\Windows\Fonts\arialbd.ttf",
               r"C:\Windows\Fonts\calibrib.ttf"]:
        try:
            font = ImageFont.truetype(fp, 28)
            break
        except:
            pass
    if font is None:
        font = ImageFont.load_default()

    text = str(building_id)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    draw.text(((300 - tw) // 2, 420 - th - 20), text, fill=orange, font=font)

    canvas.convert("RGB").save(out_png, "PNG")

def detect_columns(ws):
    header_row = 1
    for r in range(1, 16):
        row_text = " ".join(str(ws.cell(r, c).value or "").lower() for c in range(1, ws.max_column + 1))
        if "building" in row_text and ("code" in row_text or "id" in row_text):
            header_row = r
            break

    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v:
            headers[str(v).strip().lower()] = c

    def find_col(keys):
        for k, v in headers.items():
            for key in keys:
                if key in k:
                    return v
        return None

    col_building = find_col(["building code", "building id", "building", "bldg", "code", "id"]) or 1
    col_address = find_col(["national address", "address"]) or min(col_building + 1, ws.max_column)

    col_qr = find_col(["barcode", "qr"])
    if col_qr is None:
        col_qr = ws.max_column + 1
        ws.cell(header_row, col_qr, "Barcode")

    return header_row, col_building, col_address, col_qr

def ensure_labels_sheet(wb, name="LABELS"):
    if name in wb.sheetnames:
        del wb[name]
    return wb.create_sheet(name)

def set_a4_print_settings(ws):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3

def process_one_xlsx(xlsx_bytes: bytes, original_name: str) -> bytes:
    with tempfile.TemporaryDirectory() as td:
        src_path = os.path.join(td, original_name)
        with open(src_path, "wb") as f:
            f.write(xlsx_bytes)

        wb = load_workbook(src_path)
        ws = wb.active

        try:
            ws._images = []
        except:
            pass

        header_row, col_building, col_address, col_qr = detect_columns(ws)
        qr_col_letter = get_column_letter(col_qr)
        ws.column_dimensions[qr_col_letter].width = 26

        tmp_imgs = []
        for r in range(header_row + 1, ws.max_row + 1):
            bid = ws.cell(r, col_building).value
            if bid is None or str(bid).strip() == "":
                continue

            addr = ws.cell(r, col_address).value
            qr_data = str(bid).strip()
            if addr is not None and str(addr).strip() != "":
                qr_data = qr_data + "\n" + str(addr).strip()

            png_path = os.path.join(td, f"img_{r}.png")
            make_qr_block(qr_data, bid, png_path)
            tmp_imgs.append(png_path)

            img = XLImage(png_path)
            img.width = EXCEL_IMG_W
            img.height = EXCEL_IMG_H
            ws.add_image(img, f"{qr_col_letter}{r}")
            ws.row_dimensions[r].height = 285

        labels_ws = ensure_labels_sheet(wb, "LABELS")
        set_a4_print_settings(labels_ws)

        for c in range(1, LABEL_COLS + 1):
            labels_ws.column_dimensions[get_column_letter(c)].width = LABEL_COL_WIDTH

        labels_per_page = LABEL_COLS * LABEL_ROWS
        for idx, png in enumerate(tmp_imgs):
            page = idx // labels_per_page
            pos = idx % labels_per_page
            row_in_page = pos // LABEL_COLS
            col_in_page = pos % LABEL_COLS

            base_row = page * (LABEL_ROWS + 1) + 1
            rr = base_row + row_in_page
            cc = 1 + col_in_page
            labels_ws.row_dimensions[rr].height = LABEL_ROW_HEIGHT

            cell = f"{get_column_letter(cc)}{rr}"
            img = XLImage(png)
            img.width = EXCEL_IMG_W
            img.height = EXCEL_IMG_H
            labels_ws.add_image(img, cell)

        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

# ========= UI =========
st.set_page_config(page_title="QR Excel Generator", layout="centered")
st.title("QR Code Excel Generator (Logo + A4 Labels)")

files = st.file_uploader("Upload Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if files:
    st.write(f"Files selected: {len(files)}")
    if st.button("Generate"):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in files:
                ready_bytes = process_one_xlsx(f.read(), f.name)
                out_name = f.name.replace(".xlsx", "_QR_READY.xlsx")
                z.writestr(out_name, ready_bytes)

        st.success("Done ✅")
        st.download_button(
            "Download ZIP (All QR_READY Excels)",
            data=zip_buf.getvalue(),
            file_name="QR_READY_OUTPUT.zip",
            mime="application/zip",
        )
else:
    st.info("Upload one or more .xlsx files to start.")
