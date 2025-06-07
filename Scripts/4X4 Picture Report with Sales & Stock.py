import os
from datetime import datetime
import pandas as pd
import pyodbc
from PIL import Image
from tkinter import Tk, filedialog, messagebox, Toplevel
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader
import tkinter as tk
import sys

# === CONFIG ===
IMAGE_FOLDER = r"\\192.168.1.83\jewelsoft\Ashi_Images"
PAGE_WIDTH, PAGE_HEIGHT = LETTER
STYLES_PER_ROW = 4
STYLES_PER_PAGE = 4 * STYLES_PER_ROW
BOX_WIDTH = 137
BOX_HEIGHT = 160
IMG_HEIGHT = int(BOX_HEIGHT * 0.70)
IMG_WIDTH = int(BOX_WIDTH * 0.83)
HEADER_IMAGE_PATH = r"C:\Users\vaishalisinghbaghel\Desktop\FINAL_PP_AUTOMATION\ASHI-Logo-Header.jpg"
TEXT_POS = {
    "item_cd": {"x_offset": 8, "y_offset": 12},
    "desc": {"x_offset": 5, "y_start_offset": 22, "line_spacing": 8},
    "price": {"x_offset": -5, "y_offset": -15},
    "qty": {"x_offset": 5, "bottom_margin": 2}
}

# === ESCAPE HANDLER ===
def bind_escape_kill(window):
    def kill_app(event=None):
        try:
            window.destroy()
        except Exception:
            pass
        os._exit(0)  # Force kill the process
    window.bind("<Escape>", kill_app)

def fetch_details_from_sql(item_cds):
    server = '192.168.1.11'
    database = 'jewel_ashi_prod'
    username = 'sa'
    password = 'avalondb@no1Knows'
    conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    query = '''
        SELECT style_cd, style_description, sell_price_c, style_photo
        FROM stylmsstylehd
        WHERE style_cd = ?
    '''
    result = {}

    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        for cd in item_cds:
            cursor.execute(query, cd)
            row = cursor.fetchone()
            if row:
                result[cd] = {
                    "Description": row.style_description or "",
                    "Price": row.sell_price_c or "",
                    "Photo": row.style_photo or ""
                }
            else:
                result[cd] = {"Description": "", "Price": "", "Photo": ""}
    return result

def draw_header(c, title, page_num, total_pages):
    today = datetime.today().strftime("%m/%d/%Y")
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(HexColor("#000000"))
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 30, "ASHI DIAMONDS, LLC")

    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(HexColor("#0000FF"))
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 45, title.upper())

    c.setFont("Helvetica", 8)
    c.setFillColor(HexColor("#000000"))
    c.drawString(40, PAGE_HEIGHT - 60, f"Date : {today}")
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 60,
                        "800-622-2744 | www.ashidiamonds.com | orders@ashidiamonds.com")
    c.drawRightString(PAGE_WIDTH - 40, PAGE_HEIGHT - 60, f"Page : {page_num} OF {total_pages}")

def draw_legend(c):
    c.setFont("Helvetica", 6)
    c.setFillColor(HexColor("#0000FF"))
    c.drawString(21, 40, "Sales : Blue ")
    c.setFillColor(HexColor("#000000"))
    c.drawString(61, 40, "|")
    c.setFillColor(HexColor("#FF0000"))
    c.drawString(71, 40, "Open Orders : Red ")
    c.setFillColor(HexColor("#000000"))
    c.drawString(131, 40, "|")
    c.setFillColor(HexColor("#800000"))
    c.drawString(141, 40, "BOL : Maroon ")
    c.setFillColor(HexColor("#000000"))
    c.drawString(186, 40, "|")
    c.setFillColor(HexColor("#008000"))
    c.drawString(196, 40, "Stock in Hand : Green ")

    c.setFillColor(HexColor("#000000"))
    c.setFont("Helvetica", 6)
    footer_text = (
        "18 East 48th Street   |   14th Floor   |   New York, NY 10017   |   "
        "orders@ashidiamonds.com   |   Tel: 212-319-8291   |   800-622-ASHI   |   Fax: 212-319-4341"
    )
    c.drawCentredString(PAGE_WIDTH / 2, 10, footer_text)

def draw_info_grid(c, x, y, width, row):
    col_width = width / 4
    c.setFont("Helvetica", 6)
    c.setLineWidth(0.5)
    c.setStrokeColor(HexColor("#B0B0B0"))
    c.line(x, y + 4, x + width, y + 4)

    for i in range(1, 5):
        val = row.iloc[i]
        if pd.isna(val) or val == 0:
            continue
        try: val = int(float(val))
        except: pass
        col_x = x + (i - 1) * col_width
        c.setFillColor(HexColor("#0000FF"))
        c.drawCentredString(col_x + col_width / 2, y - 5, f"{val:,}")

def draw_inline_metrics(c, x, y, width, row):
    f_val, g_val, h_val = row.iloc[5:8]
    c.setFont("Helvetica", 6)
    if pd.notna(f_val) and f_val != 0:
        try: f_val = int(float(f_val))
        except: pass
        c.setFillColor(HexColor("#FF0000"))
        c.drawString(x + 3, y, f"{f_val:,}")
    if pd.notna(g_val) and g_val != 0:
        try: g_val = int(float(g_val))
        except: pass
        c.setFillColor(HexColor("#800000"))
        c.drawRightString(x + width - 3, y + 95, f"{g_val:,}")
    if pd.notna(h_val) and h_val != 0:
        try: h_val = int(float(h_val))
        except: pass
        c.setFillColor(HexColor("#008000"))
        c.drawRightString(x + width - 3, y, f"{h_val:,}")

def draw_description_and_price(c, x, code_y, width, row):
    description = str(row["Description"]) if pd.notna(row["Description"]) else ""
    price = str(row["Price"]) if pd.notna(row["Price"]) else ""
    center_x = x + width / 2
    max_width = width - 6
    font_size = 6

    if description:
        c.setFont("Helvetica", font_size)
        c.setFillColor(HexColor("#000000"))
        words = description.split()
        lines, current_line = [], ""
        for word in words:
            test = f"{current_line} {word}".strip()
            if c.stringWidth(test, "Helvetica", font_size) <= max_width:
                current_line = test
            else:
                lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
        for i, line in enumerate(lines[:2]):
            c.drawCentredString(center_x, code_y - 3 - (i * 8), line)

    if price:
        c.setFont("Helvetica-Bold", 7)
        c.setFillColor(HexColor("#000000"))
        try:
            price_str = f"{float(price):,.0f}"
        except:
            price_str = price
        c.drawCentredString(center_x, code_y - 20, f"Price: ${price_str}")

def generate_pdf(data, output_file, title):
    c = canvas.Canvas(output_file, pagesize=LETTER)
    BOX_WIDTH, BOX_HEIGHT = 142, 165
    H_SPACING, V_SPACING = 0, 0
    IMG_HEIGHT, IMG_WIDTH = int(BOX_HEIGHT * 0.60), 100
    total_grid_width = (BOX_WIDTH * STYLES_PER_ROW)
    LEFT_MARGIN = (PAGE_WIDTH - total_grid_width) / 2
    total_pages = (len(data) - 1) // STYLES_PER_PAGE + 1

    for page in range(total_pages):
        y = PAGE_HEIGHT - 80
        draw_header(c, title, page + 1, total_pages)
        x = LEFT_MARGIN
        page_data = data.iloc[page * STYLES_PER_PAGE:(page + 1) * STYLES_PER_PAGE]

        for i, row in page_data.iterrows():
            style_code = row.iloc[0]
            image_file = row.get("Photo", "")
            image_path = os.path.join(IMAGE_FOLDER, image_file) if image_file else ""

            c.setStrokeColor(HexColor("#B0B0B0"))
            c.setLineWidth(1)
            c.rect(x, y - BOX_HEIGHT, BOX_WIDTH, BOX_HEIGHT, fill=0)

            center_x = x + BOX_WIDTH / 2
            image_x, image_y = x + (BOX_WIDTH - IMG_WIDTH) / 2, y - IMG_HEIGHT - 5

            if os.path.exists(image_path):
                try:
                    with Image.open(image_path) as img:
                        img.verify()
                        img = Image.open(image_path)
                        aspect_ratio = img.width / img.height
                        if aspect_ratio > 1:
                            display_width = IMG_WIDTH
                            display_height = IMG_WIDTH / aspect_ratio
                        else:
                            display_height = IMG_HEIGHT
                            display_width = IMG_HEIGHT * aspect_ratio
                        draw_x = x + (BOX_WIDTH - display_width) / 2
                        draw_y = y - display_height - 5
                        c.drawImage(ImageReader(image_path), draw_x, draw_y,
                                    width=display_width, height=display_height)
                except Exception as e:
                    print(f" Image load error: {e} - {image_path}")

            grid_y = image_y - 10
            draw_info_grid(c, x, grid_y, BOX_WIDTH, row)
            draw_inline_metrics(c, x, image_y, BOX_WIDTH, row)

            code_y = grid_y - 25
            c.setFont("Helvetica-Bold", 8)
            c.setFillColor(HexColor("#000000"))
            c.line(x, code_y + 15, x + BOX_WIDTH, code_y + 15)
            c.drawCentredString(center_x, code_y + 5, style_code.upper())

            draw_description_and_price(c, x, code_y, BOX_WIDTH, row)

            x += BOX_WIDTH + H_SPACING
            if (i % STYLES_PER_ROW) == STYLES_PER_ROW - 1:
                x = LEFT_MARGIN
                y -= BOX_HEIGHT + V_SPACING

        draw_legend(c)
        c.showPage()

    c.save()

# === CUSTOM INFO DIALOG ===
def show_input_format_info(callback_on_close):
    info_window = Toplevel()
    info_window.title("Input File Format Information")
    info_window.geometry("700x300")
    info_window.resizable(False, False)

    # Bind ESC to kill everything
    bind_escape_kill(info_window)

    message = (
        "Please ensure your input Excel file has the following columns:\n\n"
        "Column A: STYLE CD\n"
        "Columns B-E: Annual Net Quantity for the past 4 years\n"
        "(As: ANNUAL ALL NET QTY 2022, 2023, 2024, 2025)\n"
        "Column F: CUST OPEN SALES ORDER QTY\n"
        "Column G: TOTAL OPEN BOL QTY\n"
        "Column H: STOCK IN HAND\n\n"
        "The script will read data from these columns."
    )

    label = tk.Label(info_window, text=message, justify="left", wraplength=480, padx=10, pady=10)
    label.pack()

    def on_ok():
        info_window.destroy()
        callback_on_close()

    ok_btn = tk.Button(info_window, text="OK", width=15, command=on_ok)
    ok_btn.pack(pady=10)

    info_window.grab_set()
    info_window.mainloop()

# === MAIN ENTRY ===
def main():
    root = Tk()
    root.withdraw()
    # Bind ESC to kill everything
    bind_escape_kill(root)

    def proceed_after_info():
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        if not file_path:
            messagebox.showerror("Error", "No file selected!")
            return

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file:\n{e}")
            return

        file_name_only = os.path.splitext(os.path.basename(file_path))[0]
        item_cds = df.iloc[:, 0].dropna().astype(str).tolist()

        try:
            details_dict = fetch_details_from_sql(item_cds)
        except Exception as e:
            messagebox.showerror("Error", f"SQL fetch failed:\n{e}")
            return

        df["Description"] = [details_dict.get(cd, {}).get("Description", "") for cd in item_cds]
        df["Price"] = [details_dict.get(cd, {}).get("Price", "") for cd in item_cds]
        df["Photo"] = [details_dict.get(cd, {}).get("Photo", "") for cd in item_cds]

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save PDF As",
            initialfile=file_name_only + "_output.pdf"
        )
        if not output_path:
            messagebox.showinfo("Cancelled", "No output location selected.")
            return

        generate_pdf(df, output_path, title=file_name_only.upper())
        messagebox.showinfo("Success", f"PDF generated successfully at:\n{output_path}")

    show_input_format_info(proceed_after_info)

if __name__ == "__main__":
    main()