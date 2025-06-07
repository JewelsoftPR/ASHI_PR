import pandas as pd
import pyodbc
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
from reportlab.graphics.barcode import code128
from tkinter import Tk, filedialog, Toplevel, Label, messagebox, Button
from tkinter.ttk import Progressbar
from datetime import datetime
import os
from PIL import Image

# === PAGE SETTINGS ===
PAGE_WIDTH, PAGE_HEIGHT = letter
ROWS_PER_PAGE = 4
STYLES_PER_ROW = 4
STYLES_PER_PAGE = ROWS_PER_PAGE * STYLES_PER_ROW
IMAGE_NETWORK_PATH = r"\\192.168.1.83\jewelsoft\Ashi_Images"
HEADER_IMAGE_PATH = r"C:\Users\vaishalisinghbaghel\Desktop\FINAL_PP_AUTOMATION\ASHI-Logo-Header.jpg"

# === SQL CONNECTION SETTINGS ===
SQL_SERVER = "192.168.1.11"
SQL_DATABASE = "jewel_ashi_prod"
SQL_USERNAME = "sa"
SQL_PASSWORD = "avalondb@no1Knows"

SQL_CONNECTION_STRING = (
    f"DRIVER={{SQL Server}};"
    f"SERVER={SQL_SERVER};"
    f"DATABASE={SQL_DATABASE};"
    f"UID={SQL_USERNAME};"
    f"PWD={SQL_PASSWORD};"
)

def get_style_photo_from_db(style_cd):
    try:
        with pyodbc.connect(SQL_CONNECTION_STRING) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT style_photo FROM stylmsstylehd WHERE style_cd = ?", (style_cd,))
            result = cursor.fetchone()
            if result and result[0]:
                return result[0]
    except Exception as e:
        print(f"Database error for style_cd {style_cd}: {e}")
    return None

def find_image_path(style_cd):
    photo_filename = get_style_photo_from_db(style_cd)
    if photo_filename:
        image_path = os.path.join(IMAGE_NETWORK_PATH, photo_filename)
        if os.path.exists(image_path):
            return image_path
    return None

def show_progress_bar(total_steps):
    window = Toplevel()
    window.title("Generating PDFs...")
    window.geometry("400x100")
    Label(window, text="Generating ASHI Barcode PDFs...").pack(pady=10)
    bar = Progressbar(window, length=350, mode='determinate', maximum=total_steps)
    bar.pack(pady=5)
    return window, bar

def draw_header(c, title):
    header_width = PAGE_WIDTH - 50
    header_height = 60
    header_x = 25
    header_y = PAGE_HEIGHT - header_height
    try:
        c.drawImage(HEADER_IMAGE_PATH, header_x, header_y + 3, width=header_width, height=header_height - 10, preserveAspectRatio=False, mask='auto')
    except:
        pass
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(HexColor("#1F4E79"))
    c.drawCentredString(PAGE_WIDTH / 2, header_y - 17, title.upper())

def draw_footer(c, page_num, total_pages):
    today = datetime.today().strftime("%m/%d/%Y")
    c.setFont("Helvetica", 8)
    c.setFillColor(HexColor("#000000"))
    bottom_margin = 13
    c.drawString(40, bottom_margin, f"{today}")
    c.drawRightString(PAGE_WIDTH - 30, bottom_margin, f"Page  {page_num} of {total_pages}")

def generate_pdf(data, output_file, title, progress_bar=None):
    c = canvas.Canvas(output_file, pagesize=letter)
    BOX_WIDTH = 142
    BOX_HEIGHT = 170
    H_SPACING = 0
    V_SPACING = 0
    IMG_HEIGHT = int(BOX_HEIGHT * 0.60)
    IMG_WIDTH = 98
    total_grid_width = (BOX_WIDTH * STYLES_PER_ROW) + (H_SPACING * (STYLES_PER_ROW - 1))
    LEFT_MARGIN = (PAGE_WIDTH - total_grid_width) / 2
    total_pages = (len(data) - 1) // STYLES_PER_PAGE + 1
    style_count = 0

    for page in range(total_pages):
        draw_header(c, title)
        x = LEFT_MARGIN
        y = PAGE_HEIGHT - 88
        start = page * STYLES_PER_PAGE
        end = start + STYLES_PER_PAGE
        page_data = data[start:end]

        for i, row in page_data.iterrows():
            style_code = row['STYLE_CD']
            barcode_value = str(row['BARCODE']).strip('*')
            image_path = find_image_path(style_code)

            c.setStrokeColor(HexColor("#B0B0B0"))
            c.setLineWidth(1)
            c.rect(x, y - BOX_HEIGHT, BOX_WIDTH, BOX_HEIGHT, fill=0)
            center_x = x + BOX_WIDTH / 2

            if image_path:
                try:
                    with Image.open(image_path) as img:
                        img_width, img_height = img.size
                        aspect_ratio = img_width / img_height
                        if img_width < 400 or img_height < 400:
                            upscale_factor = 1.5
                            img_width = int(img_width * upscale_factor)
                            img_height = int(img_height * upscale_factor)
                            aspect_ratio = img_width / img_height
                        if aspect_ratio > 1:
                            display_width = IMG_WIDTH
                            display_height = IMG_WIDTH / aspect_ratio
                        else:
                            display_height = IMG_HEIGHT
                            display_width = IMG_HEIGHT * aspect_ratio
                        draw_x = x + (BOX_WIDTH - display_width) / 2
                        draw_y = y - display_height - 5
                        c.drawImage(ImageReader(image_path), draw_x, draw_y, width=display_width, height=display_height, preserveAspectRatio=True, mask='auto')

                        price_value = row.get('PRICE_CX3', '')
                        if pd.notna(price_value) and price_value != '':
                            formatted_price = f"${int(price_value):,}"
                            c.setFont("Helvetica-Bold", 7)
                            c.setFillColorRGB(0, 0, 0)
                            c.drawString(draw_x + display_width - 10, draw_y, formatted_price)
                except Exception as e:
                    print(f"Image error: {e}")

            separator_y = y - IMG_HEIGHT - 8
            c.setLineWidth(0.5)
            c.line(x - 0.5, separator_y, x + BOX_WIDTH - 1, separator_y)

            text_y = separator_y - 8
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(center_x, text_y, style_code.upper())

            description = row.get('STYLE_DESCRIPTION', '')
            if pd.notna(description) and str(description).strip() != '':
                full_desc = str(description).strip()
                if len(full_desc) > 20:
                    words = full_desc.split()
                    first_line = ""
                    second_line = ""
                    for word in words:
                        if len(first_line + " " + word) <= 20:
                            first_line += (" " if first_line else "") + word
                        else:
                            second_line += (" " if second_line else "") + word
                    c.setFont("Helvetica", 6)
                    c.drawCentredString(center_x, text_y - 10, first_line.strip())
                    c.drawCentredString(center_x, text_y - 20, second_line.strip())
                else:
                    c.setFont("Helvetica", 8)
                    c.drawCentredString(center_x, text_y - 10, full_desc)
            else:
                c.setFont("Helvetica", 8)
                c.drawCentredString(center_x, text_y - 10, "DESCRIPTION MISSING")

            try:
                barcode = code128.Code128(barcode_value, barHeight=17, barWidth=0.8)
                c.setFillColor(HexColor("#333333"))
                barcode.drawOn(c, x + (BOX_WIDTH - barcode.width) / 2, text_y - 45)
                c.setFillColor(HexColor("#000000"))
                spaced_text = '    '.join(barcode_value)
                c.setFont("Helvetica", 6)
                c.drawCentredString(center_x, text_y - 50, spaced_text)
            except:
                c.setFont("Helvetica", 6)
                c.drawCentredString(center_x, text_y - 55, "Invalid barcode")

            x += BOX_WIDTH + H_SPACING
            if (i - start + 1) % STYLES_PER_ROW == 0:
                x = LEFT_MARGIN
                y -= BOX_HEIGHT + V_SPACING

            style_count += 1
            if progress_bar:
                progress_bar['value'] = style_count
                progress_bar.update()

        draw_footer(c, page + 1, total_pages)
        c.showPage()

    c.save()

def start_process():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv")])
    if not file_paths:
        print("No data file selected. Exiting.")
        return

    output_base_dir = filedialog.askdirectory(title="Select Output Folder")
    if not output_base_dir:
        print("No destination selected. Exiting.")
        return

    output_dir = os.path.join(output_base_dir, "generated_pdfs_barcode")
    os.makedirs(output_dir, exist_ok=True)

    for file_path in file_paths:
        if file_path.lower().endswith(('.xlsx', '.xls')):
            xl = pd.ExcelFile(file_path)
            sheet_name = xl.sheet_names[0]
            df = xl.parse(sheet_name)
        else:
            df = pd.read_csv(file_path)
            sheet_name = os.path.splitext(os.path.basename(file_path))[0]

        expected_columns = {
            'id': 'STYLE_CD',
            'barcode': 'BARCODE',
            'price': 'PRICE_CX3',
            'DESCRIPTION': 'STYLE_DESCRIPTION'
        }

        for col in expected_columns:
            if col not in df.columns:
                messagebox.showerror("Missing Column", f"Column '{col}' is required in file: {os.path.basename(file_path)}")
                continue

        df = df.rename(columns=expected_columns)
        df = df.dropna(subset=['STYLE_CD', 'BARCODE'])

        progress_window, progress_bar = show_progress_bar(total_steps=len(df))

        safe_name = "".join(c if c.isalnum() or c in (' ', '.', '-', '_') else "_" for c in sheet_name)
        output_pdf_path = os.path.join(output_dir, f"{safe_name}_barcode.pdf")

        generate_pdf(df, output_pdf_path, sheet_name, progress_bar)
        progress_window.destroy()

    messagebox.showinfo("Success", f"✅ Barcode PDFs generated successfully in folder:\n{output_dir}")

def main():
    root = Tk()
    root.withdraw()

    def proceed():
        dialog.destroy()
        start_process()

    dialog = Toplevel()
    dialog.title("Instructions")
    dialog.geometry("600x250")
    
    # Regular text label
    Label(dialog, text="This Script Requires The Following Columns In The Input Sheet To Generate The Report:\n\nID, Price, Description and Barcode\n\n", 
          wraplength=450, justify="left").pack(pady=(20,0))
    
    # Note text in red
    Label(dialog, text="NOTE: Make Sure To Sort The Data (A-Z or Smallest to largest ID) Before Proceeding To Generate The Report.", 
          wraplength=450, justify="left", fg="red").pack()
    
    # OK button with border
    ok_button = Button(dialog, text="OK", command=proceed, relief="solid", borderwidth=1)
    ok_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()