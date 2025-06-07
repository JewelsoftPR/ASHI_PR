import pandas as pd
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
IMAGE_FOLDER = r"D:\\Ashi_Images"
HEADER_IMAGE_PATH = r"C:\Users\vaishalisinghbaghel\Desktop\FINAL_PP_AUTOMATION\ASHI-Logo-Header.jpg"
SUPPORTED_EXTENSIONS = ['.jpg', '.jpeg', '.png']

# === INITIAL INFO DIALOG ===
def show_initial_info_dialog():
    window = Toplevel()
    window.title("Input Excel File Format")
    window.geometry("600x300")
    window.resizable(False, False)

    info_text = (
        "Please ensure your input Excel file contains the following columns:\n\n"
        "STYLE_CD\n"
        "PRICE_CX3\n"
        "BARCODE\n"
        "STYLE_DESCRIPTION \n"
        "TOP_STYLE_RANK\n"
        "From column AF onwards: Style categories (mandatory)\n\n"
        "The script will read data from these columns to generate the report."
    )
    Label(window, text=info_text, justify="left").pack(pady=20, padx=10)

    Button(window, text="OK", command=window.destroy).pack(pady=10)

    window.wait_window()

# === TREATED DIAMONDS INFO DIALOG (BARCODE) ===
def show_treated_diamonds_info_dialog_barcode():
    window = Toplevel()
    window.title("Treated Diamond File Format")

    info_text = (
        "Please select the Excel file containing Treated Diamond Style IDs.\n\n"
        "This file should have only ONE column containing the Style CDs, these are those items which will have 'Treated Diamond' written below the Style Image."
    )
    Label(window, text=info_text, justify="left").pack(pady=20, padx=10)

    Button(window, text="OK", command=window.destroy).pack(pady=10)

    window.wait_window()

# === GUI PROGRESS BAR ===
def show_progress_bar(total_steps):
    window = Toplevel()
    window.title("Generating PDFs...")
    window.geometry("400x100")
    Label(window, text="Generating ASHI Barcode PDFs...").pack(pady=10)
    bar = Progressbar(window, length=350, mode='determinate', maximum=total_steps)
    bar.pack(pady=5)
    return window, bar

# === HEADER ===
def draw_header(c, title):
    header_width = PAGE_WIDTH-50
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


# === FOOTER ===
def draw_footer(c, page_num, total_pages):
    today = datetime.today().strftime("%m/%d/%Y")
    c.setFont("Helvetica", 8)
    c.setFillColor(HexColor("#000000"))
    bottom_margin = 13
    c.drawString(40, bottom_margin, f"{today}")
    c.drawRightString(PAGE_WIDTH - 30, bottom_margin, f"Page  {page_num} of {total_pages}")

# === IMAGE LOOKUP ===
def find_image_path(style_code):
    for ext in SUPPORTED_EXTENSIONS:
        full_path = os.path.join(IMAGE_FOLDER, f"{style_code}{ext}")
        if os.path.exists(full_path):
            return full_path
    return None

# === BARCODE PDF GENERATOR ===
def generate_pdf(data, output_file, title, progress_bar=None, treated_styles=None):
    c = canvas.Canvas(output_file, pagesize=letter)

    BOX_WIDTH = 142  # Updated grid width
    BOX_HEIGHT = 170  # Updated grid height
    H_SPACING = 0
    V_SPACING = 0
    IMG_HEIGHT = int(BOX_HEIGHT * 0.60)
    IMG_WIDTH = 98

    total_grid_width = (BOX_WIDTH * STYLES_PER_ROW) + (H_SPACING * (STYLES_PER_ROW - 1))
    LEFT_MARGIN = (PAGE_WIDTH - total_grid_width) / 2

    total_pages = (len(data) - 1) // STYLES_PER_PAGE + 1
    style_count = 0

    for page in range(total_pages):
        draw_header(c,title)
        x = LEFT_MARGIN
        y = PAGE_HEIGHT - 88  # Adjusted grid start after header

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

                        if style_code in treated_styles:
                            watermark_text = "TREATED DM"
                            c.setFont("Helvetica-Bold", 6)
                            c.setFillColorRGB(0, 0, 0)
                            c.drawString(draw_x - 15, draw_y, watermark_text)
                except Exception as e:
                    print(f"Image error: {e}")

            separator_y = y - IMG_HEIGHT - 8
            c.setLineWidth(0.5)
            c.line(x - 0.5, separator_y, x + BOX_WIDTH - 1, separator_y)

            text_y = separator_y - 8

            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(center_x, text_y, style_code.upper())

            style_description = row.get('STYLE_DESCRIPTION', '')
            if pd.notna(style_description) and str(style_description).strip() != '':
                full_desc = str(style_description).strip()
                if len(full_desc) > 20:
                    words = full_desc.split()
                    first_line = ""
                    second_line = ""
                    for word in words:
                        if len(first_line + " " + word) <= 20:
                            if first_line:
                                first_line += " "
                            first_line += word
                        else:
                            if second_line:
                                second_line += " "
                            second_line += word

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
            except Exception as e:
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

# (MAIN function remains same)
# === MAIN MULTI-PDF GENERATOR ===
def main():
    root = Tk()
    root.withdraw()
    show_initial_info_dialog()

    print("Please select the main data file...")
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
    if not excel_path:
        print("No main data file selected. Exiting.")
        return

    show_treated_diamonds_info_dialog_barcode()
    print("Please select the treated diamond style IDs file...")
    treated_diamonds_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
    if not treated_diamonds_file:
        print("No treated diamond style IDs file selected. Exiting.")
        return

    df = pd.read_excel(excel_path, sheet_name=0)
    treated_df = pd.read_excel(treated_diamonds_file, sheet_name=0)
    treated_styles = treated_df['STYLE_CD'].tolist()

    # Ask user for output directory
    output_base_dir = filedialog.askdirectory(title="Select Output Folder for Barcode PDFs")
    if not output_base_dir:
        print("No output destination selected. Exiting.")
        return

    # Create the output directory
    output_dir = os.path.join(output_base_dir, "generated_pdfs_barcode")
    os.makedirs(output_dir, exist_ok=True)

    start_index = 31
    non_empty_columns = df.iloc[:, start_index:].dropna(how='all', axis=1).columns
    selected_columns = df.columns[start_index : df.columns.get_loc(non_empty_columns[-1]) + 1]

    progress_window, progress_bar = show_progress_bar(total_steps=len(selected_columns))

    generated_count = 0
    skipped_indices = []

    for idx, col in enumerate(selected_columns):
        title_for_pdf = col.split(")", 1)[-1].strip() if ")" in col else col.strip()
        required_columns = ['STYLE_CD', 'BARCODE', 'PRICE_CX3', 'STYLE_DESCRIPTION', col, 'TOP_STYLE_RANK']
        for rc in required_columns:
            if rc not in df.columns:
                messagebox.showerror("Missing Column", f"Excel must contain '{rc}' column.")
                return

        temp_df = df[required_columns].copy()
        temp_df = temp_df.dropna(subset=[col, 'BARCODE', 'TOP_STYLE_RANK'])
        top_data = temp_df.sort_values(by='TOP_STYLE_RANK', ascending=True).reset_index(drop=True)

        if top_data.empty:
            skipped_indices.append(idx + 1)
            continue

        safe_filename = "".join(c if c.isalnum() or c in (' ', '.', '-', '_') else "_" for c in col)

        output_csv_path = os.path.join(output_dir, f"{safe_filename}_order.csv")
        top_data.to_csv(output_csv_path, index=False)

        output_pdf_path = os.path.join(output_dir, f"{safe_filename}_barcode.pdf")
        generate_pdf(top_data, output_pdf_path, title_for_pdf, progress_bar, treated_styles)

        progress_bar['value'] = idx + 1
        progress_bar.update()
        generated_count += 1

    progress_window.destroy()

    skipped_message = (
        f"✅ PDFs Generated: {generated_count}\n"
        f"❌ PDFs Skipped: {len(skipped_indices)}"
    )
    if skipped_indices:
        skipped_message += f"\n\nSkipped Column Numbers:\n{', '.join(map(str, skipped_indices))}"

    messagebox.showinfo("Summary", skipped_message)

if __name__ == "__main__":
    main()