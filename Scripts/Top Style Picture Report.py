import pandas as pd
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
from tkinter import Tk, filedialog, Toplevel, Label, messagebox, Button
from tkinter.ttk import Progressbar
from datetime import datetime
import os

# === PAGE SETTINGS ===
PAGE_WIDTH, PAGE_HEIGHT = LETTER
ROWS_PER_PAGE = 4
STYLES_PER_ROW = 4
STYLES_PER_PAGE = ROWS_PER_PAGE * STYLES_PER_ROW
IMAGE_FOLDER = r"D:\Ashi_Images"
SUPPORTED_EXTENSIONS = ['.jpg', '.jpeg', '.png']

# === HEADER IMAGE PATH ===
HEADER_IMAGE_PATH = r"C:\Users\vaishalisinghbaghel\Desktop\FINAL_PP_AUTOMATION\ASHI-Logo-Header.jpg"

# === INITIAL INFO DIALOG ===
def show_initial_info_dialog():
    window = Toplevel()
    window.title("Input Excel File Format")
    window.geometry("550x250")
    window.resizable(False, False)

    info_text = (
        "Please ensure your input Excel file contains the following columns:\n\n"
        "STYLE_CD\n"
        "STYLE_DESCRIPTION\n"
        "PRICE_CX3\n"
        "TOP_STYLE_RANK\n"
        "From column AF onwards: Style categories (mandatory)\n\n"
        "The script will read data from these columns to generate the report."
    )
    Label(window, text=info_text, justify="left").pack(pady=20, padx=10)

    Button(window, text="OK", command=window.destroy).pack(pady=10)

    window.wait_window()

# === TREATED DIAMONDS INFO DIALOG ===
def show_treated_diamonds_info_dialog():
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
    Label(window, text="Generating ASHI Styles PDFs...").pack(pady=10)
    bar = Progressbar(window, length=350, mode='determinate', maximum=total_steps)
    bar.pack(pady=5)
    return window, bar

# === HEADER ===
def draw_header(c, title):
    header_width = PAGE_WIDTH + 10  # HEADER WIDTH: Reduced by 3 points
    header_height = 60             # HEADER HEIGHT
    header_x = 2                   # HEADER X: Shifted 2 points to the right
    header_y = PAGE_HEIGHT - header_height

    try:
        c.drawImage(HEADER_IMAGE_PATH, header_x, header_y, width=header_width -5 , height=header_height-10, preserveAspectRatio=True, mask='auto')
    except:
        pass

    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(HexColor("#1F4E79"))
    c.drawCentredString(PAGE_WIDTH / 2, header_y - 25, title.upper())

# === FOOTER ===
def draw_footer(c, page_num, total_pages):
    today = datetime.today().strftime("%m/%d/%Y")
    c.setFont("Helvetica", 8)
    c.setFillColor(HexColor("#000000"))

    bottom_margin = 15  # FOOTER MARGIN FROM BOTTOM
    c.drawString(40, bottom_margin, f"{today}")
    c.drawRightString(PAGE_WIDTH - 40, bottom_margin, f"Page {page_num} of {total_pages}")

# === IMAGE LOOKUP ===
def find_image_path(style_code):
    for ext in SUPPORTED_EXTENSIONS:
        full_path = os.path.join(IMAGE_FOLDER, f"{style_code}{ext}")
        if os.path.exists(full_path):
            return full_path
    return None

# === PDF GENERATOR ===
def generate_pdf(data, output_file, title, progress_bar=None, treated_styles=None):
    c = canvas.Canvas(output_file, pagesize=LETTER)

    BOX_WIDTH = 142             # GRID WIDTH
    BOX_HEIGHT = 165            # GRID HEIGHT: Reduced box height to 160 points
    H_SPACING = 0
    V_SPACING = 0
    IMG_HEIGHT = int(BOX_HEIGHT * 0.70)
    IMG_WIDTH = 100

    total_grid_width = (BOX_WIDTH * STYLES_PER_ROW) + (H_SPACING * (STYLES_PER_ROW - 1))
    LEFT_MARGIN = (PAGE_WIDTH - total_grid_width) / 2

    total_pages = (len(data) - 1) // STYLES_PER_PAGE + 1
    style_count = 0

    for page in range(total_pages):
        draw_header(c, title)

        y = PAGE_HEIGHT - 100  # Starting y position for grid after header and title
        x = LEFT_MARGIN

        start = page * STYLES_PER_PAGE
        end = start + STYLES_PER_PAGE
        page_data = data[start:end]

        for i, row in page_data.iterrows():
            style_code = row['STYLE_CD']
            price = row['PRICE_CX3']
            image_path = find_image_path(style_code)

            c.setStrokeColor(HexColor("#B0B0B0"))
            c.setLineWidth(1)
            c.rect(x, y - BOX_HEIGHT, BOX_WIDTH, BOX_HEIGHT, fill=0)

            center_x = x + BOX_WIDTH / 2
            image_x = x + (BOX_WIDTH - IMG_WIDTH) / 2
            image_y = y - IMG_HEIGHT - 5

            if image_path:
                try:
                    c.drawImage(ImageReader(image_path), image_x, image_y, width=IMG_WIDTH, height=IMG_HEIGHT, preserveAspectRatio=True)
                except:
                    pass

            separator_y = y - IMG_HEIGHT - 8
            c.setLineWidth(0.5)
            c.line(x - 0.5, separator_y, x + BOX_WIDTH - 1, separator_y)

            text_y = separator_y - 8
            c.setFont("Helvetica-Bold", 8)
            c.setFillColor(HexColor("#000000"))
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

            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(center_x, text_y - 30, f"${int(price):,}")

            if treated_styles and style_code in treated_styles:
                watermark_text = "TREATED DM"
                c.setFont("Helvetica-Bold", 6)
                wm_x = image_x - 15
                wm_y = image_y + 0
                c.setFillColor(HexColor("#000000"))
                c.drawString(wm_x, wm_y, watermark_text)

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

# === MAIN MULTI-PDF GENERATOR ===
def main():
    root = Tk()
    root.withdraw()
    show_initial_info_dialog()
    print(" Please select your Excel file with styles...")
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
    if not excel_path:
        print(" No file selected. Exiting.")
        return

    show_treated_diamonds_info_dialog()
    print(" Please select the file with treated diamond style IDs...")
    treated_diamonds_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
    if not treated_diamonds_file:
        print("No treated diamond style IDs file selected. Exiting.")
        return

    # Read data from Excel files
    try:
        df = pd.read_excel(excel_path, sheet_name=0)
        treated_df = pd.read_excel(treated_diamonds_file, sheet_name=0)
        treated_styles = treated_df['STYLE_CD'].tolist()
    except Exception as e:
        messagebox.showerror("File Reading Error", f"Error reading Excel files: {e}")
        return

    # Ask user for output directory
    output_base_dir = filedialog.askdirectory(title="Select Output Folder for PDFs")
    if not output_base_dir:
        print("No output destination selected. Exiting.")
        return

    # Create the output directory
    output_dir = os.path.join(output_base_dir, "generated_pdfs")
    os.makedirs(output_dir, exist_ok=True)

    # Determine columns to process (assuming relevant data starts from a certain index)
    start_index = 31 # Based on previous script's logic
    # Find the last non-empty column after the start_index
    if start_index >= len(df.columns):
        messagebox.showerror("Column Error", "Start index for columns is out of bounds.")
        return

    non_empty_cols_after_start = df.iloc[:, start_index:].dropna(how='all', axis=1).columns

    if non_empty_cols_after_start.empty:
        messagebox.showerror("Column Error", "No columns with data found after the start index.")
        return

    last_col_index = df.columns.get_loc(non_empty_cols_after_start[-1])
    selected_columns = df.columns[start_index : last_col_index + 1]

    # Show progress bar using the existing hidden root window
    progress_window, progress_bar = show_progress_bar(total_steps=len(selected_columns))

    for idx, col in enumerate(selected_columns):
        title_for_pdf = col.split(")", 1)[-1].strip() if ")" in col else col.strip()

        required_columns = ['STYLE_CD', 'PRICE_CX3', 'STYLE_DESCRIPTION', col, 'TOP_STYLE_RANK']
        for rc in required_columns:
            if rc not in df.columns:
                messagebox.showerror("Missing Column", f"Excel must contain '{rc}' column.")
                return

        temp_df = df[required_columns].copy()
        temp_df = temp_df.dropna(subset=[col, 'TOP_STYLE_RANK'])
        top_data = temp_df.sort_values(by='TOP_STYLE_RANK', ascending=True).reset_index(drop=True)

        if top_data.empty:
            continue

        safe_filename = "".join(c if c.isalnum() or c in (' ', '.', '-', '_') else "_" for c in col)

        output_csv_path = os.path.join(output_dir, f"{safe_filename}_order.csv")
        top_data.to_csv(output_csv_path, index=False)

        output_pdf_path = os.path.join(output_dir, f"{safe_filename}.pdf")
        generate_pdf(top_data, output_pdf_path, title_for_pdf, progress_bar, treated_styles)
        progress_bar['value'] = idx + 1
        progress_bar.update()

    progress_window.destroy()
    messagebox.showinfo("Success", " All PDFs Generated Successfully!")

if __name__ == "__main__":
    main()