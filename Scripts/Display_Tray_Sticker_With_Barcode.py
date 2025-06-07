import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.graphics.barcode import code128
from tkinter import Tk, filedialog, Toplevel, Label, messagebox, Button
from tkinter.ttk import Progressbar
from datetime import datetime
import os

# === PAGE SETTINGS ===
PAGE_WIDTH, PAGE_HEIGHT = A4
ROWS_PER_PAGE = 10
STYLES_PER_ROW = 3
STYLES_PER_PAGE = ROWS_PER_PAGE * STYLES_PER_ROW
MARGIN_X = 30
MARGIN_Y_TOP = 100
MARGIN_Y_BOTTOM = 50

# === INITIAL INFO DIALOG (STICKER) ===
def show_initial_info_dialog_sticker():
    window = Toplevel()
    window.title("Input File Format Information")
    window.geometry("600x250") # Set desired size
    window.resizable(False, False) # Make non-resizable

    info_text_part1 = (
        "Please ensure your input Excel file contains the following columns:\n\n"
        "column A: id\n"
        "column B: name\n"
        "column C: barcode\n\n"
        "The script will read data from these columns to generate the report."
    )
    info_text_part2 = "NOTE: Make sure to sort the data (A-Z or Smallest to largest ID) before proceeding to generate the report *"

    Label(window, text=info_text_part1, justify="left").pack(pady=(20, 0), padx=10)
    Label(window, text=info_text_part2, justify="left", fg="red").pack(pady=(0, 20), padx=10)

    Button(window, text="OK", command=window.destroy).pack(pady=10)

    window.wait_window()

# === GUI PROGRESS BAR ===
def show_progress_bar(total_steps):
    window = Toplevel()
    window.title("Generating PDFs...")
    window.geometry("400x100")
    Label(window, text="Generating All Barcode Entries...").pack(pady=10)
    bar = Progressbar(window, length=350, mode='determinate', maximum=total_steps)
    bar.pack(pady=5)
    return window, bar

# === HEADER ===
def draw_header(c, title, page_num, total_pages):
    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(HexColor("#000000"))
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 25, "A S H I   D I A M O N D S ,  LLC")

    c.setFont("Helvetica", 7)
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 35, "18 EAST 48TH STREET ~ 14TH FLOOR ~  NY, NY 10017")
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 45, "TEL : (212) 319-8291 ~  FAX : (212) 319-4341")
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 55, "TEL : (800) 622-2744 ~  FAX : (800) 275-2744")

    today = datetime.today().strftime("%m/%d/%Y")
    c.setFont("Helvetica", 7)
    c.drawRightString(PAGE_WIDTH - 30, PAGE_HEIGHT - 25, f"Date :  {today}")
    c.drawRightString(PAGE_WIDTH - 30, PAGE_HEIGHT - 35, f"Page :  {page_num} of {total_pages}")

    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(HexColor("#1F4E79"))
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 80, title)

# === PDF GENERATOR ===
def generate_pdf(data, output_file, title, progress_bar=None, step=0):
    c = canvas.Canvas(output_file, pagesize=A4)

    # --- Layout constants ---
    aspect_ratio = 2.59  # from reference image (width / height)
    usable_width = PAGE_WIDTH - 2 * MARGIN_X
    box_spacing = 10  # spacing between boxes horizontally

    BOX_WIDTH = (usable_width - (STYLES_PER_ROW - 1) * box_spacing) / STYLES_PER_ROW
    BOX_HEIGHT = BOX_WIDTH / aspect_ratio

    start_y = PAGE_HEIGHT - MARGIN_Y_TOP
    total_pages = (len(data) - 1) // STYLES_PER_PAGE + 1

    for page in range(total_pages):
        draw_header(c, title, page + 1, total_pages)
        start = page * STYLES_PER_PAGE
        end = start + STYLES_PER_PAGE
        page_data = data[start:end]

        for idx, row in page_data.iterrows():
            i = idx - start
            col = i % STYLES_PER_ROW
            row_pos = i // STYLES_PER_ROW

            x = MARGIN_X + col * (BOX_WIDTH + box_spacing)
            y = start_y - row_pos * BOX_HEIGHT  # Removed vertical spacing between rows

            barcode_value = str(row['barcode']).strip('*')
            description = str(row['name']).strip()
            item_id = str(row['id']).strip() if 'id' in row else ''

            # === Grid Box ===
            c.setStrokeColor(HexColor("#B0B0B0"))
            c.setLineWidth(1)
            c.roundRect(x, y - BOX_HEIGHT, BOX_WIDTH, BOX_HEIGHT, 4, fill=0)

            center_x = x + BOX_WIDTH / 2
            current_y = y - 10

            # === ID ===
            if item_id:
                c.setFont("Helvetica-Bold", 6.5)
                c.setFillColor(HexColor("#1F4E79"))
                c.drawCentredString(center_x, current_y, item_id.upper())

            # === Description ===
            current_y -= 10
            words = description.split()
            lines, line = [], ""
            for word in words:
                if len(line + " " + word) <= 28:
                    line += (" " if line else "") + word
                else:
                    lines.append(line)
                    line = word
            if line:
                lines.append(line)

            c.setFont("Helvetica", 6)
            for line in lines[:2]:
                c.drawCentredString(center_x, current_y, line.strip())
                current_y -= 9

            # === Barcode ===
            try:
                barcode_y = y - BOX_HEIGHT + 18
                barcode = code128.Code128(barcode_value, barHeight=15, barWidth=0.75)
                c.setFillColor(HexColor("#000000"))
                barcode.drawOn(c, x + (BOX_WIDTH - barcode.width) / 2, barcode_y)

                spaced_text = '    '.join(barcode_value)
                c.setFont("Helvetica", 5.5)
                c.drawCentredString(center_x, barcode_y - 5, spaced_text)
            except:
                c.setFont("Helvetica", 6)
                c.drawCentredString(center_x, y - BOX_HEIGHT + 13, "Invalid barcode")

        c.showPage()

    c.save()
    if progress_bar:
        progress_bar['value'] = step + 1
        progress_bar.update()

# === MAIN ===
def main():
    print("Select the Excel file...")
    root = Tk()
    root.withdraw()
    show_initial_info_dialog_sticker()
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
    if not excel_path:
        print("No file selected. Exiting.")
        return

    # Get the Excel file name without extension
    excel_filename = os.path.splitext(os.path.basename(excel_path))[0]

    df = pd.read_excel(excel_path)
    if not all(col in df.columns for col in ['id', 'name', 'barcode']):
        messagebox.showerror("Missing Columns", "Excel must contain 'id', 'name', and 'barcode' columns.")
        return

    # Ask user for output directory
    output_base_dir = filedialog.askdirectory(title="Select Output Folder for Barcode Stickers")
    if not output_base_dir:
        print("No output destination selected. Exiting.")
        return

    # Create the output directory
    output_dir = os.path.join(output_base_dir, "Display Program Style - Sticker")
    os.makedirs(output_dir, exist_ok=True)

    root = Tk()
    root.withdraw()
    progress_window, progress_bar = show_progress_bar(total_steps=1)

    # Use Excel filename for the PDF
    output_pdf_path = os.path.join(output_dir, f"{excel_filename}.pdf")
    generate_pdf(df, output_pdf_path, title=excel_filename, progress_bar=progress_bar, step=0)

    progress_window.destroy()
    messagebox.showinfo("Done", f"✅ PDF generated as '{excel_filename}.pdf' with {len(df)} entries!")

if __name__ == "__main__":
    main()