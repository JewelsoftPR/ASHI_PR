import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.graphics.barcode import code128
from tkinter import Tk, filedialog, Toplevel, Label, messagebox, Frame, Button
from tkinter.ttk import Progressbar
from datetime import datetime
import os

# === PAGE SETTINGS ===
PAGE_WIDTH, PAGE_HEIGHT = letter
ROWS_PER_PAGE = 9
STYLES_PER_ROW = 4
STYLES_PER_PAGE = ROWS_PER_PAGE * STYLES_PER_ROW
MARGIN_X = 30  # Left/right margin

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
    c.drawRightString(PAGE_WIDTH - 30, PAGE_HEIGHT - 45, title)

# === PDF GENERATOR ===
def generate_pdf(data, output_file, title, progress_bar=None, step=0):
    c = canvas.Canvas(output_file, pagesize=letter)

    usable_width = PAGE_WIDTH - 2 * MARGIN_X
    BOX_WIDTH = usable_width / STYLES_PER_ROW
    BOX_HEIGHT = (PAGE_HEIGHT - 150) / ROWS_PER_PAGE  # space reserved for header

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

            x = MARGIN_X + col * BOX_WIDTH
            y = PAGE_HEIGHT - 70 - row_pos * BOX_HEIGHT  # header offset = 70

            barcode_value = str(row['barcode']).strip('*')
            description = str(row['name']).strip()
            item_id = str(row['id']).strip() if 'id' in row else ''

            c.setStrokeColor(HexColor("#B0B0B0"))
            c.setLineWidth(1)
            c.rect(x, y - BOX_HEIGHT, BOX_WIDTH, BOX_HEIGHT, fill=0)

            center_x = x + BOX_WIDTH / 2
            current_y = y - 10

            # === ID ===
            if item_id:
                c.setFont("Helvetica-Bold", 6.5)
                c.setFillColor(HexColor("#1F4E79"))
                c.drawCentredString(center_x, current_y, item_id.upper())

            # === DESCRIPTION ===
            current_y -= 10
            lines = []
            if "\n" in description:
                lines = description.split("\n")
            else:
                words = description.split()
                line = ""
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

            # === BARCODE ===
            try:
              box_bottom_y = y - BOX_HEIGHT
          
              # Draw barcode a bit above the bottom
              barcode_y = box_bottom_y + 18
              barcode = code128.Code128(barcode_value, barHeight=15, barWidth=0.75)
              c.setFillColor(HexColor("#000000"))
              barcode.drawOn(c, x + (BOX_WIDTH - barcode.width) / 2, barcode_y)
          
              # Draw spaced barcode text just above it
              spaced_text = '    '.join(barcode_value)
              c.setFont("Helvetica", 5.5)
              c.drawCentredString(center_x, barcode_y - 5, spaced_text)
            except:
              c.setFont("Helvetica", 6)
              c.drawCentredString(center_x, box_bottom_y + 13, "Invalid barcode")


        c.showPage()

    c.save()
    if progress_bar:
        progress_bar['value'] = step + 1
        progress_bar.update()

# === MAIN ===
def main():
    root = Tk()
    root.withdraw()
    
    # Create dialog box
    dialog = Toplevel(root)
    dialog.title("Important Information")
    dialog.geometry("600x200")
    
    # Create frame for better organization
    frame = Frame(dialog, padx=20, pady=20)
    frame.pack(expand=True, fill='both')
    
    # Add message
    Label(frame, text="This Script Requires The Following Columns In The Input Sheet To Generate The Report:", 
          wraplength=550, justify='left').pack(anchor='w', pady=(0, 10))
    
    Label(frame, text="ID, Name and Barcode", 
          wraplength=550, justify='left').pack(anchor='w', pady=(0, 10))
    
    # Add note in red color
    Label(frame, text="NOTE: Make Sure To Sort The Data (A-Z or Smallest to largest ID) Before Proceeding To Generate The Report.", 
          wraplength=550, justify='left', fg='red').pack(anchor='w', pady=(0, 20))
    
    # Add OK button
    Button(frame, text="OK", command=dialog.destroy, width=10).pack()
    
    # Wait for dialog to close
    root.wait_window(dialog)
    
    print("Select the Excel file...")
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
    if not excel_path:
        print("No file selected. Exiting.")
        return

    df = pd.read_excel(excel_path)
    if not all(col in df.columns for col in ['id', 'name', 'barcode']):
        messagebox.showerror("Missing Columns", "Excel must contain 'id', 'name', and 'barcode' columns.")
        return

    # Ask user where to save the PDFs
    print("Select where to save the PDFs...")
    output_dir = filedialog.askdirectory(title="Select Folder to Save PDFs")
    if not output_dir:
        print("No save location selected. Exiting.")
        return

    # Create Display_Tray_PDFs folder in selected location
    output_dir = os.path.join(output_dir, "Display_Tray_PDFs")
    os.makedirs(output_dir, exist_ok=True)

    root = Tk()
    root.withdraw()
    progress_window, progress_bar = show_progress_bar(total_steps=1)

    output_pdf_path = os.path.join(output_dir, "Display Tray.pdf")
    generate_pdf(df, output_pdf_path, title="Display Tray", progress_bar=progress_bar, step=0)

    progress_window.destroy()
    messagebox.showinfo("Done", f"✅ PDF generated with {len(df)} entries!\nSaved in: {output_dir}")

if __name__ == "__main__":
    main()
