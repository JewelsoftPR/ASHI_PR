import os
import pyodbc
import pandas as pd
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
from PIL import Image
import textwrap
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Toplevel
from tkcalendar import Calendar
from threading import Thread

# === PAGE SETTINGS ===
PAGE_WIDTH, PAGE_HEIGHT = LETTER
STYLES_PER_ROW = 4
STYLES_PER_PAGE = 16
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

# === SQL CONNECTION SETTINGS ===
server = '192.168.1.11'
database = 'jewel_ashi_prod'
username = 'sa'
password = 'avalondb@no1Knows'

selected_date = None
selected_start_date = None
selected_end_date = None

def map_network_drive():
    try:
        subprocess.run(r'net use Z: \\192.168.1.83\jewelsoft\Ashi_Images', shell=True, capture_output=True)
    except Exception as e:
        print(f"Failed to map drive: {e}")

def draw_header_text(c, header):
    today = datetime.today().strftime("%m/%d/%Y")
    c.setFont("Helvetica-Bold", 7)
    c.setFillColor(HexColor("#000000"))

    header_image_x = 40
    header_image_height = 51
    header_image_y_bottom = PAGE_HEIGHT - 20

    try:
        img_width_orig, img_height_orig = Image.open(HEADER_IMAGE_PATH).size
        aspect_ratio = img_width_orig / img_height_orig
        header_image_width = header_image_height * aspect_ratio
        c.drawImage(HEADER_IMAGE_PATH, header_image_x, header_image_y_bottom - header_image_height, 
                    width=header_image_width, height=header_image_height, 
                    preserveAspectRatio=True, mask='auto')
    except Exception as e:
        print(f"Error drawing header image: {e}")

def draw_footer(c, page_num, total_pages):
    timestamp = datetime.now().strftime("%m/%d/%Y - %I:%M %p")
    c.setFont("Helvetica", 5)
    c.setFillColor(HexColor("#000000"))
    c.drawString(40, 10, timestamp)
    c.drawRightString(PAGE_WIDTH - 40, 10, f"PAGE {page_num} OF {total_pages}")

def generate_pdf(data, output_path, header):
    c = canvas.Canvas(output_path, pagesize=LETTER)
    total_pages = (len(data) - 1) // STYLES_PER_PAGE + 1
    LEFT_MARGIN = (PAGE_WIDTH - (BOX_WIDTH * STYLES_PER_ROW)) / 2
    TOP_MARGIN = PAGE_HEIGHT - 130

    for page in range(total_pages):
        draw_header_text(c, header)
        start = page * STYLES_PER_PAGE
        end = min(start + STYLES_PER_PAGE, len(data))
        page_data = data.iloc[start:end].reset_index(drop=True)

        today = header.get("trans_dt", datetime.today().strftime("%m/%d/%Y"))
        info_y = TOP_MARGIN + 40
        left_x = 50
        right_x = PAGE_WIDTH / 2 + 100
        lines = [
            ("TO:", header.get("account_id", ""), "CONTACT:", header.get("contact1", "")),
            ("FAX:", header.get("fax1", ""), "CONTACT NO:", header.get("phone1", "")),
            ("QUOTE:", header.get("quotationcode", ""), "TRANS DATE:", today)
        ]
        for label1, val1, label2, val2 in lines:
            c.drawString(left_x, info_y, label1)
            display_val1 = "" if val1 is None or str(val1).strip().lower() in ['', 'none'] else str(val1)
            c.drawString(left_x + 60, info_y, display_val1)
            if label2:
                c.drawString(right_x, info_y, label2)
                c.drawString(right_x + 60, info_y, str(val2))
            info_y -= 14
        c.setStrokeColor(HexColor("#B0B0B0"))
        c.rect(32, TOP_MARGIN + 40 - (3 * 14) + 8, PAGE_WIDTH - 65, 50)

        for idx, row in page_data.iterrows():
            item_id = str(row.get('item_id', '')).strip()
            item_desc = str(row.get('item_description', '')).strip()
            item_price = row.get('item_price', 0)
            trans_qty = row.get('trans_qty', 0)
            filename = str(row.get('style_photo', '')).strip()

            image_path = os.path.join('Z:\\', filename)
            if not os.path.exists(image_path):
                image_path = f"\\\\192.168.1.83\\jewelsoft\\Ashi_Images\\{filename}"

            row_num, col_num = divmod(idx, STYLES_PER_ROW)
            current_x = LEFT_MARGIN + col_num * BOX_WIDTH
            current_y = TOP_MARGIN - row_num * BOX_HEIGHT

            c.setStrokeColor(HexColor("#B0B0B0"))
            c.rect(current_x, current_y - BOX_HEIGHT, BOX_WIDTH, BOX_HEIGHT, fill=0)

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
                    draw_x = current_x + (BOX_WIDTH - display_width) / 2
                    draw_y = current_y - display_height - 5
                    c.drawImage(ImageReader(image_path), draw_x, draw_y, width=display_width, height=display_height)
                except Exception as e:
                    print(f"Image load error: {e} - {image_path}")

            sep_y = current_y - IMG_HEIGHT - 8
            c.line(current_x, sep_y, current_x + BOX_WIDTH, sep_y)  # <== MODIFIED LINE

            text_base_y = current_y - IMG_HEIGHT - 5
            text_x = current_x + TEXT_POS['item_cd']['x_offset']
            text_y = text_base_y - TEXT_POS['item_cd']['y_offset']
            c.setFont("Helvetica-Bold", 7)
            c.drawString(text_x, text_y, f"{item_id}    Price : ${item_price:,.2f}")

            c.setFont("Helvetica", 6.5)
            wrapped_lines = textwrap.wrap(item_desc, width=32)
            desc_y = text_base_y - TEXT_POS['desc']['y_start_offset']
            for line in wrapped_lines[:3]:
                c.drawString(current_x + TEXT_POS['desc']['x_offset'], desc_y, line)
                desc_y -= TEXT_POS['desc']['line_spacing']

            c.setFont("Helvetica-Bold", 6)
            c.drawString(current_x + TEXT_POS['qty']['x_offset'],
                         current_y - BOX_HEIGHT + TEXT_POS['qty']['bottom_margin'],
                         f"QUOT QTY: {trans_qty}")

        draw_footer(c, page + 1, total_pages)
        c.showPage()

    c.save()

def run_gui():
    global selected_date
    stop_flag = [False]

    def generate_report():
        # Get values from the two new entry fields
        start_quote_input = start_quote_entry.get().strip()
        end_quote_input = end_quote_entry.get().strip()
        # Get selected date range values
        start_date_input = selected_start_date
        end_date_input = selected_end_date

        # --- Input Validation ---
        # Require either quotation input OR date input (single or range)
        if not (start_quote_input or end_quote_input) and not (start_date_input):
             messagebox.showerror("Missing Input", "Please enter a Quotation Number Range OR select a Transaction Date or Date Range.")
             return

        # Validate quotation numbers if provided (existing logic remains)
        if start_quote_input or end_quote_input:
            if start_quote_input and not start_quote_input.isdigit():
                 messagebox.showerror("Invalid Input", "Invalid Starting Quotation Number format. Please enter a number.")
                 return
            
            if end_quote_input and not end_quote_input.isdigit():
                 messagebox.showerror("Invalid Input", "Invalid Ending Quotation Number format. Please enter a number.")
                 return
            
            if start_quote_input and end_quote_input and int(start_quote_input) > int(end_quote_input):
                messagebox.showerror("Invalid Input", "Starting Quotation Number cannot be greater than Ending Quotation Number.")
                return

        # Validate date range logic if at least a start date is provided
        if start_date_input:
            # If only start date is provided, no need for end date validation
            # If both are provided, validate the range
            if end_date_input:
                 try:
                    # Attempt to parse as datetime objects for robust comparison
                    start_dt = datetime.strptime(start_date_input, '%m/%d/%Y')
                    end_dt = datetime.strptime(end_date_input, '%m/%d/%Y')
                    if start_dt > end_dt:
                         messagebox.showerror("Invalid Input", "Start Date cannot be after End Date.")
                         return
                 except ValueError:
                      # Should not happen if date_pattern is consistent, but as a safeguard
                      messagebox.showerror("Internal Error", "Could not parse selected date(s).")
                      return


        # --- Output Directory Selection ---
        output_dir = filedialog.askdirectory(title="Select Folder to Save PDFs")
        if not output_dir:
            messagebox.showwarning("Cancelled", "Folder selection cancelled.")
            return

        # Create a timestamped subfolder within the selected directory
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_subfolder = os.path.join(output_dir, f"Quotation_Reports")

        try:
            os.makedirs(output_subfolder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create output folder:\n{e}")
            return

        # --- Build WHERE clauses and parameters based on input ---
        where_clauses = ["a.trans_bk = 'SQ01'"]
        params = []

        # Handle date filtering (single date or range)
        if start_date_input:
             if end_date_input:
                  # Both start and end dates provided: filter by date range
                  where_clauses.append("a.trans_dt BETWEEN ? AND ?")
                  params.append(start_date_input)
                  params.append(end_date_input)
             else:
                  # Only start date provided: filter by single date
                  where_clauses.append("a.trans_dt = ?")
                  params.append(start_date_input)

        # Handle quotation number filtering (existing logic remains)
        if start_quote_input:
            if end_quote_input:
                # Both start and end provided: filter by range
                where_clauses.append("a.trans_no BETWEEN ? AND ?")
                params.append(start_quote_input)
                params.append(end_quote_input)
            else:
                # Only start provided: filter by single quotation number
                where_clauses.append("a.trans_no = ?")
                params.append(start_quote_input)

        where_condition = " AND ".join(where_clauses)

        query = f"""
        SELECT 
            a.trans_no,
            a.account_id,
            c.contact1,
            c.contact2,
            c.phone1,
            c.fax1,
            b.serial_no,
            b.item_id,
            d.style_photo,
            b.item_description,
            b.item_price,
            b.trans_qty,
            a.trans_dt
        FROM saoitrinvhd a
        LEFT JOIN saoitrinvdtl b
            ON a.trans_bk = b.trans_bk AND a.trans_no = b.trans_no AND a.trans_dt = b.trans_dt
        LEFT JOIN faarmscustomer c
            ON a.account_id = c.id
        LEFT JOIN stylmsstylehd d
            ON b.item_id = d.style_cd
        WHERE {where_condition}
        ORDER BY a.trans_no, b.serial_no
        """

        btn_generate.config(state='disabled')
        status_label.config(text="Starting...")
        progress_bar['value'] = 0
        text_output.delete('1.0', tk.END)

        def process():
            try:
                map_network_drive()
                conn = pyodbc.connect(
                    f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
                )
                # Pass parameters to pd.read_sql
                df = pd.read_sql(query, conn, params=params)
                conn.close()

                if df.empty:
                    messagebox.showinfo("No Results", "No data found for the selected criteria.")
                    status_label.config(text="No records found.")
                    return

                unique_trans = df['trans_no'].nunique()
                progress_bar.config(maximum=unique_trans, mode='determinate')
                done = 0

                for trans_no, group in df.groupby("trans_no"):
                    if stop_flag[0]:
                        status_label.config("Generation stopped.")
                        return

                    group = group.sort_values(by="serial_no")
                    # Save the PDF inside the new subfolder
                    output_path = os.path.join(output_subfolder, f"Quotation_{trans_no}.pdf")
                    header = {
                        "quotationcode": trans_no,
                        "account_id": group.iloc[0].get("account_id", ""),
                        "contact1": group.iloc[0].get("contact1", ""),
                        "contact2": group.iloc[0].get("contact2", ""),
                        "phone1": group.iloc[0].get("phone1", ""),
                        "fax1": group.iloc[0].get("fax1", ""),
                        "trans_dt": group.iloc[0]["trans_dt"].strftime("%m/%d/%Y") if pd.notnull(group.iloc[0].get("trans_dt")) else datetime.today().strftime("%m/%d/%Y")
                    }
                    generate_pdf(group, output_path, header)
                    done += 1
                    progress_bar['value'] = done
                    text_output.insert(tk.END, f"Generated: Quotation_{trans_no}.pdf ({done}/{unique_trans})\n")
                    text_output.see(tk.END)

                if not stop_flag[0]:
                    os.startfile(output_dir)
                    messagebox.showinfo("Success", f"{unique_trans} PDF(s) saved to: {output_subfolder}")
                    root.after(500, root.destroy)

            except Exception as e:
                messagebox.showerror("Error", str(e))
                status_label.config(text="Error occurred!")
            finally:
                btn_generate.config(state='normal')

        Thread(target=process).start()

    def open_calendar(type):
        def select_date():
            nonlocal cal
            # Use the correct global variables for start/end date
            global selected_start_date, selected_end_date
            selected_date_str = cal.get_date()  # Get date as string
            
            if type == "start":
                selected_start_date = selected_date_str
                start_date_label.config(text=f"Selected: {selected_start_date}")
            else:
                selected_end_date = selected_date_str
                end_date_label.config(text=f"Selected: {selected_end_date}")
            popup.destroy()

        def reset_date():
            # Use the correct global variables for start/end date
            global selected_start_date, selected_end_date
            if type == "start":
                selected_start_date = None
                start_date_label.config(text="") # Clear label text
            else:
                selected_end_date = None
                end_date_label.config(text="") # Clear label text
            popup.destroy()

        popup = Toplevel(root)
        popup.title(f"Select {type.capitalize()} Transaction Date")
        popup.geometry("280x270")
        cal = Calendar(
            popup,
            selectmode='day',
            year=datetime.today().year,
            month=datetime.today().month,
            day=datetime.today().day,
            mindate=datetime(2001, 7, 16),
            maxdate=datetime.today(),
            date_pattern='mm/dd/yyyy'
        )
        cal.pack(pady=10)

        btn_frame = tk.Frame(popup)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="Select", command=select_date, bg="#4CAF50", fg="white").pack(side='left', padx=5)
        tk.Button(btn_frame, text="Reset", command=reset_date, bg="#f44336", fg="white").pack(side='left', padx=5)

    def handle_esc(event):
        stop_flag[0] = True
        root.destroy()

    global root, calendar_button, btn_generate, progress_bar, status_label, start_quote_entry, end_quote_entry, text_output, start_date_label, end_date_label
    root = tk.Tk()
    root.title("ASHI Quotation PDF Generator")
    # Adjusted window height to accommodate new fields
    root.geometry("600x600") # Increased height and width slightly for better layout
    root.configure(bg="#F4F4F7")
    root.resizable(False, False)
    root.bind("<Escape>", handle_esc)

    # Labels and Entry for Quotation Number Range
    tk.Label(root, text="Quotation Range:", bg="#F4F4F7", font=("Segoe UI", 10, "bold")).pack(pady=(10, 0))

    quote_range_frame = tk.Frame(root, bg="#F4F4F7")
    quote_range_frame.pack(pady=(0, 10))

    tk.Label(quote_range_frame, text="Start:", bg="#F4F4F7", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
    start_quote_entry = tk.Entry(quote_range_frame, width=20)
    start_quote_entry.pack(side="left", padx=(0, 10))

    tk.Label(quote_range_frame, text="End:", bg="#F4F4F7", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
    end_quote_entry = tk.Entry(quote_range_frame, width=20)
    end_quote_entry.pack(side="left")

    # --- Date Range Selection ---
    tk.Label(root, text="Transaction Date Range:", bg="#F4F4F7", font=("Segoe UI", 10, "bold")).pack(pady=(10, 0))

    date_range_frame = tk.Frame(root, bg="#F4F4F7")
    date_range_frame.pack(pady=(0, 10))

    tk.Label(date_range_frame, text="Start Date:", bg="#F4F4F7", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
    start_date_button = tk.Button(date_range_frame, text="Select Start Date", command=lambda: open_calendar("start"), bg="#2196F3", fg="white")
    start_date_button.pack(side="left", padx=(0, 5))
    # Label to display selected start date
    start_date_label = tk.Label(date_range_frame, text="", bg="#F4F4F7", font=("Segoe UI", 9), fg="gray")
    start_date_label.pack(side="left", padx=(0, 10))

    tk.Label(date_range_frame, text="End Date:", bg="#F4F4F7", font=("Segoe UI", 10)).pack(side="left", padx=(0, 5))
    end_date_button = tk.Button(date_range_frame, text="Select End Date", command=lambda: open_calendar("end"), bg="#2196F3", fg="white")
    end_date_button.pack(side="left", padx=(0, 5))
    # Label to display selected end date
    end_date_label = tk.Label(date_range_frame, text="", bg="#F4F4F7", font=("Segoe UI", 9), fg="gray")
    end_date_label.pack(side="left")

    btn_generate = tk.Button(root, text="Generate Report", command=generate_report, bg="#4C5BD4", fg="white", font=("Segoe UI", 10, "bold"))
    btn_generate.pack(pady=8)

    progress_bar = ttk.Progressbar(root, mode='determinate', length=380)
    progress_bar.pack(pady=(10, 5))

    status_label = tk.Label(root, text="", fg="green", bg="#F4F4F7", font=("Segoe UI", 9))
    status_label.pack(pady=(0, 5))

    tk.Label(root, text="PRESS ESC TO END REPORT GENERATION", bg="#F4F4F7", fg="gray", font=("Segoe UI", 9, "italic")).pack(pady=(0, 8))

    tk.Label(root, text="Generated PDFs:", bg="#F4F4F7", font=("Segoe UI", 10, "bold")).pack()
    text_output = tk.Text(root, height=10, width=64, wrap='word')
    text_output.pack(pady=(0, 10))

    root.mainloop()

run_gui()
