import os
import pymupdf  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
import re
from collections import defaultdict
# === File Picker ===
Tk().withdraw()
pdf_path = askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf")])
if not pdf_path:
    print(" No file selected. Exiting.")
    exit()
# === Ask for Save Location ===
print(" Please select where you want to save the Split_PDFs folder...")
save_location = askdirectory(title="Select Location to Save Split PDFs")
if not save_location:
    print(" No save location selected. Exiting.")
    exit()
# === Create Output Folder ===
output_folder = os.path.join(save_location, "Split_PDFs")
os.makedirs(output_folder, exist_ok=True)
# === Step 1: Scan PDF and map pages to Quotation Numbers ===
print(" Scanning PDF pages...")
doc = pymupdf.open(pdf_path)
quote_map = defaultdict(list)
for i in range(len(doc)):
    text = doc[i].get_text()
    match = re.search(r"QUOTATION\s+NUMBER\s*[:\-]?\s*(\d+)", text, re.IGNORECASE)
    if match:
        quote_num = match.group(1)
        quote_map[quote_num].append(i)
    else:
        print(f" Quotation Number not found on page {i + 1}")
doc.close()
# === Step 2: Split PDF using PyPDF2 ===
print(" Splitting PDF...")
reader = PdfReader(pdf_path)
for quote_num, pages in quote_map.items():
    writer = PdfWriter()
    for page_num in pages:
        writer.add_page(reader.pages[page_num])
    output_file = os.path.join(output_folder, f"Quotation_{quote_num}.pdf")
    with open(output_file, "wb") as f:
        writer.write(f)
    print(f"Created: {output_file} ({len(pages)} page(s))")
print("All quotations split successfully.")
print(f"PDFs have been saved in: {output_folder}")
