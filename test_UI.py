import customtkinter as ctk
import subprocess
import threading
import os
import datetime
import itertools
import re
import sys

# === Theme Setup ===
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# === App Initialization ===
app = ctk.CTk()
app.title("ASHI Script Playground")
app.geometry("1080x700")
app.minsize(900, 600)
app.resizable(True, True)

# === Color Palette ===
LIGHT_BG = "#d4e6f1"
DARK_TEXT = "#000000"
LIGHT_PANEL = "#FFFFFF"
INPUT_BOX = "#F0F0F0"
BUTTON_PRIMARY = "#fbfcfc"
BUTTON_HOVER = "#808b96"
BUTTON_SELECTED = "#808b96"
BUTTON_TEXT_DEFAULT = "black"
BUTTON_TEXT_SELECTED = "#fdfefe"
RUN_BTN = "#FF6B6B"
RUN_BTN_HOVER = "#CC5252"

# === Globals
selected_script = None
button_refs = {}
loading = False
script_buttons_disabled = False

# === Script List
script_names = [
    "Split PDF",
    "4X4 Picture Report",
    "4X4 Picture Report with Sales & Stock",
    "4×4 Picture Report with Barcode and Tag Price",
    "Top Style Picture Report with Barcode",
    "Top Style Picture Report",
    "Display Tray Picture Report",
    "Display Tray Sticker With Barcode"
]

scripts_info = {
    "Split PDF": {
        "description": "This Script Generates multiple PDFs grouped by quotation number from a single input PDF.",
        "path": "Scripts\\split_pdf.py"
    },
    "4X4 Picture Report": {
        "description": "Generates a 4x4 picture report with item code, images, description, price, and quoted quantity.",
        "path": "Scripts\\4x4_picture_report.py"
    },
    "4X4 Picture Report with Sales & Stock": {
        "description": "Includes Style Code, 4-year sales quantity, stock in hand, open sales order, and open BOL.",
        "path": "Scripts\\4X4 Picture Report with Sales & Stock.py"
    },
    "4×4 Picture Report with Barcode and Tag Price": {
        "description": "Generates 4x4 picture report with barcodes and tag prices.",
        "path": "Scripts\\4x4_Picture_Report_with_Barcode_and_Tag_Price.py"
    },
    "Top Style Picture Report with Barcode": {
        "description": "Extracts product details with barcode and image for top styles",
        "path": "Scripts\\Top Style Picture Report with Barcode.py"
    },
    "Top Style Picture Report": {
        "description": "Generates top style image report without barcode only with Price.",
        "path": "Scripts\\Top Style Picture Report.py"
    },
    "Display Tray Picture Report": {
        "description": "Generates picture report from Id, Name, barcode input.",
        "path": "Scripts\\Display Tray Picture Report.py"
    },
    "Display Tray Sticker With Barcode": {
        "description": "Creates PDF stickers with barcode layout for display trays.",
        "path": "Scripts\\Display_Tray_Sticker_With_Barcode.py"
    }
}

# === Helpers
def set_buttons_state(state):
    global script_buttons_disabled
    run_btn.configure(state=state)
    script_buttons_disabled = (state == "disabled")

def animate_loading():
    for frame in itertools.cycle(["⏳ Please wait.", "⏳ Please wait..", "⏳ Please wait..."]):
        if not loading:
            loading_label.configure(text="")
            break
        loading_label.configure(text=frame)
        app.update()
        app.after(500)

# === Script Selector
def on_script_click(name):
    global selected_script
    if script_buttons_disabled:
        return
    selected_script = name
    script_entry.configure(state="normal")
    script_entry.delete("1.0", "end")
    script_entry.insert("end", scripts_info[name]["description"])
    script_entry.configure(state="disabled")
    output_box.delete("1.0", "end")
    loading_label.configure(text="")
    for btn_name, (frame, label) in button_refs.items():
        frame.configure(fg_color=BUTTON_SELECTED if btn_name == name else BUTTON_PRIMARY)
        label.configure(text_color=BUTTON_TEXT_SELECTED if btn_name == name else BUTTON_TEXT_DEFAULT)

def on_enter(e, name):
    if not script_buttons_disabled and name != selected_script:
        frame, label = button_refs[name]
        frame.configure(fg_color=BUTTON_HOVER)
        label.configure(text_color=BUTTON_TEXT_SELECTED)

def on_leave(e, name):
    if not script_buttons_disabled and name != selected_script:
        frame, label = button_refs[name]
        frame.configure(fg_color=BUTTON_PRIMARY)
        label.configure(text_color=BUTTON_TEXT_DEFAULT)

# === Script Execution
def run_script():
    global loading
    if not selected_script:
        output_box.insert("end", "❌ Please select a script.\n")
        return

    def target():
        global loading
        try:
            path = scripts_info[selected_script]["path"]
            if not os.path.exists(path):
                output_box.insert("end", f"⚠️ File not found: {path}\n")
                loading = False
                return

            set_buttons_state("disabled")
            process = subprocess.Popen([sys.executable, path],
                                       stdout=subprocess.PIPE,
                                       stderr=subprocess.STDOUT,
                                       universal_newlines=True,
                                       bufsize=1)

            for line in iter(process.stdout.readline, ''):
                output_box.insert("end", line)
                output_box.see("end")

            process.stdout.close()
            process.wait()
        except Exception as e:
            output_box.insert("end", f"❌ Error: {e}\n")
        finally:
            loading = False
            set_buttons_state("normal")

    output_box.delete("1.0", "end")
    loading = True
    threading.Thread(target=target).start()
    threading.Thread(target=animate_loading).start()

# === Layout
app.configure(bg=LIGHT_BG)
main_frame = ctk.CTkFrame(app, corner_radius=20, fg_color=LIGHT_BG)
main_frame.pack(fill="both", expand=True, padx=20, pady=20)

left_panel = ctk.CTkFrame(main_frame, width=250, corner_radius=15, fg_color=LIGHT_PANEL)
left_panel.pack(side="left", fill="y", padx=(15, 7.5), pady=20)

right_panel = ctk.CTkFrame(main_frame, corner_radius=15, fg_color=LIGHT_PANEL)
right_panel.pack(side="right", fill="both", padx=(7.5, 15), expand=True, pady=20)

# === Left Panel Content
ctk.CTkLabel(left_panel, text="Predefined Scripts", text_color=DARK_TEXT, font=("Segoe UI", 21, "bold")).pack(pady=(20, 15))

for idx, name in enumerate(script_names, 1):
    row_frame = ctk.CTkFrame(left_panel, fg_color=BUTTON_PRIMARY, corner_radius=10)
    row_frame.pack(fill="x", padx=10, pady=5)
    row_frame.bind("<Button-1>", lambda e, n=name: on_script_click(n))
    row_frame.bind("<Enter>", lambda e, n=name: on_enter(e, n))
    row_frame.bind("<Leave>", lambda e, n=name: on_leave(e, n))

    index_label = ctk.CTkLabel(row_frame, text=f"{idx}.", width=25, font=("Segoe UI", 17, "bold"), text_color=DARK_TEXT)
    index_label.pack(side="left", padx=(5, 0))
    index_label.bind("<Button-1>", lambda e, n=name: on_script_click(n))
    index_label.bind("<Enter>", lambda e, n=name: on_enter(e, n))
    index_label.bind("<Leave>", lambda e, n=name: on_leave(e, n))

    label = ctk.CTkLabel(row_frame, text=name, font=("Segoe UI", 17), text_color=BUTTON_TEXT_DEFAULT, anchor="w")
    label.pack(side="left", fill="x", expand=True, padx=5)
    label.bind("<Button-1>", lambda e, n=name: on_script_click(n))
    label.bind("<Enter>", lambda e, n=name: on_enter(e, n))
    label.bind("<Leave>", lambda e, n=name: on_leave(e, n))

    button_refs[name] = (row_frame, label)

run_btn = ctk.CTkButton(left_panel, text="▶ Run Script", fg_color=RUN_BTN, hover_color=RUN_BTN_HOVER,
                        font=("Segoe UI", 17, "bold"), corner_radius=10, text_color="white", command=run_script)
run_btn.pack(pady=(40, 10), padx=10, fill="x")

loading_label = ctk.CTkLabel(left_panel, font=("Segoe UI", 16, "italic"), text_color="#777777", text="")
loading_label.pack(pady=(0, 5))

# === Right Panel Content
ctk.CTkLabel(right_panel, text="Script Description", text_color=DARK_TEXT, font=("Segoe UI", 21, "bold")).pack(pady=(15, 5))

script_entry = ctk.CTkTextbox(right_panel, height=100, font=("Consolas", 16), corner_radius=10, fg_color=INPUT_BOX, text_color=DARK_TEXT)
script_entry.pack(fill="x", padx=20, pady=(5, 10))
script_entry.insert("end", "")
script_entry.configure(state="disabled")

ctk.CTkLabel(right_panel, text="Message", text_color=DARK_TEXT, font=("Segoe UI", 21, "bold")).pack(pady=(10, 5))

output_box = ctk.CTkTextbox(right_panel, height=350, font=("Consolas", 16), corner_radius=10, fg_color=INPUT_BOX, text_color=DARK_TEXT)
output_box.pack(fill="both", expand=True, padx=20, pady=(5, 20))

# === Clock
def update_time():
    current_time = datetime.datetime.now().strftime("%d %b %Y | %I:%M:%S %p")
    datetime_label.configure(text=current_time)
    app.after(1000, update_time)

datetime_label = ctk.CTkLabel(left_panel, font=("Segoe UI", 14), text_color="#555555", anchor="center")
datetime_label.pack(side="bottom", pady=(10, 15))
update_time()

# === Run the App
app.mainloop()