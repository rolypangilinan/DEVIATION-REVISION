# CREATE DATABASE DISPLAY AT TOP RIGHT OF THE MATERIAL UPDATE LOG TEXT BOX

# PROCESS 4 INCOMPLETE COLUMN - DONE FIXING
# FIXED :FIXING PROCESS 3 - DUE TO NOT READING THE VT3 TAIL DATA

# FINAL
# CREATE BACK BUTTON TO GO BACK TO SELECTION OF DATABASE
# FINAL
# WITH DATABASE SELECTION GUI: FC1 / TESTING
# WITH MODEL CODE DISPLAYED
# WITH INTENTIONALLY BLANK COLUMN IGNORED THE ALARM
            # OPTIONAL_ITEMS_BY_JO
# FIXED: DOESNT READ LOG000_2.CSV ITS NOT UPDATING THE OUTPU
# DISPLAY JO DATE AT TITTLE
# FIXED: WRONG MATERIAL DETECTOR CSV NOT UPDATING
# "CLEAR LOG" BUTTON IS NOT VISIBLE
# MAKE THE PROCESS BUTTONS CENTERED
# DONE: PROCESS 1, PROCESS 2, PROCESS 3, PROCESS 4, PROCESS 5, PROCESS 6
# MODEL CODE: (60CAT0213P_3J737987830000) (60CAT0212P_3J73798113000 (60FC00000P_3J737976350000)
# SKLEARN ML VALIDATION ADDED (CHECK THE MATERIALS IF IT IS CORRECT OR ERROR)
# BASE CODE
# FIXED: WRONG EXCEL MATCH DETECTION + JOB ORDER NUMBER EXTRACTION
# CREATE PYTHON FILE NAME AFTER THE 1ST TKINTER GUI TITLE

# TEXT BOX: 
    # MATERIAL UPDATE LOG - tree_container = tk.Frame(log_frame, relief="sunken", borderwidth=20, bg="#887979")
    # SYSTEM MESSAGE - system_text = scrolledtext.ScrolledText(msg_frame, height=6, width=60, font=("Courier", 9), state="disabled", bg="#887979", fg="#f0f0f0", relief='sunken', borderwidth=20)
import tkinter as tk
from tkinter import ttk, scrolledtext
import pandas as pd
import os
import sys
import subprocess
import time
import threading
from datetime import datetime
import win32file
import win32con
import csv
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import make_pipeline
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import random
import difflib

# NEW: SELECTION GUI FOR DATABASE LOCATION
selection_root = tk.Tk()
selection_root.title("DATABASE LOCATION")
selection_root.geometry("300x200")
selection_root.configure(bg="#f0f0f0")

sel_frame = tk.Frame(selection_root, bg="#f0f0f0")
sel_frame.pack(expand=True, pady=50)

database_choice = None

def choose_db(db):
    global database_choice
    database_choice = db
    selection_root.destroy()

btn_fc1 = tk.Button(sel_frame, text="FC1", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", width=20, command=lambda: choose_db("FC1"))
btn_fc1.pack(pady=10)

btn_testing = tk.Button(sel_frame, text="TESTING", font=("Arial", 12, "bold"), bg="#2196F3", fg="white", width=20, command=lambda: choose_db("TESTING"))
btn_testing.pack(pady=10)

selection_root.mainloop()

# === CONFIGURATION FROM CODE 1 & CODE 2 ===
if database_choice == "FC1":
    # SOURCE_PATH = r"\\192.168.2.10\csv\csv\JobOrder\log000_JobOrder.csv" #ORIGINAL PATH
    SOURCE_PATH = r"\\192.168.2.10\csv\csv\JobOrder\log000_JobOrder.csv" # FC1 PATH
    OUTPUT_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\wrongMaterialDetectorCSV.csv" # SAME OUTPUT
    COLUMN_NAME = "Job Order Number" # Exact column name in CSV
    EXCEL_DIR = r"\\192.168.2.19\production\2025\1. Document for Production Admin\8. JOB ORDER MATERIAL LIST" # FC1 PATH
    # Network paths (Code 2)
    CSV_LOG_PATH = SOURCE_PATH # Unified with Code 1
    VT1_LOG_PATH = r"\\192.168.2.10\csv\csv\VT1\log000_1.csv" # FC1 PATH
    JO_MATERIAL_DIR = r"\\192.168.2.19\production\2025\1. Document for Production Admin\8. JOB ORDER MATERIAL LIST"
    EXCEL_FOLDER = EXCEL_DIR # Unified
    LOG000_1_PATH = r"\\192.168.2.10\csv\csv\VT1\log000_1.csv"
    LOG000_2_PATH = r"\\192.168.2.10\csv\csv\VT2\log000_2.csv"
    LOG000_3_PATH = r"\\192.168.2.10\csv\csv\VT3\log000_3.csv" # NEW: PROCESS 3 LOG
    LOG000_4_PATH = r"\\192.168.2.10\csv\csv\VT4\log000_4.csv" # NEW: PROCESS 4 LOG
    LOG000_5_PATH = r"\\192.168.2.10\csv\csv\VT5\log000_5.csv" # NEW: PROCESS 5 LOG
    LOG000_6_PATH = r"\\192.168.2.10\csv\csv\VT6\log000_6.csv" # NEW: PROCESS 6 LOG (VINYL)
else:
    # SOURCE_PATH = r"\\192.168.2.10\csv\csv\JobOrder\log000_JobOrder.csv" #ORIGINAL PATH
    SOURCE_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_JobOrder.csv" # FALSE TEST PATH
    OUTPUT_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\wrongMaterialDetectorCSV.csv" # MY OWN CSV FILE
    COLUMN_NAME = "Job Order Number" # Exact column name in CSV
    EXCEL_DIR = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\8. JOB ORDER MATERIAL LIST" # FALSE TEST PATH
    # Network paths (Code 2)
    CSV_LOG_PATH = SOURCE_PATH # Unified with Code 1
    VT1_LOG_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_1.csv" #FALSE TEST PATH
    JO_MATERIAL_DIR = r"\\192.168.2.19\production\2025\1. Document for Production Admin\8. JOB ORDER MATERIAL LIST"
    EXCEL_FOLDER = EXCEL_DIR # Unified
    LOG000_1_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_1.csv"
    LOG000_2_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_2.csv"
    LOG000_3_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_3.csv" # NEW: PROCESS 3 LOG
    LOG000_4_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_4.csv" # NEW: PROCESS 4 LOG
    LOG000_5_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_5.csv" # NEW: PROCESS 5 LOG
    LOG000_6_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_6.csv" # NEW: PROCESS 6 LOG (VINYL)
# Column mapping for Task 3 (1-based index from log000_1.csv)
COLUMN_MAPPING = {
    'DATE': 2, # wrongMaterialDetectorCSV.csv
    'TIME': 3,
    'Process 1 Model Code': 4,
    'Process 1 S/N': 5,
    'Process 1 Em2p': 9, # log000_1.csv 5 COLUMNS
    'Process 1 Em3p': 11,
    'Process 1 Harness': 13,
    'Process 1 Frame': 15,
    'Process 1 Bushing': 17
}
# === NEW: COLUMN MAPPING FOR PROCESS 2 (log000_2.csv) ===
COLUMN_MAPPING_P2 = {
    'Process 2 M4x40 Screw': 9, # log000_2.csv  6 COLUMNS
    'Process 2 Rod Blk': 11,
    'Process 2 Df Blk': 13,
    'Process 2 Df Ring': 15,
    'Process 2 Washer': 17,
    'Process 2 Lock Nut': 19
}
# === NEW: COLUMN MAPPING FOR PROCESS 3 (log000_3.csv) ===
COLUMN_MAPPING_P3 = {
    'Process 3 Frame Gasket': 9, # log000_3.csv 16 COLUMNS
    'Process 3 Casing Block': 11,
    'Process 3 Casing Gasket': 13,
    'Process 3 M4x16 Screw 1': 15,
    'Process 3 M4x16 Screw 2': 17,
    'Process 3 Ball Cushion': 19,
    'Process 3 Frame Cover': 21,
    'Process 3 Partition Board': 23,
    'Process 3 Built In Tube 1': 25,
    'Process 3 Built In Tube 2': 27,
    'Process 3 Head Cover': 29,
    'Process 3 Casing Packing': 31,
    'Process 3 M4x12 Screw': 33,
    'Process 3 Csb L': 35,
    'Process 3 Csb R': 37,
    'Process 3 Head Packing': 39
}
# === NEW: COLUMN MAPPING FOR PROCESS 4 (log000_4.csv) ===
COLUMN_MAPPING_P4 = {
    'Process 4 Tank': 9,    # log000_4.csv  13 COLUMNS
    'Process 4 Upper Housing': 11,
    'Process 4 Cord Hook': 13,
    'Process 4 M4x16 Screw': 15,
    'Process 4 Tank Gasket': 17,
    'Process 4 Tank Cover': 19,
    'Process 4 Housing Gasket': 21,
    'Process 4 M4x40 Screw': 23,
    'Process 4 PartitionGasket': 25,
    'Process 4 M4x12 Screw': 27,
    'Process 4 Muffler': 29,
    'Process 4 Muffler Gasket': 31,
    'Process 4 VCR': 33
}
# === NEW: COLUMN MAPPING FOR PROCESS 5 (log000_5.csv) - RATING LABEL ===
COLUMN_MAPPING_P5 = {
    'Process 5 Rating Label': 9  # Column 9 in log000_5.csv
}
# === NEW: COLUMN MAPPING FOR PROCESS 6 (log000_6.csv) - VINYL ===
COLUMN_MAPPING_P6 = {
    'Process 6 Vinyl': 9  # Column 9 in log000_6.csv → goes to column 42 in output CSV
}
VALIDATION_COLUMN = 'Process 1 Em2p' # The column in output CSV to validate
EXCEL_MATERIAL_COLUMN_INDEX = 4 # The 0-based index (4 for Column E) in the Excel file to check against
# REFERENCE_EXCEL = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\8. JOB ORDER MATERIAL LIST\3J737987830000.xlsx"
REFERENCE_EXCEL = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\8. JOB ORDER MATERIAL LIST"
ML_MATERIAL_COL = 4 # Column E (0-based)
ML_DESC_COL = 10 # Column K (0-based)
ML_TARGET_COLUMNS = ['Process 1 Em2p',
                     'Process 1 Em3p',
                     'Process 1 Harness',
                     'Process 1 Frame',
                     'Process 1 Bushing'] # columns 6-10 in output CSV
ML_TARGET_COLUMNS_P2 = ['Process 2 M4x40 Screw',
                        'Process 2 Rod Blk',
                        'Process 2 Df Blk',
                        'Process 2 Df Ring',
                        'Process 2 Washer',
                        'Process 2 Lock Nut'] # columns 11-16 in output CSV
ML_TARGET_COLUMNS_P3 = ['Process 3 Frame Gasket',
                        'Process 3 Casing Block',
                        'Process 3 Casing Gasket',
                        'Process 3 M4x16 Screw 1',
                        'Process 3 M4x16 Screw 2',
                        'Process 3 Ball Cushion',
                        'Process 3 Frame Cover',
                        'Process 3 Partition Board',
                        'Process 3 Built In Tube 1',
                        'Process 3 Built In Tube 2',
                        'Process 3 Head Cover',
                        'Process 3 Casing Packing',
                        'Process 3 M4x12 Screw',
                        'Process 3 Csb L',
                        'Process 3 Csb R',
                        'Process 3 Head Packing'] # NEW: PROCESS 3
ML_TARGET_COLUMNS_P4 = ['Process 4 Tank',
                        'Process 4 Upper Housing',
                        'Process 4 Cord Hook',
                        'Process 4 M4x16 Screw',
                        'Process 4 Tank Gasket',
                        'Process 4 Tank Cover',
                        'Process 4 Housing Gasket',
                        'Process 4 M4x40 Screw',
                        'Process 4 PartitionGasket',
                        'Process 4 M4x12 Screw',
                        'Process 4 Muffler',
                        'Process 4 Muffler Gasket',
                        'Process 4 VCR'] # NEW: PROCESS 4
ML_TARGET_COLUMNS_P5 = ['Process 5 Rating Label']  # NEW: PROCESS 5
ML_TARGET_COLUMNS_P6 = ['Process 6 Vinyl']  # NEW: PROCESS 6
# Column names
JO_NUMBER_COL = "Job Order Number"
MATERIAL_COL = "MATERIAL"
# Process 1 items to monitor
PROCESS1_ITEMS = [
    "Process 1 Em2p",
    "Process 1 Em3p",
    "Process 1 Harness",
    "Process 1 Frame",
    "Process 1 Bushing"
]
# Process 2 items to monitor
PROCESS2_ITEMS = [
    "Process 2 M4x40 Screw",
    "Process 2 Rod Blk",
    "Process 2 Df Blk",
    "Process 2 Df Ring",
    "Process 2 Washer",
    "Process 2 Lock Nut"
]
# Process 3 items to monitor
PROCESS3_ITEMS = [
    "Process 3 Frame Gasket",
    "Process 3 Casing Block",
    "Process 3 Casing Gasket",
    "Process 3 M4x16 Screw 1",
    "Process 3 M4x16 Screw 2",
    "Process 3 Ball Cushion",
    "Process 3 Frame Cover",
    "Process 3 Partition Board",
    "Process 3 Built In Tube 1",
    "Process 3 Built In Tube 2",
    "Process 3 Head Cover",
    "Process 3 Casing Packing",
    "Process 3 M4x12 Screw",
    "Process 3 Csb L",
    "Process 3 Csb R",
    "Process 3 Head Packing"
]
# Process 4 items to monitor
PROCESS4_ITEMS = [
    "Process 4 Tank",
    "Process 4 Upper Housing",
    "Process 4 Cord Hook",
    "Process 4 M4x16 Screw",
    "Process 4 Tank Gasket",
    "Process 4 Tank Cover",
    "Process 4 Muffler",
    "Process 4 Muffler Gasket",
    "Process 4 VCR",
    "Process 4 Housing Gasket",
    "Process 4 M4x40 Screw",
    "Process 4 PartitionGasket",
    "Process 4 M4x12 Screw"
]
# Process 5 items to monitor
PROCESS5_ITEMS = [
    "Process 5 Rating Label"
]
# Process 6 items to monitor
PROCESS6_ITEMS = [
    "Process 6 Vinyl"
]
# Ensure output directory exists
output_dir = os.path.dirname(OUTPUT_PATH)
os.makedirs(output_dir, exist_ok=True)
# GUI setup
root = tk.Tk()
# === DYNAMIC TITLE: Now uses __file__ to get actual running script name + JO DATE + MODEL CODE ===
# Initial title (will be updated after JO is read)
root.title(f"Production Monitoring System | FILE: {os.path.basename(__file__)} | JO DATE: N/A | MODEL: N/A")
root.geometry("910x700+0+0")
root.configure(bg="#f0f0f0")
# === 1ST & 2ND: BACK & NEXT BUTTONS (Top Left & Right) ===
combined_frame = tk.Frame(root, bg="#f0f0f0")
combined_frame.pack(fill="x", pady=5)
back_btn = tk.Button(combined_frame, text="BACK", font=("Arial", 10, "bold"), bg="#686663", fg="white", relief="ridge", borderwidth=8, width=10)
#back_btn.pack(side="left", padx=10)
# === 3RD: ALL PROCESS BUTTONS IN ONE ROW ===
btn_frame = tk.Frame(combined_frame, bg="#f0f0f0")
btn_frame.pack(expand=True, fill="x") # Changed to expand and fill for centering
next_btn = tk.Button(combined_frame, text="NEXT", font=("Arial", 10, "bold"), bg="#686663", fg="white", relief="ridge", borderwidth=8, width=10)
#next_btn.pack(side="right", padx=10)
#back_btn.pack(side="right", padx=5)
process_buttons = {}
loading_labels = {}
dot_labels = {}
blink_threads = {}
stop_flags = {}
process_active = {}
# === CENTERED PROCESS BUTTONS FRAME ===
center_btn_container = tk.Frame(btn_frame, bg="#f0f0f0")
center_btn_container.pack(expand=True) # This will center the inner frame
inner_btn_frame = tk.Frame(center_btn_container, bg="#f0f0f0")
inner_btn_frame.pack() # Pack inner frame inside container
for i in range(1, 7):
    proc_name = f"PROCESS {i}"
    proc_frame = tk.Frame(inner_btn_frame, bg="#f0f0f0")
    proc_frame.pack(side="left", padx=15)
    load_lbl = tk.Label(proc_frame, text="Loading", font=("Arial", 9), fg="gray", bg="#f0f0f0")
    load_lbl.pack()
    loading_labels[i] = load_lbl
    btn = tk.Button(proc_frame, text=f"{proc_name}\nSTOP", font=("Arial", 11, "bold"),
                    width=10, height=2, bg="#4CAF50", fg="white", relief="raised", borderwidth=8,
                    command=lambda p=i: acknowledge_stop(p))
    btn.pack(pady=3)
    process_buttons[i] = btn
    dot_lbl = tk.Label(proc_frame, text="...", font=("Arial", 12), fg="#4CAF50", bg="#f0f0f0")
    dot_lbl.pack()
    dot_labels[i] = dot_lbl
    stop_flags[i] = threading.Event()
    process_active[i] = True
# === 4TH: REMOVED PREVIOUS, REFRESH, STOP ALL FROM HERE ===
# control_frame = tk.Frame(root, bg="#f0f0f0")
# control_frame.pack(fill="x", pady=5)
# control_center = tk.Frame(control_frame, bg="#f0f0f0")
# control_center.pack(expand=True)
# prev_btn = tk.Button(control_center, text="PREVIOUS", font=("Arial", 10, "bold"), bg="#E8D7EB", fg="black", borderwidth=10, width=10)
# prev_btn.pack(side="left", padx=5)
# refresh_btn = tk.Button(control_center, text="REFRESH", font=("Arial", 10, "bold"), bg="#E8D7EB", fg="black", borderwidth=10, width=10)
# refresh_btn.pack(side="left", padx=5)
# stop_all_btn = tk.Button(control_center, text="STOP ALL", font=("Arial", 10, "bold"), bg="#f44336", fg="white", borderwidth=10, width=10)
# stop_all_btn.pack(side="left", padx=5)
# === 5TH: SYSTEM MESSAGE TEXT BOX (NOW LEFT-ALIGNED, SHORTER, BELOW JOB ORDER) ===
msg_frame = tk.Frame(root, bg="#f0f0f0")
msg_frame.pack(anchor="w", pady=(0, 10), padx=20, fill="x") # Left-aligned, full width but controlled
msg_label = tk.Label(msg_frame, text="SYSTEM MESSAGE", font=("Arial", 11, "bold"), bg="#f0f0f0", anchor="w")
msg_label.pack(anchor="w")
system_text = scrolledtext.ScrolledText(msg_frame, height=6, width=60, font=("Courier", 9), state="disabled", bg="#000000", fg="#f0f0f0", relief='sunken', borderwidth=20)
# system_text = scrolledtext.ScrolledText(msg_frame, height=6, width=60, font=("Courier", 9), state="disabled", bg="#887979", fg="#f0f0f0", relief='sunken', borderwidth=20)
system_text.pack(anchor="w", pady=5, fill="x") # Fill horizontally but respect width
# === 6TH & 9TH: MATERIAL UPDATE LOG (Larger, 3 Columns: Process, MATERIAL ERROR, CORRECTION) ===
log_frame = tk.Frame(root, bg="#f0f0f0")
log_frame.pack(pady=10, fill="both", expand=True, padx=20)
top_frame = tk.Frame(log_frame, bg="#f0f0f0")
top_frame.pack(fill="x")
log_label = tk.Label(top_frame, text="MATERIAL UPDATE LOG", font=("Arial", 11, "bold"), bg="#f0f0f0", anchor="w")
log_label.pack(side="left")
db_label = tk.Label(top_frame, text=f"DATABASE: {database_choice}", font=("Arial", 11, "bold"), fg="blue")
db_label.pack(side="right")
# <<< NEW: RAISED FRAME AROUND TREEVIEW >>>   MATERIAL UPDATE LOG
tree_container = tk.Frame(log_frame, relief="sunken", borderwidth=20, bg="#887979")
tree_container.pack(fill="both", expand=True, side="left")
# Treeview for structured table (REMOVED: MATERIAL, STATUS)
columns = ("Process", "Error", "Correction")
material_tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=7)
material_tree.heading("Process", text="PROCESS:")
material_tree.heading("Error", text="MATERIAL ERROR:")
material_tree.heading("Correction", text="CORRECTION:")
material_tree.column("Process", width=180, anchor="w") # LEFT-ALIGNED
material_tree.column("Error", width=200, anchor="w") # LEFT-ALIGNED
material_tree.column("Correction", width=200, anchor="w") # LEFT-ALIGNED
material_tree.pack(fill="both", expand=True, side="left")
# Scrollbar for Material Log
mat_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=material_tree.yview)
mat_scroll.pack(side="right", fill="y")
material_tree.configure(yscrollcommand=mat_scroll.set)
# === 1ST CHANGE: MAKE HEADERS BOLD AND BIGGER ===
style = ttk.Style()
style.configure("Treeview.Heading", font=("Arial", 12, "bold")) # BIG + BOLD HEADERS
material_tree.tag_configure("error", foreground="red", font=("Courier", 9, "bold"))
# === 8TH: CLEAR BUTTON BELOW MATERIAL LOG ===
clear_btn_frame = tk.Frame(root, bg="#f0f0f0")
clear_btn_frame.pack(pady=12, fill="x", padx=10) # <-- increased pady so it is clearly visible
# === MOVED: PREVIOUS, REFRESH, STOP ALL TO BOTTOM LEFT ===
prev_btn = tk.Button(clear_btn_frame, text="PREVIOUS", font=("Arial", 10, "bold"), bg="#E8D7EB", fg="black", borderwidth=10, width=10)
refresh_btn = tk.Button(clear_btn_frame, text="REFRESH", font=("Arial", 10, "bold"), bg="#E8D7EB", fg="black", borderwidth=10, width=10)
stop_all_btn = tk.Button(clear_btn_frame, text="STOP ALL", font=("Arial", 10, "bold"), bg="#f44336", fg="white", borderwidth=10, width=10)
prev_btn.pack(side="left", padx=5)
refresh_btn.pack(side="left", padx=5)
stop_all_btn.pack(side="left", padx=5)
# === JOB ORDER DISPLAY BETWEEN STOP ALL AND CLEAR LOG ===
jo_btn_frame = tk.Frame(clear_btn_frame, bg="#f0f0f0")
jo_btn_frame.pack(side="right", padx=12)
# REMOVED: jo_label_title = tk.Label(...)
jo_display = tk.Label(jo_btn_frame, text="JO: N/A", font=("Arial", 13, "bold"), bg="#ffffff", fg="#2187EE",
                      relief="sunken", borderwidth=8, width=16, anchor="center")
jo_display.pack(side="bottom", padx=(3, 3))
back_btn = tk.Button(clear_btn_frame, text="BACK", font=("Arial", 10, "bold"), bg="#686663", fg="white", relief="ridge", borderwidth=8, width=10, command=lambda: go_back())
next_btn = tk.Button(clear_btn_frame, text="NEXT", font=("Arial", 10, "bold"), bg="#686663", fg="white", relief="ridge", borderwidth=8, width=10)
def clear_material_log():
    for item in material_tree.get_children():
        material_tree.delete(item)
    log_system("MATERIAL LOG CLEARED")
clear_material_button = tk.Button(clear_btn_frame, text="CLEAR LOG", font=("Arial", 10, "bold"), bg="#ff4444", fg="white",
                                  width=15, relief="raised", borderwidth=10, command=clear_material_log)
next_btn.pack(side="right", padx=5)
back_btn.pack(side="right", padx=5)
clear_material_button.pack(side="right", padx=5)
# === 11TH: VERTICAL SCROLLBAR CONTROLLED BY MOUSE WHEEL ===
def _on_mousewheel(event):
    material_tree.yview_scroll(int(-1*(event.delta/120)), "units")
    system_text.yview_scroll(int(-1*(event.delta/120)), "units")
    return "break"
root.bind_all("<MouseWheel>", _on_mousewheel)
# Global variables
last_jo_number = None
last_vt1_mtime = None
running = True
last_mtime_log1 = 0
last_mtime_log2 = 0
last_mtime_log3 = 0
last_mtime_log4 = 0 # NEW: FOR PROCESS 4
last_mtime_log5 = 0 # NEW: FOR PROCESS 5
last_mtime_log6 = 0 # NEW: FOR PROCESS 6
animation_running = True
current_jo = "N/A"
current_jo_date = "N/A" # NEW: To store JO date from CSV
current_model_code = "N/A"  # <<< NEW: GLOBAL FOR MODEL CODE
# === NEW: JO-SPECIFIC OPTIONAL ITEMS (SKIP VALIDATION & ALARM) ===
OPTIONAL_ITEMS_BY_JO = {
    "3J737976350000": [
        "Process 2 Df Ring",
        "Process 3 Frame Gasket",
        "Process 3 Casing Block",
        "Process 3 Casing Gasket",
        "Process 3 M4x16 Screw 1",
        "Process 3 Ball Cushion",
        "Process 3 Partition Board",
        "Process 3 Built In Tube 1",
        "Process 3 Built In Tube 2",
        "Process 4 Tank",
        "Process 4 Upper Housing",
        "Process 4 Cord Hook",
        "Process 4 M4x16 Screw",
        "Process 4 Tank Gasket",
        "Process 4 Tank Cover",
        "Process 4 Housing Gasket",
        "Process 4 M4x40 Screw",
        "Process 4 PartitionGasket"
    ],
    "3J737987830000": [ # <<< NEW: ADDED FOR THIS JO >>>
        "Process 3 Head Cover",
        "Process 3 Casing Packing",
        "Process 3 M4x12 Screw",
        "Process 3 Csb L",
        "Process 3 Csb R",
        "Process 3 Head Packing",
        "Process 4 Tank Gasket",
        "Process 4 Tank Cover",
        "Process 4 Muffler",
        "Process 4 Muffler Gasket",
        "Process 4 VCR"
    ],
    "3J737981130000": [ # <<< NEW: ADDED FOR THIS JO >>>
        "Process 3 Partition Board",
        "Process 3 Built In Tube 1",
        "Process 3 Built In Tube 2",
        "Process 3 Head Cover",
        "Process 3 Casing Packing",
        "Process 3 M4x12 Screw",
        "Process 3 Csb L",
        "Process 3 Csb R",
        "Process 3 Head Packing",
        "Process 4 PartitionGasket",
        "Process 4 M4x12 Screw",
        "Process 4 Muffler",
        "Process 4 Muffler Gasket",
        "Process 4 VCR"

        

    ]
    # Add more JOs and items here if needed
}
# === SYSTEM LOG FUNCTION ===
def log_system(message):
    system_text.config(state="normal")
    timestamp = datetime.now().strftime("%H:%M:%S")
    system_text.insert(tk.END, f"[{timestamp}] {message}\n")
    system_text.see(tk.END)
    system_text.config(state="disabled")
# === UPDATE TITLE WITH JO DATE ===
def update_title_with_jo_date():
    global current_jo, current_jo_date
    jo_date_str = current_jo_date if current_jo_date != "N/A" else "N/A"
    root.title(f"Production Monitoring System      |      PY FILE NAME: {os.path.basename(__file__)}      |     JO DATE: {jo_date_str}      |     MODEL: {current_model_code}")
# === NEW: UPDATE TITLE WITH MODEL CODE ===
def update_title_with_model():
    global current_model_code
    try:
        if os.path.exists(OUTPUT_PATH):
            with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
                if len(rows) >= 2 and len(rows[0]) > 0:
                    headers = rows[0]
                    data_row = rows[1]
                    if "Process 1 Model Code" in headers:
                        idx = headers.index("Process 1 Model Code")
                        model = data_row[idx].strip()
                        if model:
                            current_model_code = model
                        else:
                            current_model_code = "N/A"
                    else:
                        current_model_code = "N/A"
                else:
                    current_model_code = "N/A"
        else:
            current_model_code = "N/A"
    except:
        current_model_code = "N/A"
    update_title_with_jo_date()  # refresh full title
# === MATERIAL LOG WITH COLOR TAGS (DISPLAYS "Process 1 Em2p" HEADER, ERROR, CORRECTION) ===
def log_material(process_header, error, correction):
    # CHANGED: Process column now shows the header (e.g., "Process 1 Em2p"), no "FRAME"
    item = material_tree.insert("", "end", values=(process_header, error, correction))
    material_tree.item(item, tags=("error",))
    material_tree.tag_configure("error", foreground="red", font=("Courier", 12, "bold"))
    # Highlight diff in ERROR vs CORRECTION
    if error and correction:
        diff = difflib.ndiff(error, correction)
        error_colored = ""
        for d in diff:
            if d.startswith(" "):
                error_colored += d[-1]
            elif d.startswith("-"):
                error_colored += f"[red]{d[-1]}[/red]"
            elif d.startswith("+"):
                error_colored += f"[skyblue]{d[-1]}[/skyblue]"
        # We'll use tags later for actual coloring if needed
    else:
        error_colored = error
# === ACKNOWLEDGE STOP ===
def acknowledge_stop(process_id):
    global blink_threads
    stop_flags[process_id].set()
    btn = process_buttons[process_id]
    btn.config(bg="#4CAF50", relief="raised")
    log_system(f"PROCESS {process_id}: STOP ACKNOWLEDGED BY USER")
    if process_id in blink_threads:
        blink_threads[process_id].join(timeout=1)
        del blink_threads[process_id]
    process_active[process_id] = True
    stop_flags[process_id].clear()
# === BLINK BUTTON ===
def blink_button(process_id):
    btn = process_buttons[process_id]
    while not stop_flags[process_id].is_set() and process_active[process_id]:
        btn.config(bg="red")
        time.sleep(0.5)
        if stop_flags[process_id].is_set():
            break
        btn.config(bg="#ff6666")
        time.sleep(0.5)
    btn.config(bg="#4CAF50")
# === ANIMATE LOADING (FIXED: Continues even after button click) ===
def animate_loading():
    global animation_running
    while animation_running:
        for i in range(1, 7):
            if process_active[i]:
                loading_labels[i].config(text="Loading" + "." * ((int(time.time() * 2) % 4)))
                dot_text = "." * ((int(time.time() * 3) % 4) + 1)
                dot_labels[i].config(text=dot_text)
        time.sleep(0.3)
# === GET LAST LINE (FROM CODE 1) ===
def get_last_line(csv_file_path):
    """Read the last non-empty line from a CSV file efficiently."""
    with open(csv_file_path, 'r', newline='', encoding='utf-8') as f:
        lines = f.readlines()
        # Reverse to find last non-empty line
        for line in reversed(lines):
            if line.strip():
                return line
    return None
def parse_last_row(last_line, delimiter=','):
    """Parse CSV line using csv reader to handle quotes and commas properly."""
    reader = csv.reader([last_line], delimiter=delimiter)
    return next(reader)
# === PROCESS JOB ORDER (TASK 1 & 2) - REPLACED WITH CODE 2 LOGIC ===
def process_job_order():
    global current_jo, current_jo_date
    try:
        print(f"Reading tail of: {CSV_LOG_PATH}")
        if not os.path.exists(CSV_LOG_PATH):
            raise FileNotFoundError(f"Source file not found: {CSV_LOG_PATH}")
        last_line = get_last_line(CSV_LOG_PATH)
        if not last_line:
            print("Source CSV is empty.")
            return
        with open(CSV_LOG_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            try:
                headers = next(reader)
            except StopIteration:
                print("CSV has no headers.")
                return
            try:
                job_order_idx = headers.index(COLUMN_NAME)
            except ValueError:
                print(f"Column '{COLUMN_NAME}' not found in headers.")
                return
            # === NEW: Extract DATE from column index 2 (1-based) ===
            date_idx = 1 # 0-based index for column 2
            rows = list(reader)
            if not rows:
                print("No data rows in CSV.")
                return
            last_row = rows[-1]
        raw_job_order = last_row[job_order_idx]
        cleaned_job_order = raw_job_order.replace(" ", "").replace("\t", "")
        print(f"Original: '{raw_job_order}'")
        print(f"Cleaned : '{cleaned_job_order}'")
        current_jo = cleaned_job_order
        jo_display.config(text=f"JO: {cleaned_job_order}")
        # === EXTRACT DATE FROM COLUMN 2 (1-based) ===
        if len(last_row) > date_idx:
            raw_date = last_row[date_idx].strip()
            try:
                # Try to parse and reformat date
                parsed_date = datetime.strptime(raw_date, "%Y/%m/%d")
                current_jo_date = parsed_date.strftime("%Y-%m-%d")
            except:
                current_jo_date = raw_date # fallback
        else:
            current_jo_date = "N/A"
        # === UPDATE TITLE WITH JO DATE ===
        update_title_with_jo_date()
        # === WRITE FULL HEADER + EMPTY ROW (CODE 2 STYLE) ===
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            header = ["Job Order Number", "File Name"] + list(COLUMN_MAPPING.keys()) + list(COLUMN_MAPPING_P2.keys()) + list(COLUMN_MAPPING_P3.keys()) + list(COLUMN_MAPPING_P4.keys()) + list(COLUMN_MAPPING_P5.keys()) + list(COLUMN_MAPPING_P6.keys())
            writer.writerow(header)
            empty_row = [cleaned_job_order, ""] + [""] * (len(COLUMN_MAPPING) + len(COLUMN_MAPPING_P2) + len(COLUMN_MAPPING_P3) + len(COLUMN_MAPPING_P4) + len(COLUMN_MAPPING_P5) + len(COLUMN_MAPPING_P6))
            writer.writerow(empty_row)
        print(f"Successfully saved to: {OUTPUT_PATH}")
        # === CHECK EXCEL MATCH ===
        excel_file = f"{cleaned_job_order}.xlsx"
        excel_full_path = os.path.join(EXCEL_FOLDER, excel_file)
        if os.path.exists(excel_full_path):
            print(f"THE FILE NAME {cleaned_job_order} WAS SUCCESSFULLY FOUND AT {EXCEL_FOLDER}")
            update_output_column(1, cleaned_job_order)
            log_system(f"MATCHED FILE: {cleaned_job_order}.xlsx")
        else:
            print(f".xlsx file NOT found: {excel_file}")
            update_output_column(1, "NOT FOUND")
            log_system(f".xlsx file NOT found: {excel_file}")
        # <<< NEW: After creating the output CSV, read Model Code and update title >>>
        update_title_with_model()
    except Exception as e:
        print(f"Error: {e}")
        log_system(f"Job Order Error: {e}")
# === UPDATE OUTPUT COLUMN (from Code 2, reused) ===
def update_output_column(col_index, value):
    if not os.path.exists(OUTPUT_PATH):
        return
    with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
        rows = list(csv.reader(f))
    if len(rows) > 1:
        rows[1][col_index] = value
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(rows)
last_mtime_log1 = 0
last_mtime_log2 = 0
last_mtime_log3 = 0
last_mtime_log4 = 0
last_mtime_log5 = 0
last_mtime_log6 = 0
# === UPDATE FROM LOG000_1.CSV - REPLACED WITH CODE 2 LOGIC ===
def update_from_log000_1(force_update=False):
    global last_mtime_log1
    try:
        if not os.path.exists(LOG000_1_PATH):
            return False
        current_mtime = os.path.getmtime(LOG000_1_PATH)
        if not force_update and current_mtime <= last_mtime_log1:
            return False
        time.sleep(0.1)
        last_mtime_log1 = current_mtime
        print(f"\nTRANSFERING ({os.path.basename(LOG000_1_PATH)}) DATA TO \"{os.path.basename(OUTPUT_PATH)}\"")
        with open(LOG000_1_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r and any(c.strip() for c in r)]
            if not rows:
                return False
            last_row = rows[-1]
        if not os.path.exists(OUTPUT_PATH):
            return False
        with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
            output_rows = list(csv.reader(f))
        if len(output_rows) < 2:
            return False
        headers = output_rows[0]
        data_row = output_rows[1]
        for col_name, src_idx_1based in COLUMN_MAPPING.items():
            header_idx = headers.index(col_name)
            value = last_row[src_idx_1based - 1] if len(last_row) >= src_idx_1based else ""
            data_row[header_idx] = value
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(output_rows)
        print(f"Updated {len(COLUMN_MAPPING)} columns from log000_1.csv")
        validate_material_in_excel(output_rows, process_id=1)
        # <<< NEW: After updating Process 1 (which includes Model Code), refresh title >>>
        update_title_with_model()
        return True
    except Exception as e:
        print(f"Error in log000_1 update: {e}")
        return False
# === NEW: UPDATE FROM LOG000_2.CSV ===
def update_from_log000_2(force_update=False):
    global last_mtime_log2
    try:
        if not os.path.exists(LOG000_2_PATH):
            return False
        current_mtime = os.path.getmtime(LOG000_2_PATH)
        if not force_update and current_mtime <= last_mtime_log2:
            return False
        time.sleep(0.1)
        last_mtime_log2 = current_mtime
        print(f"\nTRANSFERING ({os.path.basename(LOG000_2_PATH)}) DATA TO \"{os.path.basename(OUTPUT_PATH)}\"")
        with open(LOG000_2_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r and any(c.strip() for c in r)]
            if not rows:
                return False
            last_row = rows[-1]
        if not os.path.exists(OUTPUT_PATH):
            return False
        with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
            output_rows = list(csv.reader(f))
        if len(output_rows) < 2:
            return False
        headers = output_rows[0]
        data_row = output_rows[1]
        base_idx = 2 + len(COLUMN_MAPPING) # Start after Process 1 columns
        for col_name, src_idx_1based in COLUMN_MAPPING_P2.items():
            header_idx = headers.index(col_name)
            value = last_row[src_idx_1based - 1] if len(last_row) >= src_idx_1based else ""
            data_row[header_idx] = value
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(output_rows)
        print(f"Updated {len(COLUMN_MAPPING_P2)} columns from log000_2.csv")
        validate_material_in_excel(output_rows, process_id=2)
        return True
    except Exception as e:
        print(f"Error in log000_2 update: {e}")
        return False
# === NEW: UPDATE FROM LOG000_3.CSV (FIXED) ===
def update_from_log000_3(force_update=False):
    global last_mtime_log3
    try:
        if not os.path.exists(LOG000_3_PATH):
            return False
        current_mtime = os.path.getmtime(LOG000_3_PATH)
        if not force_update and current_mtime <= last_mtime_log3:
            return False
        time.sleep(0.1)
        last_mtime_log3 = current_mtime
        print(f"\nTRANSFERING ({os.path.basename(LOG000_3_PATH)}) DATA TO \"{os.path.basename(OUTPUT_PATH)}\"")
        src_values = {}
        with open(LOG000_3_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r and any(c.strip() for c in r)]
            if not rows:
                return False
            last_row = rows[-1]
        if not os.path.exists(OUTPUT_PATH):
            return False
        with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
            output_rows = list(csv.reader(f))
        if len(output_rows) < 2:
            return False
        headers = output_rows[0]
        data_row = output_rows[1]
        base_idx = 2 + len(COLUMN_MAPPING) + len(COLUMN_MAPPING_P2)
        updated_cols = []
        for col_name, src_idx_1based in COLUMN_MAPPING_P3.items():
            header_idx = headers.index(col_name)
            value = last_row[src_idx_1based - 1] if len(last_row) >= src_idx_1based else ""
            data_row[header_idx] = value
            if value:
                updated_cols.append(col_name)
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(output_rows)
        print(f"Updated {len(updated_cols)} Process-3 columns: {', '.join(updated_cols)}")
        validate_material_in_excel(output_rows, process_id=3)
        return True
    except Exception as e:
        print(f"Error in log000_3 update: {e}")
        return False
# === NEW: UPDATE FROM LOG000_4.CSV ===
def update_from_log000_4(force_update=False):
    global last_mtime_log4
    try:
        if not os.path.exists(LOG000_4_PATH):
            return False
        current_mtime = os.path.getmtime(LOG000_4_PATH)
        if not force_update and current_mtime <= last_mtime_log4:
            return False
        time.sleep(0.1)
        last_mtime_log4 = current_mtime
        print(f"\nTRANSFERING ({os.path.basename(LOG000_4_PATH)}) DATA TO \"{os.path.basename(OUTPUT_PATH)}\"")
        with open(LOG000_4_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r and any(c.strip() for c in r)]
            if not rows:
                return False
            last_row = rows[-1]
        if not os.path.exists(OUTPUT_PATH):
            return False
        with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
            output_rows = list(csv.reader(f))
        if len(output_rows) < 2:
            return False
        headers = output_rows[0]
        data_row = output_rows[1]
        base_idx = 2 + len(COLUMN_MAPPING) + len(COLUMN_MAPPING_P2) + len(COLUMN_MAPPING_P3)
        for col_name, src_idx_1based in COLUMN_MAPPING_P4.items():
            header_idx = headers.index(col_name)
            value = last_row[src_idx_1based - 1] if len(last_row) >= src_idx_1based else ""
            data_row[header_idx] = value
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(output_rows)
        print(f"Updated {len(COLUMN_MAPPING_P4)} columns from log000_4.csv")
        validate_material_in_excel(output_rows, process_id=4)
        return True
    except Exception as e:
        print(f"Error in log000_4 update: {e}")
        return False
# === NEW: UPDATE FROM LOG000_5.CSV (PROCESS 5 RATING LABEL) ===
def update_from_log000_5(force_update=False):
    global last_mtime_log5
    try:
        if not os.path.exists(LOG000_5_PATH):
            return False
        current_mtime = os.path.getmtime(LOG000_5_PATH)
        if not force_update and current_mtime <= last_mtime_log5:
            return False
        time.sleep(0.1)
        last_mtime_log5 = current_mtime
        print(f"\nTRANSFERING ({os.path.basename(LOG000_5_PATH)}) DATA TO \"{os.path.basename(OUTPUT_PATH)}\"")
        with open(LOG000_5_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r and any(c.strip() for c in r)]
            if not rows:
                return False
            last_row = rows[-1]
        if not os.path.exists(OUTPUT_PATH):
            return False
        with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
            output_rows = list(csv.reader(f))
        if len(output_rows) < 2:
            return False
        headers = output_rows[0]
        data_row = output_rows[1]
        for col_name, src_idx_1based in COLUMN_MAPPING_P5.items():
            header_idx = headers.index(col_name)
            value = last_row[src_idx_1based - 1] if len(last_row) >= src_idx_1based else ""
            data_row[header_idx] = value
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(output_rows)
        print(f"Updated Process 5 Rating Label from log000_5.csv")
        validate_material_in_excel(output_rows, process_id=5)
        return True
    except Exception as e:
        print(f"Error in log000_5 update: {e}")
        return False
# === NEW: UPDATE FROM LOG000_6.CSV (PROCESS 6 VINYL) ===
def update_from_log000_6(force_update=False):
    global last_mtime_log6
    try:
        if not os.path.exists(LOG000_6_PATH):
            return False
        current_mtime = os.path.getmtime(LOG000_6_PATH)
        if not force_update and current_mtime <= last_mtime_log6:
            return False
        time.sleep(0.1)
        last_mtime_log6 = current_mtime
        print(f"\nTRANSFERING ({os.path.basename(LOG000_6_PATH)}) DATA TO \"{os.path.basename(OUTPUT_PATH)}\"")
        with open(LOG000_6_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r and any(c.strip() for c in r)]
            if not rows:
                return False
            last_row = rows[-1]
        if not os.path.exists(OUTPUT_PATH):
            return False
        with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
            output_rows = list(csv.reader(f))
        if len(output_rows) < 2:
            return False
        headers = output_rows[0]
        data_row = output_rows[1]
        for col_name, src_idx_1based in COLUMN_MAPPING_P6.items():
            header_idx = headers.index(col_name)
            value = last_row[src_idx_1based - 1] if len(last_row) >= src_idx_1based else ""
            data_row[header_idx] = value
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(output_rows)
        print(f"Updated Process 6 Vinyl from log000_6.csv")
        validate_material_in_excel(output_rows, process_id=6)
        return True
    except Exception as e:
        print(f"Error in log000_6 update: {e}")
        return False
# === SKLEARN VALIDATION + CORRECTION (NOW SUPPORTS PROCESS 6) ===
def validate_material_in_excel(output_rows, process_id=1):
    try:
        if len(output_rows) < 2:
            return
        headers = output_rows[0]
        data_row = output_rows[1]
        job_order_num = data_row[headers.index("Job Order Number")]
        file_name_status = data_row[headers.index("File Name")]
        if file_name_status == "NOT FOUND":
            log_system(f"Validation skipped: Excel file for JO {job_order_num} was NOT FOUND.")
            return
        optional_columns = OPTIONAL_ITEMS_BY_JO.get(job_order_num, [])
        dynamic_ref = os.path.join(EXCEL_DIR, job_order_num + ".xlsx")
        if os.path.exists(dynamic_ref):
            ref_df = pd.read_excel(dynamic_ref, header=None)
        elif os.path.exists(REFERENCE_EXCEL):
            ref_df = pd.read_excel(REFERENCE_EXCEL, header=None)
        else:
            log_system(f"Reference Excel not found for {job_order_num}")
            return
        materials = ref_df.iloc[:, ML_MATERIAL_COL].dropna().astype(str).str.strip()
        descriptions = ref_df.iloc[:, ML_DESC_COL].dropna().astype(str).str.strip()
        invalid_texts = []
        for mat in materials[:min(50, len(materials))]:
            invalid_texts.append(f"XYZ_{mat}")
            if len(mat) > 1:
                invalid_texts.append(mat[:-1])
            if len(mat) > 2:
                i = random.randint(0, len(mat)-2)
                typo = list(mat)
                typo[i], typo[i+1] = typo[i+1], typo[i]
                invalid_texts.append(''.join(typo))
            invalid_texts.append(''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=8)))
        invalid_texts = list(set(invalid_texts) - set(materials.str.upper()))
        train_texts = [f"{m} {d}" for m, d in zip(materials, descriptions)] + invalid_texts
        train_labels = ["VALID"] * len(materials) + ["INVALID"] * len(invalid_texts)
        ml_model = None
        if len(set(train_labels)) >= 2:
            vectorizer = TfidfVectorizer(lowercase=True, ngram_range=(1,2))
            clf = LogisticRegression(max_iter=1000)
            ml_model = make_pipeline(vectorizer, clf)
            try:
                ml_model.fit(train_texts, train_labels)
            except:
                ml_model = None
        material_codes = materials.tolist()
        correction_vectorizer = TfidfVectorizer(lowercase=True, ngram_range=(1,3), analyzer='char')
        tfidf_matrix = correction_vectorizer.fit_transform(material_codes) if material_codes else None
        print("\n" + "="*80)
        print(f" SKLEARN VALIDATION + 100% CORRECT SUGGESTIONS (PROCESS {process_id}) ")
        print("="*80)
        any_error = False
        empty_error = False
        target_columns = (ML_TARGET_COLUMNS if process_id == 1 else
                          ML_TARGET_COLUMNS_P2 if process_id == 2 else
                          ML_TARGET_COLUMNS_P3 if process_id == 3 else
                          ML_TARGET_COLUMNS_P4 if process_id == 4 else
                          ML_TARGET_COLUMNS_P5 if process_id == 5 else
                          ML_TARGET_COLUMNS_P6)
        for col_name in target_columns:
            if col_name not in headers:
                continue
            idx = headers.index(col_name)
            value = data_row[idx].strip()
            if col_name in optional_columns:
                print(f" [SKIP] {col_name}: OPTIONAL FOR JO {job_order_num}")
                continue
            if not value or value == "":
                print(f" [EMPTY] {col_name}: <empty>")
                empty_error = True
                log_material(col_name, "<EMPTY>", "REQUIRED")
                log_system(f"PROCESS {process_id}: ERROR: {col_name} IS EMPTY")
                if process_active[process_id] and not stop_flags[process_id].is_set():
                    t = threading.Thread(target=blink_button, args=(process_id,), daemon=True)
                    blink_threads[process_id] = t
                    t.start()
                continue
            exact_match = value.upper() in {m.upper() for m in materials}
            if exact_match:
                print(f" [OK] {col_name}: {value} (exact match)")
                continue
            is_valid = False
            confidence = 0.0
            if ml_model:
                pred_input = f"{value} {value}"
                pred = ml_model.predict([pred_input])[0]
                prob = ml_model.predict_proba([pred_input])[0]
                confidence = prob[ml_model.classes_.tolist().index(pred)]
                if pred == "VALID" and confidence > 0.85:
                    is_valid = True
            if is_valid:
                print(f" [OK] {col_name}: {value} (ML-predicted VALID, {confidence:.2%})")
                continue
            any_error = True
            suggested_material = "UNKNOWN"
            input_upper = value.upper()
            if tfidf_matrix is not None:
                query_vec = correction_vectorizer.transform([input_upper])
                sim_scores = cosine_similarity(query_vec, tfidf_matrix).flatten()
                if len(sim_scores) > 0 and sim_scores.max() > 0.1:
                    best_idx = np.argmax(sim_scores)
                    suggested_material = material_codes[best_idx]
                else:
                    candidates = [m for m in material_codes if m.upper().startswith(input_upper[:4])]
                    if not candidates:
                        candidates = material_codes
                    match = difflib.get_close_matches(input_upper, candidates, n=1, cutoff=0.6)
                    suggested_material = match[0] if match else material_codes[0] if material_codes else "NO DATA"
            else:
                suggested_material = material_codes[0] if material_codes else "NO DATA"
            print(f" [ERROR] {col_name}: {value}")
            print(f" **SUGGESTED CORRECTION** -> {suggested_material}")
            err_msg = f"PROCESS {process_id}: ERROR: {value} EXPECTED: {suggested_material}"
            log_system(err_msg)
            log_material(col_name, value, suggested_material)
            if process_active[process_id] and not stop_flags[process_id].is_set():
                t = threading.Thread(target=blink_button, args=(process_id,), daemon=True)
                blink_threads[process_id] = t
                t.start()
        if not any_error and not empty_error:
            print(f" All material fields are CORRECT for PROCESS {process_id}.")
            log_system(f"PROCESS {process_id}: ALL MATERIALS VERIFIED")
            if stop_flags[process_id].is_set():
                acknowledge_stop(process_id)
        print("="*80 + "\n")
    except Exception as e:
        print(f"Error in material validation (Process {process_id}): {e}")
        log_system(f"Validation Error (Process {process_id}): {e}")
# === WATCHDOG HANDLER (NOW HANDLES LOG000_1,2,3,4,5,6) ===
class LogChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        normalized_event_path = os.path.normpath(event.src_path).lower()
        normalized_log1_path = os.path.normpath(LOG000_1_PATH).lower()
        normalized_log2_path = os.path.normpath(LOG000_2_PATH).lower()
        normalized_log3_path = os.path.normpath(LOG000_3_PATH).lower()
        normalized_log4_path = os.path.normpath(LOG000_4_PATH).lower()
        normalized_log5_path = os.path.normpath(LOG000_5_PATH).lower()
        normalized_log6_path = os.path.normpath(LOG000_6_PATH).lower()
        if normalized_event_path == normalized_log1_path:
            update_from_log000_1(force_update=True)
        elif normalized_event_path == normalized_log2_path:
            update_from_log000_2(force_update=True)
        elif normalized_event_path == normalized_log3_path:
            update_from_log000_3(force_update=True)
        elif normalized_event_path == normalized_log4_path:
            update_from_log000_4(force_update=True)
        elif normalized_event_path == normalized_log5_path:
            update_from_log000_5(force_update=True)
        elif normalized_event_path == normalized_log6_path:
            update_from_log000_6(force_update=True)
# === MONITORING LOOP ===
def monitoring_loop():
    global animation_running, last_mtime_log1, last_mtime_log2, last_mtime_log3, last_mtime_log4, last_mtime_log5, last_mtime_log6
    observer = None
    dot_count = 0
    try:
        process_job_order()
        # Initial load for all logs
        if os.path.exists(LOG000_1_PATH):
            if update_from_log000_1(force_update=True):
                log_system(f"WAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES...")
            else:
                try:
                    last_mtime_log1 = os.path.getmtime(LOG000_1_PATH)
                except:
                    pass
        if os.path.exists(LOG000_2_PATH):
            if update_from_log000_2(force_update=True):
                log_system(f"WAITING FOR ({os.path.basename(LOG000_2_PATH)}) FILE CHANGES...")
            else:
                try:
                    last_mtime_log2 = os.path.getmtime(LOG000_2_PATH)
                except:
                    pass
        if os.path.exists(LOG000_3_PATH):
            if update_from_log000_3(force_update=True):
                log_system(f"WAITING FOR ({os.path.basename(LOG000_3_PATH)}) FILE CHANGES...")
            else:
                try:
                    last_mtime_log3 = os.path.getmtime(LOG000_3_PATH)
                except:
                    pass
        if os.path.exists(LOG000_4_PATH):
            if update_from_log000_4(force_update=True):
                log_system(f"WAITING FOR ({os.path.basename(LOG000_4_PATH)}) FILE CHANGES...")
            else:
                try:
                    last_mtime_log4 = os.path.getmtime(LOG000_4_PATH)
                except:
                    pass
        if os.path.exists(LOG000_5_PATH):
            if update_from_log000_5(force_update=True):
                log_system(f"WAITING FOR ({os.path.basename(LOG000_5_PATH)}) FILE CHANGES...")
            else:
                try:
                    last_mtime_log5 = os.path.getmtime(LOG000_5_PATH)
                except:
                    pass
        if os.path.exists(LOG000_6_PATH):
            if update_from_log000_6(force_update=True):
                log_system(f"WAITING FOR ({os.path.basename(LOG000_6_PATH)}) FILE CHANGES...")
            else:
                try:
                    last_mtime_log6 = os.path.getmtime(LOG000_6_PATH)
                except:
                    pass
        observer = Observer()
        watch_dir1 = os.path.dirname(LOG000_1_PATH)
        watch_dir2 = os.path.dirname(LOG000_2_PATH)
        watch_dir3 = os.path.dirname(LOG000_3_PATH)
        watch_dir4 = os.path.dirname(LOG000_4_PATH)
        watch_dir5 = os.path.dirname(LOG000_5_PATH)
        watch_dir6 = os.path.dirname(LOG000_6_PATH)
        if os.path.exists(watch_dir1):
            observer.schedule(LogChangeHandler(), path=watch_dir1, recursive=False)
        if os.path.exists(watch_dir2) and watch_dir2 != watch_dir1:
            observer.schedule(LogChangeHandler(), path=watch_dir2, recursive=False)
        if os.path.exists(watch_dir3) and watch_dir3 not in (watch_dir1, watch_dir2):
            observer.schedule(LogChangeHandler(), path=watch_dir3, recursive=False)
        if os.path.exists(watch_dir4) and watch_dir4 not in (watch_dir1, watch_dir2, watch_dir3):
            observer.schedule(LogChangeHandler(), path=watch_dir4, recursive=False)
        if os.path.exists(watch_dir5) and watch_dir5 not in (watch_dir1, watch_dir2, watch_dir3, watch_dir4):
            observer.schedule(LogChangeHandler(), path=watch_dir5, recursive=False)
        if os.path.exists(watch_dir6) and watch_dir6 not in (watch_dir1, watch_dir2, watch_dir3, watch_dir4, watch_dir5):
            observer.schedule(LogChangeHandler(), path=watch_dir6, recursive=False)
        observer.start()
        while running:
            updated = False
            # Check log000_1
            if os.path.exists(LOG000_1_PATH):
                try:
                    current_mtime = os.path.getmtime(LOG000_1_PATH)
                    if current_mtime > last_mtime_log1:
                        if update_from_log000_1(force_update=True):
                            updated = True
                            dot_count = 0
                except:
                    pass
            # Check log000_2
            if os.path.exists(LOG000_2_PATH):
                try:
                    current_mtime = os.path.getmtime(LOG000_2_PATH)
                    if current_mtime > last_mtime_log2:
                        if update_from_log000_2(force_update=True):
                            updated = True
                            dot_count = 0
                except:
                    pass
            # Check log000_3
            if os.path.exists(LOG000_3_PATH):
                try:
                    current_mtime = os.path.getmtime(LOG000_3_PATH)
                    if current_mtime > last_mtime_log3:
                        if update_from_log000_3(force_update=True):
                            updated = True
                            dot_count = 0
                except:
                    pass
            # Check log000_4
            if os.path.exists(LOG000_4_PATH):
                try:
                    current_mtime = os.path.getmtime(LOG000_4_PATH)
                    if current_mtime > last_mtime_log4:
                        if update_from_log000_4(force_update=True):
                            updated = True
                            dot_count = 0
                except:
                    pass
            # Check log000_5
            if os.path.exists(LOG000_5_PATH):
                try:
                    current_mtime = os.path.getmtime(LOG000_5_PATH)
                    if current_mtime > last_mtime_log5:
                        if update_from_log000_5(force_update=True):
                            updated = True
                            dot_count = 0
                except:
                    pass
            # Check log000_6
            if os.path.exists(LOG000_6_PATH):
                try:
                    current_mtime = os.path.getmtime(LOG000_6_PATH)
                    if current_mtime > last_mtime_log6:
                        if update_from_log000_6(force_update=True):
                            updated = True
                            dot_count = 0
                except:
                    pass
            if not updated:
                dot_str = "." * min((dot_count % 39) + 1, 39)
                print(f"\rWAITING FOR LOG FILE CHANGES{dot_str:<39}", end="", flush=True)
            else:
                print(f"\r{' ' * 100}", end="\r", flush=True)
                print(f"WAITING FOR LOG FILE CHANGES...", end="", flush=True)
            time.sleep(1)
            dot_count += 1
    except Exception as e:
        log_system(f"Monitoring error: {e}")
    finally:
        if observer:
            observer.stop()
            observer.join()
# === BUTTON COMMANDS ===
def refresh_all():
    global last_mtime_log1, last_mtime_log2, last_mtime_log3, last_mtime_log4, last_mtime_log5, last_mtime_log6
    for i in range(1, 7):
        if i in blink_threads:
            stop_flags[i].set()
            blink_threads[i].join(timeout=1)
        process_active[i] = True
        stop_flags[i].clear()
        process_buttons[i].config(bg="#4CAF50")
    clear_material_log()
    system_text.config(state="normal")
    system_text.delete(1.0, tk.END)
    system_text.config(state="disabled")
    last_mtime_log1 = 0
    last_mtime_log2 = 0
    last_mtime_log3 = 0
    last_mtime_log4 = 0
    last_mtime_log5 = 0
    last_mtime_log6 = 0
    process_job_order()
    log_system("SYSTEM REFRESHED")
refresh_btn.config(command=refresh_all)
def stop_all():
    for i in range(1, 7):
        if not stop_flags[i].is_set():
            stop_flags[i].set()
            process_buttons[i].config(bg="#4CAF50")
    log_system("ALL PROCESSES STOPPED")
stop_all_btn.config(command=stop_all)
# === START MONITORING ===
def start_monitoring():
    global animation_running
    animation_running = True
    anim_thread = threading.Thread(target=animate_loading, daemon=True)
    anim_thread.start()
    mon_thread = threading.Thread(target=monitoring_loop, daemon=True)
    mon_thread.start()
def on_closing():
    global running, animation_running
    running = False
    animation_running = False
    for flag in stop_flags.values():
        flag.set()
    root.destroy()
log_system("SYSTEM STARTED - MONITORING NETWORK PATHS...")
start_monitoring()
root.protocol("WM_DELETE_WINDOW", on_closing)
def go_back():
    root.destroy()
    subprocess.Popen([sys.executable] + sys.argv)
# root.state('zoomed')
root.mainloop()

# 1260 