# AI
# AVG IS NOT ROUND OFF AND IT IS NOT CHANGING WHILE RUNNING
# FROM LTGHis_20.py
# FINAL WORKING
# ERROR: THE AVG AT TKINTER IS CHANGING WHILE RUNNING
# ALREADY RUNNING AT FC1 FOR A MONTH
#  WITH BACK BUTTON (WORKING)

# 1,936
# TRANSFER MAX, MID, MIN MEASUREMENTS BELOW SERIAL NUMBER TEXT BOX (DONE)
# DISPLAY FILE NAME
# 75 ADJUST THE WINDOW HEIGHT OF THE 1ST TKINTER GUI
# WITH LTG HISTORY GOOD
# FIX EXACT 200 ROWS 
# CREATE SAMPLE NUMBER SELECTION (WORKING)
# TRIED TO SWITCH TO HORIZONTAL LAYOUT AND SIZE 5 OF SAMPLE NUMBER
# PUT VERTICAL SCROLLBAR AT DISPLAY SETTINGS

# FOR DATAFRAME GO TO deviation1b.py
# dataframe =   df (original dataframe)
#               model_summary (average) to be vlooked up at compiledFrame               
#               compiledFrame (df) 

# SEARCH
# ROUND OFF

import pandas as pd
import numpy as np
import time
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import logging
from tkinter.ttk import Combobox
import warnings
import threading
import sys

# Suppress FutureWarnings
warnings.filterwarnings('ignore', category=FutureWarning)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# File path  
# file_path = r"\\192.168.2.19\general\INSPECTION-MACHINE\FC1\CompiledPIMachine.csv"  # FC1
file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledPiMachine\CompiledPIMachine.csv"  #ORIGINAL PATH
# file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledPiMachine\CompiledPIMachine.csv"  # FALSE TEST
# output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_CSV.csv"

# Global variables
dataList = []
compiledFrame = pd.DataFrame()
last_modified = None
stop_dots = False

# Function to display "WAITING FOR FILE CHANGES..." with moving dots
def print_waiting_dots():
    global stop_dots
    dots = 0
    while True:
        if stop_dots:
            sys.stdout.write("\r" + " " * 50 + "\r")
            sys.stdout.flush()
            break
        sys.stdout.write("\rWAITING FOR FILE CHANGES" + "." * dots)
        sys.stdout.flush()
        dots = (dots + 1) % 4
        time.sleep(0.5)
    stop_dots = False
    threading.Thread(target=print_waiting_dots, daemon=True).start()

# Start the waiting dots thread
threading.Thread(target=print_waiting_dots, daemon=True).start()

class DatabaseSelection:
    def __init__(self, root):
        self.root = root
        self.root.title("Select Database Location")
        self.root.geometry(f"{int(self.root.winfo_screenwidth() * 0.3)}x{int(self.root.winfo_screenheight() * 1.0)}") # ADJUST THE WINDOW HEIGHT OF THE 1ST TKINTER GUI
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(main_container)
        self.scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        ttk.Label(self.scrollable_frame, text="Choose Database Location:", font=('Arial', 12)).pack(pady=10)
        self.location_var = tk.StringVar(value="FC1-QC")
        locations = [
            ("FC1", "FC1"),
            ("FC1-QC","FC1-QC"),
            ("TESTING", "TESTING")
        ]
        for text, location in locations:
            ttk.Radiobutton(self.scrollable_frame, text=text, variable=self.location_var,
                          value=location).pack(anchor='w', padx=40)
        ttk.Label(self.scrollable_frame, text="TOLERANCE:", font=('Arial', 12)).pack(pady=10)
        self.threshold_var = tk.StringVar(value="5")
        threshold_options = [
            ("3%", "3"),
            ("5%", "5"),
            ("Others", "others")
        ]
        for text, val in threshold_options:
            ttk.Radiobutton(self.scrollable_frame, text=text, variable=self.threshold_var, value=val).pack(anchor='w', padx=40)
        self.other_frame = ttk.Frame(self.scrollable_frame)
        ttk.Label(self.other_frame, text="Enter %:").pack(side=tk.LEFT)
        self.other_entry = ttk.Entry(self.other_frame, width=5)
        self.other_entry.insert(0, "3")
        self.other_entry.pack(side=tk.LEFT)
        self.threshold_var.trace("w", self.on_threshold_change)
        # Add CSV generation radio buttons
        ttk.Label(self.scrollable_frame, text="Generate CSV File:", font=('Arial', 12)).pack(pady=10)
        self.generate_csv_var = tk.StringVar(value="NO")
        csv_options = [
            ("YES", "YES"),
            ("NO", "NO")
        ]
        for text, val in csv_options:
            ttk.Radiobutton(self.scrollable_frame, text=text, variable=self.generate_csv_var, value=val).pack(anchor='w', padx=40)
        ttk.Label(self.scrollable_frame, text="Choose Model Code:", font=('Arial', 12)).pack(pady=10)
        self.model_var = tk.StringVar(value="60CAT0213P")
        model_options = [
            ("60CAT0213P", "60CAT0213P"),
            ("60CAT0212P", "60CAT0212P")
        ]
        for text, val in model_options:
            ttk.Radiobutton(self.scrollable_frame, text=text, variable=self.model_var, value=val).pack(anchor='w', padx=40)
        ttk.Label(self.scrollable_frame, text="Process Inspection Selection:", font=('Arial', 12)).pack(pady=10)
        self.measurement_vars = {}
        measurements = [
            "VOLTAGE MAX (V)",
            "WATTAGE MAX (W)",
            "CLOSED PRESSURE_MAX (kPa)",
            "VOLTAGE Middle (V)",
            "WATTAGE Middle (W)",
            "AMPERAGE Middle (A)",
            "CLOSED PRESSURE Middle (kPa)",
            "VOLTAGE MIN (V)",
            "WATTAGE MIN (W)",
            "CLOSED PRESSURE MIN (kPa)"
        ]
        self.measurement_groups = {
            "max": ["VOLTAGE MAX (V)", "WATTAGE MAX (W)", "CLOSED PRESSURE_MAX (kPa)"],
            "mid": ["VOLTAGE Middle (V)", "WATTAGE Middle (W)", "AMPERAGE Middle (A)", "CLOSED PRESSURE Middle (kPa)"],
            "min": ["VOLTAGE MIN (V)", "WATTAGE MIN (W)", "CLOSED PRESSURE MIN (kPa)"]
        }
        for meas in measurements:
            var = tk.BooleanVar(value=False)
            self.measurement_vars[meas] = var
            ttk.Checkbutton(self.scrollable_frame, text=meas, variable=var).pack(anchor='w', padx=40)
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="MAX", command=lambda: self.select_group("max")).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="MID", command=lambda: self.select_group("mid")).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="MIN", command=lambda: self.select_group("min")).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="SELECT ALL", command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.scrollable_frame, text="Confirm", command=self.confirm_selection).pack(pady=10)
        self.select_group("mid")
    def select_group(self, group):
        for meas, var in self.measurement_vars.items():
            if meas in self.measurement_groups[group]:
                var.set(True)
            else:
                var.set(False)
    def select_all(self):
        for var in self.measurement_vars.values():
            var.set(True)
    def on_threshold_change(self, *args):
        if self.threshold_var.get() == "others":
            self.other_frame.pack(anchor='w', padx=40, pady=5)
            self.other_entry.focus()
            self.root.geometry(f"{self.root.winfo_width()}x{int(self.root.winfo_screenheight() * 0.6)}")
            self.root.update_idletasks()
        else:
            self.other_frame.pack_forget()
            self.root.geometry(f"{self.root.winfo_width()}x{int(self.root.winfo_screenheight() * 0.55)}")
            self.root.update_idletasks()
    def confirm_selection(self):
        location = self.location_var.get()
        threshold_str = self.threshold_var.get()
        generate_csv = self.generate_csv_var.get()
        model = self.model_var.get()
        selected_measurements = [meas for meas, var in self.measurement_vars.items() if var.get()]
        if threshold_str == "others":
            threshold_str = self.other_entry.get().strip()
        try:
            threshold = float(threshold_str)
        except ValueError:
            threshold = 3.0
        if location:
            self.root.destroy()
            root = tk.Tk()
            app = FluctuationMonitor(root, location, threshold, generate_csv, model, selected_measurements)
            root.protocol("WM_DELETE_WINDOW", app.on_closing)
            root.mainloop()
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, callback):
        self.callback = callback
    def on_modified(self, event):
        if event.src_path.endswith('.csv'):
            self.callback()

class FluctuationMonitor:
    def __init__(self, root, location, threshold=3.0, generate_csv="NO", model="60CAT0213P", selected_measurements=None):
        self.root = root
        self.root.title(f"Historical Measurement Trend             FILE NAME: {os.path.basename(__file__)}") # DISPLAY FILE NAME
        self.root.geometry(f"{int(self.root.winfo_screenwidth() * 0.8)}x{int(self.root.winfo_screenheight() * 0.8)}")
        self.root.state('zoomed')
        self.zoom_level = 1.0
        self.threshold = threshold / 100
        self.generate_csv = generate_csv
        self.x_axis_mode = tk.StringVar(value="Numerical")
        self.layout_mode = tk.StringVar(value="VERTICAL")
        self.show_title_var = tk.BooleanVar(value=False)
        self.num_ticks = 10
        self.tick_strategy = "full"
        self.line_width_var = tk.StringVar(value="MEDIUM")
        self.line_width_map = {
            "EXTRA SMALL": 0.5,
            "SMALL": 1.0,
            "MEDIUM": 2.0,
            "LARGE": 3.0,
            "EXTRA LARGE": 4.0
        }
        self.main_container = ttk.Frame(root)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(self.main_container)
        self.scrollbar_y = ttk.Scrollbar(self.main_container, orient="vertical", command=self.canvas.yview)
        self.scrollbar_x = ttk.Scrollbar(self.main_container, orient="horizontal", command=self.canvas.xview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.pack(side="top", fill="both", expand=True)
        self.scrollbar_y.pack(side="right", fill="y")
        self.scrollbar_x.pack(side="bottom", fill="x")
        zoom_frame = ttk.Frame(self.scrollable_frame)
        zoom_frame.pack(pady=2, anchor='ne')
        ttk.Button(zoom_frame, text="Zoom In (+)", command=lambda: self.zoom(1.1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="Zoom Out (-)", command=lambda: self.zoom(0.9)).pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="Reset Zoom", command=lambda: self.zoom(1.0)).pack(side=tk.LEFT, padx=2)
        self.status_vars = {}
        self.status_labels = {}
        self.current_model = model
        self.last_date = None
        self.last_good_values = {}
        self.last_good_serial = None
        self.fluctuation_count = 0
        self.last_row_count = 0
        self.all_measurements = [
            "VOLTAGE MAX (V)", "WATTAGE MAX (W)", "CLOSED PRESSURE_MAX (kPa)",
            "VOLTAGE Middle (V)", "WATTAGE Middle (W)", "AMPERAGE Middle (A)", "CLOSED PRESSURE Middle (kPa)",
            "VOLTAGE MIN (V)", "WATTAGE MIN (W)", "CLOSED PRESSURE MIN (kPa)"
        ]
        self.selected_measurements = selected_measurements or self.all_measurements
        self.measurements = self.selected_measurements[:]
        self.section_map = {
            "MAX Measurements": ["VOLTAGE MAX (V)", "WATTAGE MAX (W)", "CLOSED PRESSURE_MAX (kPa)"],
            "MID Measurements": ["VOLTAGE Middle (V)", "WATTAGE Middle (W)", "AMPERAGE Middle (A)", "CLOSED PRESSURE Middle (kPa)"],
            "MIN Measurements": ["VOLTAGE MIN (V)", "WATTAGE MIN (W)", "CLOSED PRESSURE MIN (kPa)"]
        }
        self.short_map = {
            "VOLTAGE MAX (V)": 'V_MAX',
            "WATTAGE MAX (W)": 'W_MAX',
            "CLOSED PRESSURE_MAX (kPa)": 'CP_MAX',
            "VOLTAGE Middle (V)": 'V_MID',
            "WATTAGE Middle (W)": 'W_MID',
            "AMPERAGE Middle (A)": 'A_MID',
            "CLOSED PRESSURE Middle (kPa)": 'CP_MID',
            "VOLTAGE MIN (V)": 'V_MIN',
            "WATTAGE MIN (W)": 'W_MIN',
            "CLOSED PRESSURE MIN (kPa)": 'CP_MIN'
        }
        self.dev_map = {
            'V_MAX': 'DEV V_MAX PASS',
            'W_MAX': 'DEV WATTAGE MAX (W)',
            'CP_MAX': 'DEV CLOSED PRESSURE_MAX (kPa)',
            'V_MID': 'DEV VOLTAGE Middle (V)',
            'W_MID': 'DEV WATTAGE Middle (W)',
            'A_MID': 'DEV AMPERAGE Middle (A)',
            'CP_MID': 'DEV CLOSED PRESSURE Middle (kPa)',
            'V_MIN': 'DEV VOLTAGE MIN (V)',
            'W_MIN': 'DEV WATTAGE MIN (W)',
            'CP_MIN': 'DEV CLOSED PRESSURE MIN (kPa)'
        }
        self.fluctuation_measurements = [self.short_map[m] for m in self.selected_measurements]
        self.measurements_map = {
            "VOLTAGE MAX (V)": ("VOLTAGE MAX (V)", "AVE V_MAX PASS", "DEV V_MAX PASS", "DEV V_MAX PASS FLUCTUATED"),
            "WATTAGE MAX (W)": ("WATTAGE MAX (W)", "AVE WATTAGE MAX (W)", "DEV WATTAGE MAX (W)", "DEV WATTAGE MAX (W) FLUCTUATED"),
            "CLOSED PRESSURE_MAX (kPa)": ("CLOSED PRESSURE_MAX (kPa)", "AVE CLOSED PRESSURE_MAX (kPa)", "DEV CLOSED PRESSURE_MAX (kPa)", "DEV CLOSED PRESSURE_MAX (kPa) FLUCTUATED"),
            "VOLTAGE Middle (V)": ("VOLTAGE Middle (V)", "AVE VOLTAGE Middle (V)", "DEV VOLTAGE Middle (V)", "DEV VOLTAGE Middle (V) FLUCTUATED"),
            "WATTAGE Middle (W)": ("WATTAGE Middle (W)", "AVE WATTAGE Middle (W)", "DEV WATTAGE Middle (W)", "DEV WATTAGE Middle (W) FLUCTUATED"),
            "AMPERAGE Middle (A)": ("AMPERAGE Middle (A)", "AVE AMPERAGE Middle (A)", "DEV AMPERAGE Middle (A)", "DEV AMPERAGE Middle (A) FLUCTUATED"),
            "CLOSED PRESSURE Middle (kPa)": ("CLOSED PRESSURE Middle (kPa)", "AVE CLOSED PRESSURE Middle (kPa)", "DEV CLOSED PRESSURE Middle (kPa)", "DEV CLOSED PRESSURE Middle (kPa) FLUCTUATED"),
            "VOLTAGE MIN (V)": ("VOLTAGE MIN (V)", "AVE VOLTAGE MIN (V)", "DEV VOLTAGE MIN (V)", "DEV VOLTAGE MIN (V) FLUCTUATED"),
            "WATTAGE MIN (W)": ("WATTAGE MIN (W)", "AVE WATTAGE MIN (W)", "DEV WATTAGE MIN (W)", "DEV WATTAGE MIN (W) FLUCTUATED"),
            "CLOSED PRESSURE MIN (kPa)": ("CLOSED PRESSURE MIN (kPa)", "AVE CLOSED PRESSURE MIN (kPa)", "DEV CLOSED PRESSURE MIN (kPa)", "DEV CLOSED PRESSURE MIN (kPa) FLUCTUATED"),
        }
        self.avgs_map = {
            "VOLTAGE MAX (V)": 'V-MAX PASS AVG',
            "WATTAGE MAX (W)": 'WATTAGE MAX AVG',
            "CLOSED PRESSURE_MAX (kPa)": 'CLOSED PRESSURE_MAX AVG',
            "VOLTAGE Middle (V)": 'VOLTAGE Middle AVG',
            "WATTAGE Middle (W)": 'WATTAGE Middle AVG',
            "AMPERAGE Middle (A)": 'AMPERAGE Middle AVG',
            "CLOSED PRESSURE Middle (kPa)": 'CLOSED PRESSURE Middle AVG',
            "VOLTAGE MIN (V)": 'VOLTAGE MIN (V) AVG',
            "WATTAGE MIN (W)": 'WATTAGE MIN AVG',
            "CLOSED PRESSURE MIN (kPa)": 'CLOSED PRESSURE MIN AVG',
        }
        self.devs_map = {
            "VOLTAGE MAX (V)": 'DEV V_MAX PASS',
            "WATTAGE MAX (W)": 'DEV WATTAGE MAX (W)',
            "CLOSED PRESSURE_MAX (kPa)": 'DEV CLOSED PRESSURE_MAX (kPa)',
            "VOLTAGE Middle (V)": 'DEV VOLTAGE Middle (V)',
            "WATTAGE Middle (W)": 'DEV WATTAGE Middle (W)',
            "AMPERAGE Middle (A)": 'DEV AMPERAGE Middle (A)',
            "CLOSED PRESSURE Middle (kPa)": 'DEV CLOSED PRESSURE Middle (kPa)',
            "VOLTAGE MIN (V)": 'DEV VOLTAGE MIN (V)',
            "WATTAGE MIN (W)": 'DEV WATTAGE MIN (W)',
            "CLOSED PRESSURE MIN (kPa)": 'DEV CLOSED PRESSURE MIN (kPa)',
        }
        if location == "FC1":
            self.file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledPiMachine\CompiledPIMachine.csv" # ORIGINAL PATH
            self.output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_CSV.csv"
            self.log_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_Log.txt"
            # self.file_path = r"C:\Users\ai.pc\OneDrive\Desktop\CompiledPIMachine.csv"  # FALSE TEST
        elif location == "FC1-QC":
            self.file_path = r"\\192.168.2.19\general\INSPECTION-MACHINE\FC1\CompiledPIMachine.csv"
            self.output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_CSV.csv"
            self.log_path = r"\\192.168.2.19\general\INSPECTION-MACHINE\FC1\Deviation_Log.txt"
        elif location == "TESTING":
            self.file_path = r"C:\Users\ai.pc\OneDrive\Desktop\CompiledPIMachine.csv"
            self.output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_CSV.csv"
            self.log_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_Log.txt"
        else:
            self.file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledPiMachine\CompiledPIMachine.csv"
            self.output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_CSV.csv"
            self.log_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_Log.txt"
        self.df_columns = None
        emptyColumn = [
            "DATE", "TIME", "MODEL CODE", "S/N", "PASS/NG",
            "VOLTAGE MAX (V)", "V_MAX PASS", "AVE V_MAX PASS", "DEV V_MAX PASS", "DEV V_MAX PASS FLUCTUATED",
            "WATTAGE MAX (W)", "WATTAGE MAX PASS", "AVE WATTAGE MAX (W)", "DEV WATTAGE MAX (W)", "DEV WATTAGE MAX (W) FLUCTUATED",
            "CLOSED PRESSURE_MAX (kPa)", "CLOSED PRESSURE_MAX PASS", "AVE CLOSED PRESSURE_MAX (kPa)", "DEV CLOSED PRESSURE_MAX (kPa)", "DEV CLOSED PRESSURE_MAX (kPa) FLUCTUATED",
            "VOLTAGE Middle (V)", "VOLTAGE Middle PASS", "AVE VOLTAGE Middle (V)", "DEV VOLTAGE Middle (V)", "DEV VOLTAGE Middle (V) FLUCTUATED",
            "WATTAGE Middle (W)", "WATTAGE Middle (W) PASS", "AVE WATTAGE Middle (W)", "DEV WATTAGE Middle (W)", "DEV WATTAGE Middle (W) FLUCTUATED",
            "AMPERAGE Middle (A)", "AMPERAGE Middle (A) PASS", "AVE AMPERAGE Middle (A)", "DEV AMPERAGE Middle (A)", "DEV AMPERAGE Middle (A) FLUCTUATED",
            "CLOSED PRESSURE Middle (kPa)", "CLOSED PRESSURE Middle (kPa) PASS", "AVE CLOSED PRESSURE Middle (kPa)", "DEV CLOSED PRESSURE Middle (kPa)", "DEV CLOSED PRESSURE Middle (kPa) FLUCTUATED",
            "VOLTAGE MIN (V)", "VOLTAGE MIN (V) PASS", "AVE VOLTAGE MIN (V)", "DEV VOLTAGE MIN (V)", "DEV VOLTAGE MIN (V) FLUCTUATED",
            "WATTAGE MIN (W)", "WATTAGE MIN (W) PASS", "AVE WATTAGE MIN (W)", "DEV WATTAGE MIN (W)", "DEV WATTAGE MIN (W) FLUCTUATED",
            "CLOSED PRESSURE MIN (kPa)", "CLOSED PRESSURE MIN (kPa) PASS", "AVE CLOSED PRESSURE MIN (kPa)", "DEV CLOSED PRESSURE MIN (kPa)", "DEV CLOSED PRESSURE MIN (kPa) FLUCTUATED",
            "REFERENCE SERIAL", "DATETIME"
        ]
        dtypes = {
            "DATE": str,
            "TIME": str,
            "MODEL CODE": str,
            "S/N": str,
            "PASS/NG": float,
            "VOLTAGE MAX (V)": float,
            "V_MAX PASS": float,
            "AVE V_MAX PASS": float,
            "DEV V_MAX PASS": float,
            "DEV V_MAX PASS FLUCTUATED": float,
            "WATTAGE MAX (W)": float,
            "WATTAGE MAX PASS": float,
            "AVE WATTAGE MAX (W)": float,
            "DEV WATTAGE MAX (W)": float,
            "DEV WATTAGE MAX (W) FLUCTUATED": float,
            "CLOSED PRESSURE_MAX (kPa)": float,
            "CLOSED PRESSURE_MAX PASS": float,
            "AVE CLOSED PRESSURE_MAX (kPa)": float,
            "DEV CLOSED PRESSURE_MAX (kPa)": float,
            "DEV CLOSED PRESSURE_MAX (kPa) FLUCTUATED": float,
            "VOLTAGE Middle (V)": float,
            "VOLTAGE Middle PASS": float,
            "AVE VOLTAGE Middle (V)": float,
            "DEV VOLTAGE Middle (V)": float,
            "DEV VOLTAGE Middle (V) FLUCTUATED": float,
            "WATTAGE Middle (W)": float,
            "WATTAGE Middle (W) PASS": float,
            "AVE WATTAGE Middle (W)": float,
            "DEV WATTAGE Middle (W)": float,
            "DEV WATTAGE Middle (W) FLUCTUATED": float,
            "AMPERAGE Middle (A)": float,
            "AMPERAGE Middle (A) PASS": float,
            "AVE AMPERAGE Middle (A)": float,
            "DEV AMPERAGE Middle (A)": float,
            "DEV AMPERAGE Middle (A) FLUCTUATED": float,
            "CLOSED PRESSURE Middle (kPa)": float,
            "CLOSED PRESSURE Middle (kPa) PASS": float,
            "AVE CLOSED PRESSURE Middle (kPa)": float,
            "DEV CLOSED PRESSURE Middle (kPa)": float,
            "DEV CLOSED PRESSURE Middle (kPa) FLUCTUATED": float,
            "VOLTAGE MIN (V)": float,
            "VOLTAGE MIN (V) PASS": float,
            "AVE VOLTAGE MIN (V)": float,
            "DEV VOLTAGE MIN (V)": float,
            "DEV VOLTAGE MIN (V) FLUCTUATED": float,
            "WATTAGE MIN (W)": float,
            "WATTAGE MIN (W) PASS": float,
            "AVE WATTAGE MIN (W)": float,
            "DEV WATTAGE MIN (W)": float,
            "DEV WATTAGE MIN (W) FLUCTUATED": float,
            "CLOSED PRESSURE MIN (kPa)": float,
            "CLOSED PRESSURE MIN (kPa) PASS": float,
            "AVE CLOSED PRESSURE MIN (kPa)": float,
            "DEV CLOSED PRESSURE MIN (kPa)": float,
            "DEV CLOSED PRESSURE MIN (kPa) FLUCTUATED": float,
            "REFERENCE SERIAL": str,
            "DATETIME": str
        }
        if self.generate_csv == "YES" and os.path.exists(self.output_path) and False:
            try:
                self.compiledFrame = pd.read_csv(self.output_path, encoding='utf-8-sig', dtype=dtypes, on_bad_lines='skip')
                self.compiledFrame['DATETIME'] = pd.to_datetime(self.compiledFrame['DATETIME'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
                self.compiledFrame = self.compiledFrame[self.compiledFrame['PASS/NG'] == 1].reset_index(drop=True)
            except Exception as e:
                logging.error(f"Error loading existing CSV: {e}")
                self.compiled_frame = pd.DataFrame(columns=emptyColumn)
        else:
            self.compiledFrame = pd.DataFrame(columns=emptyColumn)
        self.main_frame = ttk.Frame(self.scrollable_frame)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X)
        self.left_frame = ttk.Frame(self.top_frame)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Button(self.left_frame, text="<- Back", command=self.go_back).pack(anchor='nw', pady=5)
        ttk.Label(self.left_frame, text="Deviation Status Monitor", 
                 font=('Arial', 14, 'bold')).pack(pady=5)
        ttk.Label(self.left_frame, text=f"Location: {location}", 
                 font=('Arial', 12)).pack(pady=2)
        self.tolerance_display = ttk.Label(self.left_frame, text=f"Tolerance: {self.threshold * 100:.1f}%", 
                 font=('Arial', 12))
        self.tolerance_display.pack(pady=2)
        self.model_display = tk.Label(self.left_frame, text="MODEL CODE: N/A", 
                                    font=('Arial', 12), fg="purple")
        self.model_display.pack(pady=2)
        self.serial_display = tk.Label(self.left_frame, text="SERIAL No.", 
                                      font=('Arial', 16, 'bold'), fg="blue")
        self.serial_display.pack(pady=2)
        self.status_box = tk.Canvas(self.left_frame, width=300, height=150,
                                   bg="gray", highlightthickness=2, 
                                   highlightbackground="black")
        self.status_box.pack(pady=10)
        self.status_text = self.status_box.create_text(150, 50, text="NO FLUCTUATION DETECTED", 
                                                     font=('Arial', 14, 'bold'), fill="white", anchor="center")
        self.counter_text = self.status_box.create_text(150, 100, text="Fluctuations: 0/10", 
                                                      font=('Arial', 10), fill="white", anchor="center")
        self.serial_log = tk.Text(self.left_frame, height=4, width=40)
        self.serial_log.pack(pady=5)
        self.serial_log.insert(tk.END, "Serial Number Log:\n")
        self.serial_log.config(state='disabled')
        for title in ["MAX Measurements", "MID Measurements", "MIN Measurements"]:
            selected_in_sec = [m for m in self.section_map[title] if m in self.selected_measurements]
            if selected_in_sec:
                fluct_cols = [self.measurements_map[m][3] for m in selected_in_sec]
                self.create_section(title, fluct_cols)
                ttk.Separator(self.left_frame, orient='horizontal').pack(fill='x', pady=5)
        button_frame_min = ttk.Frame(self.left_frame)
        button_frame_min.pack(pady=5, fill=tk.X)
        inner_frame = ttk.Frame(button_frame_min)
        inner_frame.pack(anchor='center')
        ttk.Button(inner_frame, text="Refresh Values", command=self.process_and_update).pack(side=tk.LEFT, padx=2)
        ttk.Button(inner_frame, text="Reset All", command=self.reset_all_fluctuations).pack(side=tk.LEFT, padx=2)
        self.details_frame = ttk.Frame(self.top_frame)
        self.details_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        self.right_frame = ttk.Frame(self.top_frame)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10)
        #self.create_bar_graph()
        self.spc_graph_frame = ttk.Frame(self.details_frame)
        self.spc_graph_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.line_graph_frame = ttk.Frame(self.details_frame)
        #self.line_graph_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        #self.create_line_graph()
        self.mid_measurements = ["VOLTAGE Middle (V)", "WATTAGE Middle (W)", "AMPERAGE Middle (A)", "CLOSED PRESSURE Middle (kPa)"]
        self.figs = {}
        self.axs = {}
        self.canvases = {}
        self.ucl_labels = {}
        self.lcl_labels = {}
        self.avg_labels = {}
        self.start_labels = {}
        self.end_labels = {}
        self.message_labels = {}
        self.last_labels = {}
        self.rebuild_spc_graphs()
        self.combo_frame = ttk.Frame(self.main_frame)
        self.combo_frame.pack(pady=5)
        ttk.Label(self.combo_frame, text="Select Model Code:").pack(side=tk.LEFT)
        self.combo = Combobox(self.combo_frame)
        self.combo.pack(side=tk.LEFT)
        self.combo.bind("<<ComboboxSelected>>", self.on_select)
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=5)
        ttk.Button(button_frame, text="Focus Trends", command=self.open_focus_window).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Display Settings", command=self.open_display_window).pack(side=tk.LEFT, padx=2)
        # ttk.Button(button_frame, text="Reset Log", command=self.reset_log).pack(side=tk.LEFT, padx=2)   # DISABLE "LOG RESET" BUTTON
        ttk.Button(button_frame, text="X-Axis", command=self.open_ticks_window).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="COMPUTATION", command=self.open_computation_window).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="LTG HISTORY", command=self.open_ltg_history_window).pack(side=tk.LEFT, padx=2)
        self.log_frame = ttk.Frame(self.main_frame)
        self.log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        ttk.Label(self.log_frame, text="Fluctuation Log", font=('Arial', 14, 'bold')).pack(pady=5)
        self.log_text = tk.Text(self.log_frame, wrap='word')
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar = ttk.Scrollbar(self.log_frame, orient='vertical', command=self.log_text.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        self.last_modified = os.path.getmtime(self.file_path)
        event_handler = FileChangeHandler(self.check_file_update)
        self.observer = Observer()
        self.observer.schedule(event_handler, os.path.dirname(self.file_path))
        self.observer.start()
        self.after_id = None
        self.process_and_update()
        self.after_id = self.root.after(1000, self.periodic_check)
        self.old_width = None
        self.old_height = None
        self.root.bind("<Configure>", self.on_resize)
        self.root.update_idletasks()
        fake_event = type('Event', (), {'width': self.root.winfo_width(), 'height': self.root.winfo_screenheight() * 0.8, 'widget': self.root})()
        self.on_resize(fake_event)
    def go_back(self):
        if self.after_id is not None:
            self.root.after_cancel(self.after_id)
        self.observer.stop()
        self.observer.join()
        self.root.destroy()
        new_root = tk.Tk()
        DatabaseSelection(new_root)
        new_root.mainloop()

    def rebuild_spc_graphs(self):
        for widget in self.spc_graph_frame.winfo_children():
            widget.destroy()
        self.figs = {}
        self.axs = {}
        self.canvases = {}
        self.ucl_labels = {}
        self.lcl_labels = {}
        self.avg_labels = {}
        self.start_labels = {}
        self.end_labels = {}
        self.message_labels = {}
        self.last_labels = {}
        if self.layout_mode.get() == "HORIZONTAL":
            row1 = ttk.Frame(self.spc_graph_frame)
            row1.pack(fill=tk.BOTH, expand=True, pady=5)
            row2 = ttk.Frame(self.spc_graph_frame)
            row2.pack(fill=tk.BOTH, expand=True, pady=5)
            row_frames = [row1, row2]
            meas_chunks = [self.mid_measurements[:2], self.mid_measurements[2:]]
            for i, row_meas in enumerate(meas_chunks):
                row_frame = row_frames[i]
                for meas in row_meas:
                    sub_frame = ttk.Frame(row_frame)
                    sub_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
                    graph_sub = ttk.Frame(sub_frame)
                    graph_sub.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                    fig = Figure(figsize=(10 * self.num_ticks / 10, 3))
                    ax = fig.add_subplot(111)
                    canvas = FigureCanvasTkAgg(fig, master=graph_sub)
                    canvas.draw()
                    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                    self.figs[meas] = fig
                    self.axs[meas] = ax
                    self.canvases[meas] = canvas
                    right_sub = ttk.Frame(sub_frame)
                    right_sub.pack(side=tk.RIGHT, fill=tk.Y, padx=10)
                    ttk.Label(right_sub, text="UCL:").pack(anchor='e')
                    ucl_lab = ttk.Label(right_sub, text="N/A")
                    ucl_lab.pack(anchor='w')
                    self.ucl_labels[meas] = ucl_lab
                    ttk.Label(right_sub, text="LCL:").pack(anchor='e', pady=(10,0))
                    lcl_lab = ttk.Label(right_sub, text="N/A")
                    lcl_lab.pack(anchor='w')
                    self.lcl_labels[meas] = lcl_lab
                    ttk.Label(right_sub, text="Avg:").pack(anchor='e', pady=(10,0))
                    avg_lab = ttk.Label(right_sub, text="N/A")
                    avg_lab.pack(anchor='w')
                    self.avg_labels[meas] = avg_lab
                    ttk.Label(right_sub, text="Last:").pack(anchor='e', pady=(10,0))
                    last_lab = ttk.Label(right_sub, text="N/A")
                    last_lab.pack(anchor='w')
                    self.last_labels[meas] = last_lab
                    ttk.Label(right_sub, text="Start:").pack(anchor='e', pady=(10,0))
                    start_lab = ttk.Label(right_sub, text="N/A")
                    start_lab.pack(anchor='w')
                    self.start_labels[meas] = start_lab
                    ttk.Label(right_sub, text="End:").pack(anchor='e', pady=(10,0))
                    end_lab = ttk.Label(right_sub, text="N/A")
                    end_lab.pack(anchor='w')
                    self.end_labels[meas] = end_lab
                    message_lab = ttk.Label(right_sub, text="", font=('Arial', 8, 'bold'), foreground="red")
                    message_lab.pack(anchor='w', pady=(10,0))
                    self.message_labels[meas] = message_lab
        else:
            for meas in self.mid_measurements:
                sub_frame = ttk.Frame(self.spc_graph_frame)
                sub_frame.pack(fill=tk.BOTH, expand=True, pady=5)
                graph_sub = ttk.Frame(sub_frame)
                graph_sub.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                fig = Figure(figsize=(10 * self.num_ticks / 10, 3))
                ax = fig.add_subplot(111)
                canvas = FigureCanvasTkAgg(fig, master=graph_sub)
                canvas.draw()
                canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                self.figs[meas] = fig
                self.axs[meas] = ax
                self.canvases[meas] = canvas
                right_sub = ttk.Frame(sub_frame)
                right_sub.pack(side=tk.RIGHT, fill=tk.Y, padx=10)
                ttk.Label(right_sub, text="UCL:").pack(anchor='e')
                ucl_lab = ttk.Label(right_sub, text="N/A")
                ucl_lab.pack(anchor='w')
                self.ucl_labels[meas] = ucl_lab
                ttk.Label(right_sub, text="LCL:").pack(anchor='e', pady=(10,0))
                lcl_lab = ttk.Label(right_sub, text="N/A")
                lcl_lab.pack(anchor='w')
                self.lcl_labels[meas] = lcl_lab
                ttk.Label(right_sub, text="Avg:").pack(anchor='e', pady=(10,0))
                avg_lab = ttk.Label(right_sub, text="N/A")
                avg_lab.pack(anchor='w')
                self.avg_labels[meas] = avg_lab
                ttk.Label(right_sub, text="Last:").pack(anchor='e', pady=(10,0))
                last_lab = ttk.Label(right_sub, text="N/A")
                last_lab.pack(anchor='w')
                self.last_labels[meas] = last_lab
                ttk.Label(right_sub, text="Start:").pack(anchor='e', pady=(10,0))
                start_lab = ttk.Label(right_sub, text="N/A")
                start_lab.pack(anchor='w')
                self.start_labels[meas] = start_lab
                ttk.Label(right_sub, text="End:").pack(anchor='e', pady=(10,0))
                end_lab = ttk.Label(right_sub, text="N/A")
                end_lab.pack(anchor='w')
                self.end_labels[meas] = end_lab
                message_lab = ttk.Label(right_sub, text="", font=('Arial', 8, 'bold'), foreground="red")
                message_lab.pack(anchor='w', pady=(10,0))
                self.message_labels[meas] = message_lab

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def zoom(self, factor):
        if factor != 1.0:
            self.zoom_level *= factor
        else:
            self.zoom_level = 1.0
        self.zoom_level = max(0.5, min(self.zoom_level, 2.0))
        widgets = [
            (self.serial_display, 16),
            (self.model_display, 12),
        ]
        for widget, base_size in widgets:
            current_font = widget.cget("font")
            font_name = 'Arial'
            if isinstance(current_font, str):
                font_name = current_font.split()[0]
            new_size = int(base_size * self.zoom_level)
            widget.config(font=(font_name, new_size, 'bold' if 'bold' in str(current_font) else ''))
        new_width = int(300 * self.zoom_level)
        new_height = int(150 * self.zoom_level)
        self.status_box.config(width=new_width, height=new_height)
        status_y = int(50 * self.zoom_level)
        counter_y = int(100 * self.zoom_level)
        min_spacing = int(30 * self.zoom_level)
        if counter_y - status_y < min_spacing:
            counter_y = status_y + min_spacing
        current_font = self.status_box.itemcget(self.status_text, "font")
        font_parts = current_font.split()
        font_name = 'Arial'
        if len(font_parts) > 0:
            font_name = font_parts[0]
        new_size = int(14 * self.zoom_level)
        font_weight = 'bold' if 'bold' in current_font else ''
        self.status_box.itemconfig(self.status_text, font=(font_name, new_size, font_weight))
        self.status_box.coords(self.status_text, new_width / 2, status_y)
        current_font = self.status_box.itemcget(self.counter_text, "font")
        font_parts = current_font.split()
        font_name = 'Arial'
        if len(font_parts) > 0:
            font_name = font_parts[0]
        new_size = int(10 * self.zoom_level)
        font_weight = '' if 'bold' in current_font else ''
        self.status_box.itemconfig(self.counter_text, font=(font_name, new_size, font_weight))
        self.status_box.coords(self.counter_text, new_width / 2, counter_y)
        for column in self.status_vars:
            new_size = int(10 * self.zoom_level)
            self.status_labels[column].config(font=('Arial', new_size, 'bold'))
        for frame in self.details_frame.winfo_children():
            for widget in frame.winfo_children():
                if isinstance(widget, ttk.Label) and widget.cget('text') in ["MAX Measurements", "MID Measurements", "MIN Measurements"]:
                    new_size = int(12 * self.zoom_level)
                    widget.config(font=('Arial', new_size, 'bold'))
        for meas in self.mid_measurements:
            self.figs[meas].set_size_inches(10 * self.zoom_level, 3 * self.zoom_level)
            self.canvases[meas].draw()


    def update_spc_graphs(self):
        global compiledFrame
        try:
            if compiledFrame.empty:
                return
            df = compiledFrame[compiledFrame["PASS/NG"] == 1].copy()
            if self.current_model:
                df = df[df['MODEL CODE'] == self.current_model].copy()
            df['DATE'] = pd.to_datetime(df['DATE'], format='%Y/%m/%d', errors='coerce')
            df['DATETIME'] = pd.to_datetime(df['DATETIME'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
            df = df.sort_values('DATETIME')
            today = pd.to_datetime(datetime.now().date())
            if df.empty:
                return
            model_row = self.model_summary[self.model_summary['MODEL CODE'] == self.current_model].iloc[0] if hasattr(self, 'model_summary') and not self.model_summary.empty else None
            start_d = model_row['START_DATE'] if model_row is not None else 'N/A'
            end_d = model_row['END_DATE'] if model_row is not None else 'N/A'
            start_str = start_d.strftime('%Y-%m-%d') if pd.notnull(start_d) and not isinstance(start_d, str) else str(start_d) if start_d != 'N/A' else 'N/A'
            end_str = end_d.strftime('%Y-%m-%d') if pd.notnull(end_d) and not isinstance(end_d, str) else str(end_d) if end_d != 'N/A' else 'N/A'
            if 'END_SERIAL' in model_row:
                end_str += f" ({model_row['END_SERIAL']})"
            line_width = self.line_width_map[self.line_width_var.get()]
            df_prev = df[df['DATE'].dt.date < today.date()]
            df_plot = df.tail(200)
            for meas in self.mid_measurements:
                full_data_prev = df_prev[meas].tail(200)
                mean_val = full_data_prev.mean()
                self.avg_labels[meas].config(text=f"{mean_val:.6f}")        #ROUND OFF self.avg_labels[meas].config(text=str(mean_val))
                self.start_labels[meas].config(text=start_str)
                self.end_labels[meas].config(text=end_str)
                upper_limit = mean_val * (1 + self.threshold)
                lower_limit = mean_val * (1 - self.threshold)
                ax = self.axs[meas]
                ax.clear()
                full_data = df_plot[meas]
                data = full_data.copy()
                if self.tick_strategy == "last":
                    tick_rel_pos = []
                    tail_size = len(df_plot)
                    if self.num_ticks == 3:
                        tick_rel_pos = [0,22,45]
                        tail_size = 46
                    elif self.num_ticks == 5:
                        tick_rel_pos = [0,22,44,66,89]
                        tail_size = 90
                    if tail_size > len(df_plot):
                        tail_size = len(df_plot)
                        tick_rel_pos = np.linspace(0, tail_size - 1, self.num_ticks, dtype=int)
                    full_n = len(df_plot)
                    start_sample = full_n - tail_size + 1
                    data = full_data.tail(tail_size)
                    if self.x_axis_mode.get() == "Numerical":
                        tick_labels = [str(start_sample + p) for p in tick_rel_pos]
                    elif self.x_axis_mode.get() == "DateTime":
                        tick_labels = df_plot['DATETIME'].tail(tail_size).iloc[tick_rel_pos].dt.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        tick_labels = ['' for _ in tick_rel_pos]
                    tick_positions = tick_rel_pos
                else:
                    data = full_data
                    tick_positions = np.linspace(0, len(data)-1, self.num_ticks, dtype=int)
                    if self.x_axis_mode.get() == "Numerical":
                        tick_labels = [str(pos + 1) for pos in tick_positions]
                    elif self.x_axis_mode.get() == "DateTime":
                        tick_labels = df_plot['DATETIME'].iloc[tick_positions].dt.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        tick_labels = ['' for _ in tick_positions]
                    tick_positions = np.linspace(0, len(data)-1, self.num_ticks, dtype=int)
                x = np.arange(len(data))
                last_row = df_plot.iloc[-1]
                fluct_col = self.measurements_map[meas][3]
                is_fluctuated = df_plot[fluct_col].iloc[-1] > self.threshold if fluct_col in df_plot.columns else False
                last_color = 'lightcoral' if is_fluctuated else 'green'
                n = len(data)
                if n > 1:
                    ax.plot(x[:-1], data.values[:-1], marker='o', linestyle='-', color='blue', linewidth=line_width)
                    ax.plot(x[-2:], data.values[-2:], marker='o', linestyle='--', color=last_color, linewidth=line_width)
                else:
                    ax.plot(x, data.values, marker='o', linestyle='-', color='blue', linewidth=line_width)
                last_x = len(data) - 1
                last_y = data.iloc[-1]
                circle_color = 'red' if is_fluctuated else 'none'
                ax.scatter(last_x, last_y, color=circle_color, s=100, marker='o', zorder=5, linewidth=2, edgecolors='black' if is_fluctuated else 'none')
                ax.axhline(mean_val, color='green', linestyle='--', linewidth=line_width)
                ax.axhline(upper_limit, color='red', linestyle='--', linewidth=line_width)
                ax.axhline(lower_limit, color='red', linestyle='--', linewidth=line_width)
                out_mask = (data > upper_limit) | (data < lower_limit)
                if out_mask.any():
                    out_x = np.where(out_mask)[0]
                    ax.scatter(out_x, data[out_mask].values, color='red', s=100, zorder=5)
                if self.show_title_var.get():
                    ax.set_title(meas)
                else:
                    ax.set_title('')
                if self.x_axis_mode.get() == "Numerical":
                    ax.set_xlabel('Sample Number')
                elif self.x_axis_mode.get() == "DateTime":
                    ax.set_xlabel('DateTime')
                else:
                    ax.set_xlabel('')
                    ax.set_xticks([])
                    ax.set_xticklabels([])
                ax.set_ylabel(meas)
                ax.grid(True, linestyle='--', alpha=0.7)
                ax.set_xticks(tick_positions)
                ax.set_xticklabels(tick_labels)
                
                # LINE TREND MARGIN (left=0.10, right=0.95, bottom=0.1, top=0.95)
                self.figs[meas].subplots_adjust(left=0.10, right=0.95, bottom=0.1, top=0.95)
                self.canvases[meas].draw()
                self.ucl_labels[meas].config(text=f"{upper_limit:.6f}")     #ROUND OFF .6f (6 DECIMAL PLACES)
                self.lcl_labels[meas].config(text=f"{lower_limit:.6f}")     #ROUND OFF .6f (6 DECIMAL PLACES)
                self.message_labels[meas].config(text="FLUCTUATED" if is_fluctuated else "")
                self.last_labels[meas].config(text=f"{last_y:.2f}")
        except Exception as e:
            logging.error(f"Error updating SPC graphs: {e}")


    def update_status_box(self, has_fluctuation, count=0):
        if has_fluctuation:
            self.status_box.configure(bg="red")
            self.status_box.itemconfig(self.status_text, text="FLUCTUATION DETECTED!")
            self.status_box.itemconfig(self.counter_text, text=f"Fluctuations: {count}/{len(self.selected_measurements)}")
        else:
            self.status_box.configure(bg="green")
            self.status_box.itemconfig(self.status_text, text="NO FLUCTUATION DETECTED")
            self.status_box.itemconfig(self.counter_text, text=f"Fluctuations: 0/{len(self.selected_measurements)}")

    def periodic_check(self):
        self.check_file_update()
        self.after_id = self.root.after(1000, self.periodic_check)

    def check_file_update(self):
        global stop_dots
        try:
            current_modified = os.path.getmtime(self.file_path)
            if current_modified > self.last_modified:
                self.last_modified = current_modified
                stop_dots = True
                self.process_and_update()
                print("\rDATA LOADED")
                threading.Thread(target=print_waiting_dots, daemon=True).start()
        except:
            pass

    def process_and_update(self):
        global compiledFrame, dataList
        try:
            dataList = []
            try:
                if self.last_row_count == 0:
                    df = pd.read_csv(self.file_path, encoding='latin1', dtype={'S/N': str, 'MODEL CODE': str})
                    self.df_columns = df.columns.tolist()
                else:
                    df = pd.read_csv(self.file_path, encoding='latin1', skiprows=self.last_row_count, header=None, names=self.df_columns, dtype={'S/N': str, 'MODEL CODE': str})
            except UnicodeDecodeError:
                if self.last_row_count == 0:
                    df = pd.read_csv(self.file_path, encoding='ISO-8859-1', errors='replace', dtype={'S/N': str, 'MODEL CODE': str})
                    self.df_columns = df.columns.tolist()
                else:
                    df = pd.read_csv(self.file_path, encoding='ISO-8859-1', errors='replace', skiprows=self.last_row_count, header=None, names=self.df_columns, dtype={'S/N': str, 'MODEL CODE': str})
            if self.last_row_count == 0:
                self.last_row_count = len(df)
            else:
                self.last_row_count += len(df)
            df = df[~df["MODEL CODE"].isin(['60CAT0203M', '60CAT0202P', '60CAT0203P', '60FC00000P', 
                                            '60FC00902P', '60FC00903P', '60FC00905P', '60FCXP001P', 
                                            '30FCXP001P'])]
            df = df[df["PASS/NG"].isin([0, 1])]
            df['S/N'] = df['S/N'].astype(str)
            df['MODEL CODE'] = df['MODEL CODE'].astype(str)
            # df = df[df['MODEL CODE'] == self.current_model]
            df['DATE'] = pd.to_datetime(df['DATE'], format='%Y/%m/%d', errors='coerce')
            df = df.dropna(subset=['DATE'])
            df = df[df['S/N'].str.len() >= 8]
            df = df[~df['MODEL CODE'].str.contains('M')]
            df = df[~df['MODEL CODE'].str.contains('60CAT0213T')]
            def fix_time(time_str):
                try:
                    hours, minutes, seconds = map(int, time_str.split(':'))
                    if hours >= 24:
                        hours = hours % 24
                        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    return time_str
                except:
                    return time_str
            df['TIME'] = df['TIME'].apply(fix_time)
            df['DATETIME'] = pd.to_datetime(
                df['DATE'].astype(str) + ' ' + df['TIME'],
                format='%Y-%m-%d %H:%M:%S',
                errors='coerce'
            )
            df = df.dropna(subset=['DATETIME'])
            df = df.sort_values('DATETIME')
            emptyColumn = [
                "DATE", "TIME", "MODEL CODE", "S/N", "PASS/NG",
                "VOLTAGE MAX (V)", "V_MAX PASS", "AVE V_MAX PASS", "DEV V_MAX PASS", "DEV V_MAX PASS FLUCTUATED",
                "WATTAGE MAX (W)", "WATTAGE MAX PASS", "AVE WATTAGE MAX (W)", "DEV WATTAGE MAX (W)", "DEV WATTAGE MAX (W) FLUCTUATED",
                "CLOSED PRESSURE_MAX (kPa)", "CLOSED PRESSURE_MAX PASS", "AVE CLOSED PRESSURE_MAX (kPa)", "DEV CLOSED PRESSURE_MAX (kPa)", "DEV CLOSED PRESSURE_MAX (kPa) FLUCTUATED",
                "VOLTAGE Middle (V)", "VOLTAGE Middle PASS", "AVE VOLTAGE Middle (V)", "DEV VOLTAGE Middle (V)", "DEV VOLTAGE Middle (V) FLUCTUATED",
                "WATTAGE Middle (W)", "WATTAGE Middle (W) PASS", "AVE WATTAGE Middle (W)", "DEV WATTAGE Middle (W)", "DEV WATTAGE Middle (W) FLUCTUATED",
                "AMPERAGE Middle (A)", "AMPERAGE Middle (A) PASS", "AVE AMPERAGE Middle (A)", "DEV AMPERAGE Middle (A)", "DEV AMPERAGE Middle (A) FLUCTUATED",
                "CLOSED PRESSURE Middle (kPa)", "CLOSED PRESSURE Middle (kPa) PASS", "AVE CLOSED PRESSURE Middle (kPa)", "DEV CLOSED PRESSURE Middle (kPa)", "DEV CLOSED PRESSURE Middle (kPa) FLUCTUATED",
                "VOLTAGE MIN (V)", "VOLTAGE MIN (V) PASS", "AVE VOLTAGE MIN (V)", "DEV VOLTAGE MIN (V)", "DEV VOLTAGE MIN (V) FLUCTUATED",
                "WATTAGE MIN (W)", "WATTAGE MIN (W) PASS", "AVE WATTAGE MIN (W)", "DEV WATTAGE MIN (W)", "DEV WATTAGE MIN (W) FLUCTUATED",
                "CLOSED PRESSURE MIN (kPa)", "CLOSED PRESSURE MIN (kPa) PASS", "AVE CLOSED PRESSURE MIN (kPa)", "DEV CLOSED PRESSURE MIN (kPa)", "DEV CLOSED PRESSURE MIN (kPa) FLUCTUATED",
                "REFERENCE SERIAL", "DATETIME"
            ]
            if not hasattr(self, 'compiledFrame') or self.compiledFrame.empty:
                self.compiledFrame = pd.DataFrame(columns=emptyColumn)
            today = pd.to_datetime(datetime.now().date())
            results = []
            past_compiled = self.compiledFrame[self.compiledFrame['PASS/NG'] == 1].copy() if not self.compiledFrame.empty else pd.DataFrame()
            if not past_compiled.empty:
                past_compiled['DATE'] = pd.to_datetime(past_compiled['DATE'], format='%Y/%m/%d', errors='coerce')
                past_compiled = past_compiled.dropna(subset=['DATE'])
                past_data = past_compiled[past_compiled['DATE'].dt.date < today.date()]
            else:
                past_data = pd.DataFrame()
            use_df_past = past_data.empty
            if use_df_past:
                df_past = df[df["PASS/NG"] == 1].copy()
                if not df_past.empty:
                    df_past['DATE'] = pd.to_datetime(df_past['DATE'], format='%Y/%m/%d', errors='coerce')
                    df_past = df_past.dropna(subset=['DATE'])
                    past_data = df_past[df_past['DATE'].dt.date < today.date()]
            if not past_data.empty:
                group = past_data[past_data['MODEL CODE'] == self.current_model]
                if not group.empty:
                    past_data_sorted = group.sort_values('DATE', ascending=False)
                    dates = sorted(past_data_sorted['DATE'].dt.date.unique(), reverse=True)
                    if dates:
                        accumulated_rows = pd.DataFrame(columns=group.columns)
                        count = 0
                        for date in dates:
                            daily_rows = past_data_sorted[past_data_sorted['DATE'].dt.date == date]
                            accumulated_rows = pd.concat([accumulated_rows, daily_rows], ignore_index=True)
                            count += len(daily_rows)
                            if count >= 200:
                                break
                        accumulated_rows = accumulated_rows.sort_values('DATETIME', ascending=False).head(200)
                        count = len(accumulated_rows)
                        if count > 0:
                            newest_date = accumulated_rows.iloc[0]['DATE'].date()
                            oldest_date = accumulated_rows.iloc[-1]['DATE'].date()
                            end_serial = accumulated_rows.iloc[-1]['S/N']
                        results.append({
                                'MODEL CODE': self.current_model,
                                'START_DATE': newest_date,
                                'END_DATE': oldest_date,
                                'NUM_ROWS': count,
                                'V-MAX PASS AVG': accumulated_rows["VOLTAGE MAX (V)"].mean(),
                                'WATTAGE MAX AVG': accumulated_rows["WATTAGE MAX (W)"].mean(),
                                'CLOSED PRESSURE_MAX AVG': accumulated_rows["CLOSED PRESSURE_MAX (kPa)"].mean(),
                                'VOLTAGE Middle AVG': accumulated_rows["VOLTAGE Middle (V)"].mean(),
                                'WATTAGE Middle AVG': accumulated_rows["WATTAGE Middle (W)"].mean(),
                                'AMPERAGE Middle AVG': accumulated_rows["AMPERAGE Middle (A)"].mean(),
                                'CLOSED PRESSURE Middle AVG': accumulated_rows["CLOSED PRESSURE Middle (kPa)"].mean(),
                                'VOLTAGE MIN (V) AVG': accumulated_rows["VOLTAGE MIN (V)"].mean(),
                                'WATTAGE MIN AVG': accumulated_rows["WATTAGE MIN (W)"].mean(),
                                'CLOSED PRESSURE MIN AVG': accumulated_rows["CLOSED PRESSURE MIN (kPa)"].mean(),
                                'END_SERIAL': end_serial
                            })
            model_summary = pd.DataFrame(results)
            self.model_summary = model_summary
            pass_avg_map = model_summary.set_index("MODEL CODE")["V-MAX PASS AVG"].to_dict()
            wattage_avg_map = model_summary.set_index("MODEL CODE")["WATTAGE MAX AVG"].to_dict()
            closedPressure_avg_map = model_summary.set_index("MODEL CODE")['CLOSED PRESSURE_MAX AVG'].to_dict()
            voltageMiddle_avg_map = model_summary.set_index("MODEL CODE")["VOLTAGE Middle AVG"].to_dict()
            wattageMiddle_avg = model_summary.set_index("MODEL CODE")["WATTAGE Middle AVG"].to_dict()
            amperageMiddle_avg = model_summary.set_index("MODEL CODE")["AMPERAGE Middle AVG"].to_dict()
            closePressureMiddle_avg = model_summary.set_index("MODEL CODE")["CLOSED PRESSURE Middle AVG"].to_dict()
            voltageMin_avg = model_summary.set_index("MODEL CODE")["VOLTAGE MIN (V) AVG"].to_dict()
            wattageMin_avg = model_summary.set_index("MODEL CODE")["WATTAGE MIN AVG"].to_dict()
            closePressureMin_avg = model_summary.set_index("MODEL CODE")["CLOSED PRESSURE MIN AVG"].to_dict()
            previous_values = {
                'V_MAX': None,
                'W_MAX': None,
                'CP_MAX': None,
                'V_MID': None,
                'W_MID': None,
                'A_MID': None,
                'CP_MID': None,
                'V_MIN': None,
                'W_MIN': None,
                'CP_MIN': None
            }
            previous_date = None
            previous_model = None
            existing_sn = set(self.compiledFrame['S/N'].astype(str)) if not self.compiledFrame.empty else set()
            new_df = df.copy()
            if new_df.empty:
                logging.info("No new data to process")
                self.update_combo()
                #self.update_line_graph()
                self.update_display()
                return
            new_df = new_df[new_df["PASS/NG"] == 1]
            new_df = new_df.sort_values('DATETIME')
            dataList = []
            for _, row in new_df.iterrows():
                model_code = row["MODEL CODE"]
                serial_no = row["S/N"]
                current_date = row["DATE"]
                self.serial_display.config(text=f"SERIAL No.: {serial_no}")
                self.model_display.config(text=f"MODEL CODE: {model_code}")
                self.serial_log.config(state='normal')
                self.serial_log.insert(tk.END, f"{serial_no}\n")
                self.serial_log.see(tk.END)
                self.serial_log.config(state='disabled')
                self.current_model = model_code
                self.last_date = current_date
                dataFrame = {
                    "DATE": row["DATE"].strftime('%Y/%m/%d'),
                    "TIME": row["TIME"],
                    "MODEL CODE": model_code,
                    "S/N": serial_no,
                    "PASS/NG": row["PASS/NG"],
                    "VOLTAGE MAX (V)": row["VOLTAGE MAX (V)"],
                    "V_MAX PASS": row["VOLTAGE MAX (V)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE V_MAX PASS": pass_avg_map.get(model_code),
                    "DEV V_MAX PASS": 0,
                    "DEV V_MAX PASS FLUCTUATED": 0,
                    "WATTAGE MAX (W)": row["WATTAGE MAX (W)"],
                    "WATTAGE MAX PASS": row["WATTAGE MAX (W)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE WATTAGE MAX (W)": wattage_avg_map.get(model_code),
                    "DEV WATTAGE MAX (W)": 0,
                    "DEV WATTAGE MAX (W) FLUCTUATED": 0,
                    "CLOSED PRESSURE_MAX (kPa)": row["CLOSED PRESSURE_MAX (kPa)"],
                    "CLOSED PRESSURE_MAX PASS": row["CLOSED PRESSURE_MAX (kPa)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE CLOSED PRESSURE_MAX (kPa)": closedPressure_avg_map.get(model_code),
                    "DEV CLOSED PRESSURE_MAX (kPa)": 0,
                    "DEV CLOSED PRESSURE_MAX (kPa) FLUCTUATED": 0,
                    "VOLTAGE Middle (V)": row["VOLTAGE Middle (V)"],
                    "VOLTAGE Middle PASS": row["VOLTAGE Middle (V)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE VOLTAGE Middle (V)": voltageMiddle_avg_map.get(model_code),
                    "DEV VOLTAGE Middle (V)": 0,
                    "DEV VOLTAGE Middle (V) FLUCTUATED": 0,
                    "WATTAGE Middle (W)": row["WATTAGE Middle (W)"],
                    "WATTAGE Middle (W) PASS": row["WATTAGE Middle (W)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE WATTAGE Middle (W)": wattageMiddle_avg.get(model_code),
                    "DEV WATTAGE Middle (W)": 0,
                    "DEV WATTAGE Middle (W) FLUCTUATED": 0,
                    "AMPERAGE Middle (A)": row["AMPERAGE Middle (A)"],
                    "AMPERAGE Middle (A) PASS": row["AMPERAGE Middle (A)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE AMPERAGE Middle (A)": amperageMiddle_avg.get(model_code),
                    "DEV AMPERAGE Middle (A)": 0,
                    "DEV AMPERAGE Middle (A) FLUCTUATED": 0,
                    "CLOSED PRESSURE Middle (kPa)": row["CLOSED PRESSURE Middle (kPa)"],
                    "CLOSED PRESSURE Middle (kPa) PASS": row["CLOSED PRESSURE Middle (kPa)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE CLOSED PRESSURE Middle (kPa)": closePressureMiddle_avg.get(model_code),
                    "DEV CLOSED PRESSURE Middle (kPa)": 0,
                    "DEV CLOSED PRESSURE Middle (kPa) FLUCTUATED": 0,
                    "VOLTAGE MIN (V)": row["VOLTAGE MIN (V)"],
                    "VOLTAGE MIN (V) PASS": row["VOLTAGE MIN (V)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE VOLTAGE MIN (V)": voltageMin_avg.get(model_code),
                    "DEV VOLTAGE MIN (V)": 0,
                    "DEV VOLTAGE MIN (V) FLUCTUATED": 0,
                    "WATTAGE MIN (W)": row["WATTAGE MIN (W)"],
                    "WATTAGE MIN (W) PASS": row["WATTAGE MIN (W)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE WATTAGE MIN (W)": wattageMin_avg.get(model_code),
                    "DEV WATTAGE MIN (W)": 0,
                    "DEV WATTAGE MIN (W) FLUCTUATED": 0,
                    "CLOSED PRESSURE MIN (kPa)": row["CLOSED PRESSURE MIN (kPa)"],
                    "CLOSED PRESSURE MIN (kPa) PASS": row["CLOSED PRESSURE MIN (kPa)"] if row["PASS/NG"] == 1 else np.nan,
                    "AVE CLOSED PRESSURE MIN (kPa)": closePressureMin_avg.get(model_code),
                    "DEV CLOSED PRESSURE MIN (kPa)": 0,
                    "DEV CLOSED PRESSURE MIN (kPa) FLUCTUATED": 0,
                    "REFERENCE SERIAL": "N/A",
                    "DATETIME": row["DATETIME"]
                }
                if row["PASS/NG"] == 1:
                    dataFrame["V_MAX PASS"] = row["VOLTAGE MAX (V)"]
                    dataFrame["WATTAGE MAX PASS"] = row["WATTAGE MAX (W)"]
                    dataFrame["CLOSED PRESSURE_MAX PASS"] = row["CLOSED PRESSURE_MAX (kPa)"]
                    dataFrame["VOLTAGE Middle PASS"] = row["VOLTAGE Middle (V)"]
                    dataFrame["WATTAGE Middle (W) PASS"] = row["WATTAGE Middle (W)"]
                    dataFrame["AMPERAGE Middle (A) PASS"] = row["AMPERAGE Middle (A)"]
                    dataFrame["CLOSED PRESSURE Middle (kPa) PASS"] = row["CLOSED PRESSURE Middle (kPa)"]
                    dataFrame["VOLTAGE MIN (V) PASS"] = row["VOLTAGE MIN (V)"]
                    dataFrame["WATTAGE MIN (W) PASS"] = row["WATTAGE MIN (W)"]
                    dataFrame["CLOSED PRESSURE MIN (kPa) PASS"] = row["CLOSED PRESSURE MIN (kPa)"]
                previous_pass = self.compiledFrame[(self.compiledFrame['MODEL CODE'] == model_code) & (self.compiledFrame['PASS/NG'] == 1)].copy()
                previous_pass['DATE'] = pd.to_datetime(previous_pass['DATE'], format='%Y/%m/%d', errors='coerce')
                previous_pass = previous_pass[previous_pass['DATE'].dt.date < current_date.date()]
                previous_pass = previous_pass.tail(200)
                ave_v_max = previous_pass["VOLTAGE MAX (V)"].mean() if not previous_pass.empty else 0
                ave_w_max = previous_pass["WATTAGE MAX (W)"].mean() if not previous_pass.empty else 0
                ave_cp_max = previous_pass["CLOSED PRESSURE_MAX (kPa)"].mean() if not previous_pass.empty else 0
                ave_v_mid = previous_pass["VOLTAGE Middle (V)"].mean() if not previous_pass.empty else 0
                ave_w_mid = previous_pass["WATTAGE Middle (W)"].mean() if not previous_pass.empty else 0
                ave_a_mid = previous_pass["AMPERAGE Middle (A)"].mean() if not previous_pass.empty else 0
                ave_cp_mid = previous_pass["CLOSED PRESSURE Middle (kPa)"].mean() if not previous_pass.empty else 0
                ave_v_min = previous_pass["VOLTAGE MIN (V)"].mean() if not previous_pass.empty else 0
                ave_w_min = previous_pass["WATTAGE MIN (W)"].mean() if not previous_pass.empty else 0
                ave_cp_min = previous_pass["CLOSED PRESSURE MIN (kPa)"].mean() if not previous_pass.empty else 0
                dataFrame["AVE V_MAX PASS"] = ave_v_max
                dataFrame["AVE WATTAGE MAX (W)"] = ave_w_max
                dataFrame["AVE CLOSED PRESSURE_MAX (kPa)"] = ave_cp_max
                dataFrame["AVE VOLTAGE Middle (V)"] = ave_v_mid
                dataFrame["AVE WATTAGE Middle (W)"] = ave_w_mid
                dataFrame["AVE AMPERAGE Middle (A)"] = ave_a_mid
                dataFrame["AVE CLOSED PRESSURE Middle (kPa)"] = ave_cp_mid
                dataFrame["AVE VOLTAGE MIN (V)"] = ave_v_min
                dataFrame["AVE WATTAGE MIN (W)"] = ave_w_min
                dataFrame["AVE CLOSED PRESSURE MIN (kPa)"] = ave_cp_min
                current_values = {
                    'V_MAX': row["VOLTAGE MAX (V)"],
                    'W_MAX': row["WATTAGE MAX (W)"],
                    'CP_MAX': row["CLOSED PRESSURE_MAX (kPa)"],
                    'V_MID': row["VOLTAGE Middle (V)"],
                    'W_MID': row["WATTAGE Middle (W)"],
                    'A_MID': row["AMPERAGE Middle (A)"],
                    'CP_MID': row["CLOSED PRESSURE Middle (kPa)"],
                    'V_MIN': row["VOLTAGE MIN (V)"],
                    'W_MIN': row["WATTAGE MIN (W)"],
                    'CP_MIN': row["CLOSED PRESSURE MIN (kPa)"]
                }
                fluctuations = {}
                ref_serial = self.last_good_serial if self.last_good_serial else "N/A"
                for key in current_values:
                    if current_values[key] == 0:
                        fluctuations[key] = 0
                    else:
                        avg_key = {
                            'V_MAX': "AVE V_MAX PASS",
                            'W_MAX': "AVE WATTAGE MAX (W)",
                            'CP_MAX': "AVE CLOSED PRESSURE_MAX (kPa)",
                            'V_MID': "AVE VOLTAGE Middle (V)",
                            'W_MID': "AVE WATTAGE Middle (W)",
                            'A_MID': "AVE AMPERAGE Middle (A)",
                            'CP_MID': "AVE CLOSED PRESSURE Middle (kPa)",
                            'V_MIN': "AVE VOLTAGE MIN (V)",
                            'W_MIN': "AVE WATTAGE MIN (W)",
                            'CP_MIN': "AVE CLOSED PRESSURE MIN (kPa)"
                        }[key]
                        avg_value = dataFrame[avg_key]
                        if pd.notnull(avg_value) and avg_value != 0:
                            fluctuations[key] = (avg_value - current_values[key]) / avg_value
                        else:
                            fluctuations[key] = 0
                has_fluct = any(abs(fluctuations[self.short_map[m]]) > self.threshold for m in self.selected_measurements)
                ref_serial = self.last_good_serial if self.last_good_serial else "N/A"
                dataFrame["REFERENCE SERIAL"] = ref_serial
                if not has_fluct:
                    self.last_good_values = current_values.copy()
                    self.last_good_serial = serial_no
                    ref_serial = serial_no
                dataFrame["DEV V_MAX PASS"] = fluctuations['V_MAX']
                dataFrame["DEV WATTAGE MAX (W)"] = fluctuations['W_MAX']
                dataFrame["DEV CLOSED PRESSURE_MAX (kPa)"] = fluctuations['CP_MAX']
                dataFrame["DEV VOLTAGE Middle (V)"] = fluctuations['V_MID']
                dataFrame["DEV WATTAGE Middle (W)"] = fluctuations['W_MID']
                dataFrame["DEV AMPERAGE Middle (A)"] = fluctuations['A_MID']
                dataFrame["DEV CLOSED PRESSURE Middle (kPa)"] = fluctuations['CP_MID']
                dataFrame["DEV VOLTAGE MIN (V)"] = fluctuations['V_MIN']
                dataFrame["DEV WATTAGE MIN (W)"] = fluctuations['W_MIN']
                dataFrame["DEV CLOSED PRESSURE MIN (kPa)"] = fluctuations['CP_MIN']
                dataFrame["DEV V_MAX PASS FLUCTUATED"] = abs(fluctuations['V_MAX']) if fluctuations['V_MAX'] != 0 else 0
                dataFrame["DEV WATTAGE MAX (W) FLUCTUATED"] = abs(fluctuations['W_MAX']) if fluctuations['W_MAX'] != 0 else 0
                dataFrame["DEV CLOSED PRESSURE_MAX (kPa) FLUCTUATED"] = abs(fluctuations['CP_MAX']) if fluctuations['CP_MAX'] != 0 else 0
                dataFrame["DEV VOLTAGE Middle (V) FLUCTUATED"] = abs(fluctuations['V_MID']) if fluctuations['V_MID'] != 0 else 0
                dataFrame["DEV WATTAGE Middle (W) FLUCTUATED"] = abs(fluctuations['W_MID']) if fluctuations['W_MID'] != 0 else 0
                dataFrame["DEV AMPERAGE Middle (A) FLUCTUATED"] = abs(fluctuations['A_MID']) if fluctuations['A_MID'] != 0 else 0
                dataFrame["DEV CLOSED PRESSURE Middle (kPa) FLUCTUATED"] = abs(fluctuations['CP_MID']) if fluctuations['CP_MID'] != 0 else 0
                dataFrame["DEV VOLTAGE MIN (V) FLUCTUATED"] = abs(fluctuations['V_MIN']) if fluctuations['V_MIN'] != 0 else 0
                dataFrame["DEV WATTAGE MIN (W) FLUCTUATED"] = abs(fluctuations['W_MIN']) if fluctuations['W_MIN'] != 0 else 0
                dataFrame["DEV CLOSED PRESSURE MIN (kPa) FLUCTUATED"] = abs(fluctuations['CP_MIN']) if fluctuations['CP_MIN'] != 0 else 0
                dataFrame["REFERENCE SERIAL"] = ref_serial
                previous_values = current_values.copy()
                previous_date = current_date
                previous_model = model_code
                dataList.append(dataFrame)
            if dataList:
                new_data = pd.DataFrame(dataList)
                for col in emptyColumn:
                    if col not in new_data.columns:
                        new_data[col] = pd.Series(dtype='object')
                self.compiledFrame = pd.concat([self.compiledFrame, new_data]).drop_duplicates(subset=['S/N', 'DATETIME'], keep='last').sort_values('DATETIME').reset_index(drop=True)
                self.compiledFrame = self.compiledFrame[self.compiledFrame['PASS/NG'] == 1].reset_index(drop=True)
                if self.generate_csv == "YES":
                    try:
                        self.compiledFrame.to_csv(self.output_path, index=False, encoding='utf-8-sig')
                    except Exception as e:
                        logging.error(f"Error saving CSV to {self.output_path}: {e}")
            compiledFrame = self.compiledFrame
            results = []
            unique_models = compiledFrame['MODEL CODE'].unique()
            for model in unique_models:
                group = compiledFrame[compiledFrame['MODEL CODE'] == model].copy()
                group['DATE'] = pd.to_datetime(group['DATE'], format='%Y/%m/%d', errors='coerce')
                group = group[group['DATE'].dt.date < datetime.now().date()]
                accumulated_rows = group.sort_values('DATETIME', ascending=False).head(200)
                count = len(accumulated_rows)
                if count > 0:
                    oldest_date = accumulated_rows.iloc[0]['DATE'].date()
                    newest_date = accumulated_rows.iloc[-1]['DATE'].date()
                    end_serial = accumulated_rows.iloc[0]['S/N']
                    results.append({
                            'MODEL CODE': model,
                            'START_DATE': newest_date,
                            'END_DATE': oldest_date,
                            'NUM_ROWS': count,
                            'V-MAX PASS AVG': accumulated_rows["VOLTAGE MAX (V)"].mean(),
                            'WATTAGE MAX AVG': accumulated_rows["WATTAGE MAX (W)"].mean(),
                            'CLOSED PRESSURE_MAX AVG': accumulated_rows["CLOSED PRESSURE_MAX (kPa)"].mean(),
                            'VOLTAGE Middle AVG': accumulated_rows["VOLTAGE Middle (V)"].mean(),
                            'WATTAGE Middle AVG': accumulated_rows["WATTAGE Middle (W)"].mean(),
                            'AMPERAGE Middle AVG': accumulated_rows["AMPERAGE Middle (A)"].mean(),
                            'CLOSED PRESSURE Middle AVG': accumulated_rows["CLOSED PRESSURE Middle (kPa)"].mean(),
                            'VOLTAGE MIN (V) AVG': accumulated_rows["VOLTAGE MIN (V)"].mean(),
                            'WATTAGE MIN AVG': accumulated_rows["WATTAGE MIN (W)"].mean(),
                            'CLOSED PRESSURE MIN AVG': accumulated_rows["CLOSED PRESSURE MIN (kPa)"].mean(),
                            'END_SERIAL': end_serial
                        })
            self.model_summary = pd.DataFrame(results)
            self.update_combo()
            #self.update_line_graph()
            self.update_display()
            logging.info("Data loaded and processed successfully")
        except Exception as e:
            logging.error(f"Error processing file: {e}")

    def update_display(self):
        try:
            last_row = self.compiledFrame.iloc[-1]
            has_fluctuation = False
            fluctuation_count = 0
            for column in self.status_vars:
                if column in last_row:
                    value = last_row[column]
                    status = 1 if value > self.threshold else 0
                    self.status_vars[column].set(f"{status}")
                    color = "red" if status == 1 else "green"
                    self.status_labels[column].configure(foreground=color)
                    if status == 1:
                        has_fluctuation = True
                        fluctuation_count += 1
            self.update_status_box(has_fluctuation, fluctuation_count)
            #self.update_bar_graph(last_row)
            if has_fluctuation:
                self.add_to_log(last_row, fluctuation_count)
            if "MODEL CODE" in last_row:
                model_code = last_row["MODEL CODE"]
                self.model_display.config(text=f"MODEL CODE: {model_code}")
            self.update_spc_graphs()
        except Exception as e:
            logging.error(f"Error updating display: {e}")

    def create_section(self, title, columns):
        frame = ttk.Frame(self.left_frame)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=title, font=('Arial', 12, 'bold')).pack(anchor='w')
        for col in columns:
            self.create_status_row(frame, col)

    def create_status_row(self, parent, column_name):
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, pady=1)
        ttk.Label(row_frame, text=f"{column_name.replace(' FLUCTUATED', '')}:", 
                 width=20, anchor='w').pack(side=tk.LEFT)
        status_frame = ttk.Frame(row_frame)
        status_frame.pack(side=tk.LEFT, padx=2)
        self.status_vars[column_name] = tk.StringVar()
        self.status_labels[column_name] = ttk.Label(
            status_frame, 
            textvariable=self.status_vars[column_name],
            font=('Arial', 10, 'bold'),
            width=8
        )
        self.status_labels[column_name].pack(side=tk.LEFT)
        ttk.Button(
            status_frame, 
            text="Reset", 
            width=5,
            command=lambda: self.reset_fluctuation(column_name)
        ).pack(side=tk.LEFT, padx=2)

    def reset_fluctuation(self, column_name):
        try:
            self.compiledFrame.at[self.compiledFrame.index[-1], column_name] = 0
            self.status_vars[column_name].set("0")
            self.status_labels[column_name].configure(foreground="green")
            last_row = self.compiledFrame.iloc[-1]
            if all(last_row[col] <= self.threshold for col in self.status_vars if col in last_row):
                self.last_good_values = {
                    'V_MAX': last_row["VOLTAGE MAX (V)"],
                    'W_MAX': last_row["WATTAGE MAX (W)"],
                    'CP_MAX': last_row["CLOSED PRESSURE_MAX (kPa)"],
                    'V_MID': last_row["VOLTAGE Middle (V)"],
                    'W_MID': last_row["WATTAGE Middle (W)"],
                    'A_MID': last_row["AMPERAGE Middle (A)"],
                    'CP_MID': last_row["CLOSED PRESSURE Middle (kPa)"],
                    'V_MIN': last_row["VOLTAGE MIN (V)"],
                    'W_MIN': last_row["WATTAGE MIN (W)"],
                    'CP_MIN': last_row["CLOSED PRESSURE MIN (kPa)"]
                }
                self.last_good_serial = last_row["S/N"]
                self.compiledFrame.at[self.compiledFrame.index[-1], "REFERENCE SERIAL"] = self.last_good_serial
            if self.generate_csv == "YES":
                try:
                    self.compiledFrame.to_csv(self.output_path, index=False, encoding='utf-8-sig')
                except Exception as e:
                    logging.error(f"Error saving CSV in reset_fluctuation: {e}")
            self.update_display()
        except Exception as e:
            logging.error(f"Error resetting fluctuation: {e}")

    def reset_all_fluctuations(self):
        try:
            for column in self.status_vars:
                self.compiledFrame.at[self.compiledFrame.index[-1], column] = 0
                self.status_vars[column].set("0")
                self.status_labels[column].configure(foreground="green")
            last_row = self.compiledFrame.iloc[-1]
            self.last_good_values = {
                'V_MAX': last_row["VOLTAGE MAX (V)"],
                'W_MAX': last_row["WATTAGE MAX (W)"],
                'CP_MAX': last_row["CLOSED PRESSURE_MAX (kPa)"],
                'V_MID': last_row["VOLTAGE Middle (V)"],
                'W_MID': last_row["WATTAGE Middle (W)"],
                'A_MID': last_row["AMPERAGE Middle (A)"],
                'CP_MID': last_row["CLOSED PRESSURE Middle (kPa)"],
                'V_MIN': last_row["VOLTAGE MIN (V)"],
                'W_MIN': last_row["WATTAGE MIN (W)"],
                'CP_MIN': last_row["CLOSED PRESSURE MIN (kPa)"]
            }
            self.last_good_serial = last_row["S/N"]
            self.compiledFrame.at[self.compiledFrame.index[-1], "REFERENCE SERIAL"] = self.last_good_serial
            if self.generate_csv == "YES":
                try:
                    self.compiledFrame.to_csv(self.output_path, index=False, encoding='utf-8-sig')
                except Exception as e:
                    logging.error(f"Error saving CSV in reset_all_fluctuations: {e}")
            self.update_status_box(False)
            self.update_display()
        except Exception as e:
            logging.error(f"Error resetting all fluctuations: {e}")

    def on_closing(self):
        if self.after_id is not None:
            self.root.after_cancel(self.after_id)
        self.observer.stop()
        self.observer.join()
        self.root.destroy()

    def update_combo(self):
        global compiledFrame
        if compiledFrame is not None:
            model_groups = compiledFrame.groupby('MODEL CODE').size()
            unique_models = sorted([model for model, size in model_groups.items() if size >= 10])
            self.combo['values'] = unique_models
            if self.current_model not in unique_models:
                self.current_model = None
                self.combo.set('')

    def on_select(self, event):
        self.current_model = self.combo.get()
        self.update_spc_graphs()

    def on_resize(self, event):
        if event.widget == self.root:
            if self.old_width is not None and event.width == self.old_width and event.height == self.old_height:
                return
            self.old_width = event.width
            self.old_height = event.height
            try:
                if hasattr(self, 'spc_graph_frame') and self.spc_graph_frame.winfo_height() > 0 and self.spc_graph_frame.winfo_width() > 0:
                    dpi = 100
                    if self.layout_mode.get() == "HORIZONTAL":
                        num_cols = 2
                        num_rows = 2
                        new_width = self.spc_graph_frame.winfo_width() / num_cols / dpi * 0.95
                        new_height = self.spc_graph_frame.winfo_height() / num_rows / dpi * 0.95
                    else:  # VERTICAL
                        new_width = 10  # fixed
                        new_height = 3  # fixed
                    for fig in self.figs.values():
                        fig.set_size_inches(new_width, new_height)
                    for canvas in self.canvases.values():
                        canvas.draw()
            except Exception as e:
                logging.error(f"Error in on_resize: {e}")
# 3RD TKINTER FOCUS
    def open_focus_window(self):
        self.focus_window = tk.Toplevel(self.root)
        self.focus_window.title("Focus on Historical Historical Trends")
        self.focus_window.geometry(f"{int(self.root.winfo_screenwidth() * 0.9)}x{int(self.root.winfo_screenheight() * 0.9)}")
        button_frame = ttk.Frame(self.focus_window)
        button_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        ttk.Label(button_frame, text="Select Measurement:", font=('Arial', 14, 'bold')).pack(pady=10)
        self.measurement_buttons = {}
        for measurement in self.selected_measurements:
            btn = ttk.Button(button_frame, text=measurement, command=lambda m=measurement: self.show_focus_graph(m))
            btn.pack(fill=tk.X, pady=5)
            self.measurement_buttons[measurement] = btn
        self.focus_graph_frame = ttk.Frame(self.focus_window)
        self.focus_graph_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.focus_message_label = ttk.Label(self.focus_graph_frame, text="", font=('Arial', 10, 'bold'), foreground="red")
        self.focus_message_label.pack(anchor='s')
        back_btn = ttk.Button(self.focus_window, text="Back", command=self.focus_window.destroy)
        back_btn.pack(side=tk.BOTTOM, pady=10)

    def show_focus_graph(self, measurement):
        for widget in self.focus_graph_frame.winfo_children():
            if widget != self.focus_message_label:
                widget.destroy()
        fig = Figure(figsize=(12, 8), dpi=100)
        ax = fig.add_subplot(111)
        # ax.set_title(f'Historical Trend for {measurement}')
        ax.set_ylabel('Value')
        ax.grid(True)
        df = self.compiledFrame[self.compiledFrame["PASS/NG"] == 1].copy()
        if self.current_model:
            df = df[df['MODEL CODE'] == self.current_model].copy()
        df['DATE'] = pd.to_datetime(df['DATE'], format='%Y/%m/%d', errors='coerce')
        today = pd.to_datetime(datetime.now().date())
        df = df.tail(200)
        df['DATETIME'] = pd.to_datetime(df['DATETIME'], errors='coerce')
        df = df.sort_values('DATETIME')
        if not df.empty:
            x = np.arange(len(df))
            y = df[measurement].values
            last_row = df.iloc[-1]
            fluct_col = self.measurements_map[measurement][3]
            is_fluctuated = last_row[fluct_col] > self.threshold if fluct_col in last_row else False
            last_color = 'lightcoral' if is_fluctuated else 'green'
            n = len(df)
            if n > 1:
                ax.plot(x[:-1], y[:-1], color='blue', linewidth=2)
                ax.plot(x[-2:], y[-2:], color=last_color, linewidth=2, linestyle='--')
            else:
                ax.plot(x, y, color='blue', linewidth=2)
            last_x = x[-1]
            last_y = y[-1]
            circle_color = 'red' if is_fluctuated else 'none'
            ax.scatter(last_x, last_y, color=circle_color, s=100, marker='o', zorder=5, linewidth=2, edgecolors='black' if is_fluctuated else 'none')
            num_ticks = min(10, len(df))
            if num_ticks > 0:
                tick_positions = np.linspace(0, len(df)-1, num_ticks, dtype=int)
                if self.x_axis_mode.get() == "Numerical":
                    tick_labels = [str(pos + 1) for pos in tick_positions]
                    ax.set_xlabel('Sample Number')
                elif self.x_axis_mode.get() == "DateTime":
                    tick_labels = df['DATETIME'].iloc[tick_positions].dt.strftime('%Y-%m-%d %H:%M:%S')
                    ax.set_xlabel('DateTime')
                else:
                    tick_labels = ['' for _ in tick_positions]
                    ax.set_xlabel('')
                    ax.set_xticks([])
                    ax.set_xticklabels([])
                ax.set_xticks(tick_positions)
                ax.set_xticklabels(tick_labels, rotation=45, ha='right')
        fig.subplots_adjust(left=0.15, right=0.95, bottom=0.1, top=0.95)
        canvas = FigureCanvasTkAgg(fig, master=self.focus_graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.focus_message_label.config(text="FLUCTUATED" if is_fluctuated else "")

    def open_display_window(self):
        self.display_window = tk.Toplevel(self.root)
        self.display_window.title("Display Settings")
        self.display_window.geometry(f"{int(self.root.winfo_screenwidth() * 0.3)}x{int(self.root.winfo_screenheight() * 0.3)}")
        main_container = ttk.Frame(self.display_window)
        main_container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(main_container)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        ttk.Label(scrollable_frame, text="X-AXIS:", font=('Arial', 12)).pack(pady=10)
        ttk.Radiobutton(scrollable_frame, text="Numerical Numbering", variable=self.x_axis_mode, value="Numerical").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="DateTime", variable=self.x_axis_mode, value="DateTime").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="None", variable=self.x_axis_mode, value="NONE").pack(anchor='w', padx=20)
        ttk.Checkbutton(scrollable_frame, text="Show Title", variable=self.show_title_var).pack(anchor='w', padx=20)
        ttk.Label(scrollable_frame, text="LAYOUT:", font=('Arial', 12)).pack(pady=10)
        ttk.Radiobutton(scrollable_frame, text="Vertical", variable=self.layout_mode, value="VERTICAL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="Horizontal", variable=self.layout_mode, value="HORIZONTAL").pack(anchor='w', padx=20)
        ttk.Label(scrollable_frame, text="SPC LINE TREND WIDTH:", font=('Arial', 12)).pack(pady=10)
        ttk.Radiobutton(scrollable_frame, text="EXTRA SMALL", variable=self.line_width_var, value="EXTRA SMALL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="SMALL", variable=self.line_width_var, value="SMALL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="MEDIUM", variable=self.line_width_var, value="MEDIUM").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="LARGE", variable=self.line_width_var, value="LARGE").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="EXTRA LARGE", variable=self.line_width_var, value="EXTRA LARGE").pack(anchor='w', padx=20)
        ttk.Label(scrollable_frame, text="SAMPLE NUMBER:", font=('Arial', 12)).pack(pady=10)
        self.sample_num_var = tk.StringVar()
        if self.num_ticks in [3,5]:
            self.sample_num_var.set(str(self.num_ticks))
        else:
            self.sample_num_var.set("others")
        ttk.Radiobutton(scrollable_frame, text="3", variable=self.sample_num_var, value="3").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="5", variable=self.sample_num_var, value="5").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="Others", variable=self.sample_num_var, value="others").pack(anchor='w', padx=20)
        self.sample_other_frame = ttk.Frame(scrollable_frame)
        ttk.Label(self.sample_other_frame, text="Enter Number:").pack(side=tk.LEFT)
        self.sample_other_entry = ttk.Entry(self.sample_other_frame, width=5)
        self.sample_other_entry.insert(0, str(self.num_ticks))
        self.sample_other_entry.pack(side=tk.LEFT)
        self.sample_num_var.trace("w", self.on_sample_change)
        if self.sample_num_var.get() == "others":
            self.sample_other_frame.pack(anchor='w', padx=20, pady=5)
        ttk.Button(scrollable_frame, text="Apply", command=self.apply_display_settings).pack(pady=10)
        ttk.Button(scrollable_frame, text="Back", command=self.display_window.destroy).pack(pady=5)

    def on_sample_change(self, *args):
        if self.sample_num_var.get() == "others":
            self.sample_other_frame.pack(anchor='w', padx=20, pady=5)
            self.sample_other_entry.focus()
        else:
            self.sample_other_frame.pack_forget()

    def apply_display_settings(self):
        #self.update_line_graph()
        self.rebuild_spc_graphs()
        self.update_spc_graphs()
        if hasattr(self, 'focus_window') and self.focus_window.winfo_exists():
            current_meas = [m for m in self.measurement_buttons if self.measurement_buttons[m]['relief'] == 'sunken']
            if current_meas:
                self.show_focus_graph(current_meas[0])
        sample_str = self.sample_num_var.get()
        if sample_str == "others":
            sample_str = self.sample_other_entry.get().strip()
        try:
            self.num_ticks = int(sample_str)
            if sample_str in ["3", "5"]:
                self.tick_strategy = "last"
            else:
                self.tick_strategy = "full"
        except ValueError:
            pass
        self.display_window.destroy()

    def reset_log(self):
        self.log_text.delete('1.0', tk.END)
        open(self.log_path, 'w').close()

    def add_to_log(self, row, count):
        tol_percent = self.threshold * 100
        log_content = f"SERIAL NO.: {row['S/N']}       DATE: {row['DATE']}      TIME: {row['TIME']}      MODEL CODE: {row['MODEL CODE']}\n"
        log_content += "{:<25} {:<15} {:<15}\n".format("PROCESS INSPECTION:", "VALUE:", "TOLERANCE:")
        for name, (value_col, ave_col, dev_col, fluct_col) in self.measurements_map.items():
            if name in self.selected_measurements and row[fluct_col] > self.threshold:
                ave = row[ave_col]
                dev = row[dev_col]
                current = row[value_col]
                upper_limit = ave * (1 + self.threshold)
                lower_limit = ave * (1 - self.threshold)
                if current > upper_limit:
                    sign = "UCL"
                    limit_value = upper_limit
                elif current < lower_limit:
                    sign = "LCL"
                    limit_value = lower_limit
                else:
                    continue
                log_content += "{:<25} {:<15.2f} {:<5}:{:.2f}\n".format("CLOSED P Middle (kPa)" if name == "CLOSED PRESSURE Middle (kPa)" else name, current, sign, limit_value)
        self.log_text.insert(tk.END, log_content + "\n")
        self.log_text.see(tk.END)
        with open(self.log_path, 'a', encoding='utf-8') as f:
            f.write(log_content + "\n")
        if False:
            if not self.model_summary.empty and self.current_model in self.model_summary['MODEL CODE'].values.tolist():
                model_row = self.model_summary[self.model_summary['MODEL CODE'] == self.current_model].iloc[0]
                avg_content = "MODEL AVERAGES:\n"
                avg_content += "{:<40} {:<15}\n".format("MEASUREMENT:", "AVERAGE:")
                for name, key in self.avgs_map.items():
                    if name in self.selected_measurements:
                        avg_content += "{:<40} {:<15.2f}\n".format(name,model_row[key])
                self.log_text.insert(tk.END, avg_content + "\n\n")
                with open(self.log_path, 'a', encoding='utf-8') as f:
                    f.write(avg_content + "\n\n")
            model_df = self.compiledFrame[(self.compiledFrame['MODEL CODE'] == self.current_model) & (self.compiledFrame['PASS/NG'] == 1)]
            dev_content = "AVERAGE DEVIATION VALUES:\n"
            dev_content += "{:<40} {:<15}\n".format("MEASUREMENT:", "AVG DEVIATION (%):")
            for name, col in self.devs_map.items():
                if name in self.selected_measurements:
                    avg_dev = model_df[col].mean().mean() * 100 if not model_df.empty else 0
                    dev_content += "{:<40} {:<15.2f}\n".format(name, avg_dev)
            self.log_text.insert(tk.END, dev_content + "\n\n")
            with open(self.log_path, 'a', encoding='utf-8') as f:
                f.write(dev_content + "\n\n")
            tol_content = f"TOLERANCE: {self.threshold * 100:.1f}%\n\n"
            self.log_text.insert(tk.END, tol_content)
            with open(self.log_path, 'a', encoding='utf-8') as f:
                f.write(tol_content)

    def open_ticks_window(self):
        ticks_window = tk.Toplevel(self.root)
        ticks_window.title("X-Axis Settings")
        ticks_window.geometry(f"{int(self.root.winfo_screenwidth() * 0.3)}x{int(self.root.winfo_screenheight() * 0.3)}")
        ttk.Label(ticks_window, text="SELECT HOW HOW MANY SAMPLE NUMBER WOULD YOU LIKE:", font=('Arial', 12)).pack(pady=10)
        self.ticks_var = tk.StringVar(value="10")
        ttk.Radiobutton(ticks_window, text="3", variable=self.ticks_var, value="3").pack(anchor='w', padx=20)
        ttk.Radiobutton(ticks_window, text="5", variable=self.ticks_var, value="5").pack(anchor='w', padx=20)
        ttk.Radiobutton(ticks_window, text="Others", variable=self.ticks_var, value="others").pack(anchor='w', padx=20)
        self.ticks_other_frame = ttk.Frame(ticks_window)
        ttk.Label(self.ticks_other_frame, text="Enter Number:").pack(side=tk.LEFT)
        self.ticks_other_entry = ttk.Entry(self.ticks_other_frame, width=5)
        self.ticks_other_entry.insert(0, "10")
        self.ticks_other_entry.pack(side=tk.LEFT)
        self.ticks_var.trace("w", self.on_ticks_change)
        ttk.Button(ticks_window, text="Confirm", command=lambda: self.confirm_ticks_selection(ticks_window)).pack(pady=10)

    def on_ticks_change(self, *args):
        if self.ticks_var.get() == "others":
            self.ticks_other_frame.pack(anchor='w', padx=20, pady=5)
            self.ticks_other_entry.focus()
        else:
            self.ticks_other_frame.pack_forget()

    def confirm_ticks_selection(self, ticks_window):
        ticks_str = self.ticks_var.get()
        if ticks_str == "others":
            ticks_str = self.ticks_other_entry.get().strip()
        try:
            self.num_ticks = int(ticks_str)
            if ticks_str in ["3", "5"]:
                self.tick_strategy = "last"
            else:
                self.tick_strategy = "full"
            self.rebuild_spc_graphs()
            self.update_spc_graphs()
        except ValueError:
            pass
        ticks_window.destroy()

    def open_computation_window(self):
        comp_window = tk.Toplevel(self.root)
        comp_window.title("Computation")
        comp_window.geometry(f"{int(self.root.winfo_screenwidth() * 0.5)}x{int(self.root.winfo_screenheight() * 0.5)}")
        tree = ttk.Treeview(comp_window, columns=("PROCESS INSPECTION", "AVE", "(+) TOL", "(-) TOL", "UCL", "LCL"), show="headings")
        tree.heading("PROCESS INSPECTION", text="PROCESS INSPECTION")
        tree.heading("AVE", text="AVE")
        tree.heading("(+) TOL", text="(+) TOL")
        tree.heading("(-) TOL", text="(-) TOL")
        tree.heading("UCL", text="UCL")
        tree.heading("LCL", text="LCL")
        tree.pack(fill=tk.BOTH, expand=True)
        df = self.compiledFrame[self.compiledFrame["PASS/NG"] == 1].copy()
        if self.current_model:
            df = df[df['MODEL CODE'] == self.current_model].copy()
        df['DATE'] = pd.to_datetime(df['DATE'], format='%Y/%m/%d', errors='coerce')
        today = pd.to_datetime(datetime.now().date())
        df = df[df['DATE'].dt.date < today.date()]
        df = df.tail(200)
        for meas in self.mid_measurements:
            data = df[meas]
            if pd.isna(data).all():
                continue
            mean_val = data.mean()
            plus_tol = 1 + self.threshold
            minus_tol = 1 - self.threshold
            ucl = mean_val * plus_tol
            lcl = mean_val * minus_tol
            proc_name = "CLOSED P Middle (kPa)" if meas == "CLOSED PRESSURE Middle (kPa)" else meas
            tree.insert("", "end", values=(proc_name, f"{mean_val:.2f}", f"{plus_tol:.2f}", f"{minus_tol:.2f}", f"{ucl:.2f}", f"{lcl:.2f}"))
        ttk.Button(comp_window, text="Close", command=comp_window.destroy).pack(pady=10)

    def open_ltg_history_window(self):
        date_window = tk.Toplevel(self.root)
        date_window.title("Date Selection")
        date_window.geometry(f"{int(self.root.winfo_screenwidth() * 0.3)}x{int(self.root.winfo_screenheight() * 0.6)}")
        main_container = ttk.Frame(date_window)
        main_container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(main_container)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        ttk.Label(scrollable_frame, text="DATE SELECTION:").pack(pady=10)
        start_frame = ttk.Frame(scrollable_frame)
        start_frame.pack(pady=5)
        ttk.Label(start_frame, text="START DATE:").pack(side=tk.LEFT)
        self.start_date_entry = ttk.Entry(start_frame)
        self.start_date_entry.pack(side=tk.LEFT)
        end_frame = ttk.Frame(scrollable_frame)
        end_frame.pack(pady=5)
        ttk.Label(end_frame, text="END DATE:").pack(side=tk.LEFT)
        self.end_date_entry = ttk.Entry(end_frame)
        self.end_date_entry.pack(side=tk.LEFT)
        self.all_var = tk.BooleanVar()
        ttk.Checkbutton(scrollable_frame, text="ALL (START AT NOV 2024 ONWARDS)", variable=self.all_var).pack(pady=5)
        ttk.Label(scrollable_frame, text="X-AXIS:", font=('Arial', 12)).pack(pady=10)
        self.history_x_axis_mode = tk.StringVar(value="Numerical")
        ttk.Radiobutton(scrollable_frame, text="Numerical Numbering", variable=self.history_x_axis_mode, value="Numerical").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="Month/Year", variable=self.history_x_axis_mode, value="Month/Year").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="Serial No.", variable=self.history_x_axis_mode, value="Serial No.").pack(anchor='w', padx=20)
        self.history_show_title_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(scrollable_frame, text="Show Title", variable=self.history_show_title_var).pack(anchor='w', padx=20, pady=5)
        ttk.Label(scrollable_frame, text="LAYOUT:", font=('Arial', 12)).pack(pady=10)
        self.history_layout_mode = tk.StringVar(value="HORIZONTAL")
        ttk.Radiobutton(scrollable_frame, text="HORIZONTAL", variable=self.history_layout_mode, value="HORIZONTAL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="VERTICAL", variable=self.history_layout_mode, value="VERTICAL").pack(anchor='w', padx=20)
        ttk.Label(scrollable_frame, text="SPC LINE TREND WIDTH", font=('Arial', 12)).pack(pady=10)
        self.history_line_width_var = tk.StringVar(value="MEDIUM")
        ttk.Radiobutton(scrollable_frame, text="EXTRA EXTRA SMALL", variable=self.history_line_width_var, value="EXTRA EXTRA SMALL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="EXTRA SMALL", variable=self.history_line_width_var, value="EXTRA SMALL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="SMALL", variable=self.history_line_width_var, value="SMALL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="MEDIUM", variable=self.history_line_width_var, value="MEDIUM").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="OTHERS", variable=self.history_line_width_var, value="OTHERS").pack(anchor='w', padx=20)
        self.history_line_other_frame = ttk.Frame(scrollable_frame)
        ttk.Label(self.history_line_other_frame, text="Enter Width:").pack(side=tk.LEFT)
        self.history_line_other_entry = ttk.Entry(self.history_line_other_frame, width=5)
        self.history_line_other_entry.insert(0, "3.0")
        self.history_line_other_entry.pack(side=tk.LEFT)
        self.history_line_width_var.trace("w", self.on_history_line_change)
        if self.history_line_width_var.get() == "OTHERS":
            self.history_line_other_frame.pack(anchor='w', padx=20, pady=5)
        ttk.Label(scrollable_frame, text="DATAPOINTS", font=('Arial', 12)).pack(pady=10)
        self.history_datapoints_var = tk.StringVar(value="Scatter")
        ttk.Radiobutton(scrollable_frame, text="Scatter", variable=self.history_datapoints_var, value="Scatter").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="None", variable=self.history_datapoints_var, value="None").pack(anchor='w', padx=20)
        ttk.Label(scrollable_frame, text="SCATTER COLOR:", font=('Arial', 12)).pack(pady=10)
        self.history_scatter_color_var = tk.StringVar(value="Fluctuation Red")
        ttk.Radiobutton(scrollable_frame, text="Normal Blue", variable=self.history_scatter_color_var, value="Normal Blue").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="Fluctuation Red", variable=self.history_scatter_color_var, value="Fluctuation Red").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="None", variable=self.history_scatter_color_var, value="None").pack(anchor='w', padx=20)
        ttk.Label(scrollable_frame, text="LIMITATIONS", font=('Arial', 12)).pack(pady=10)
        self.history_limitations_var = tk.StringVar(value="UCL/LCL")
        ttk.Radiobutton(scrollable_frame, text="UCL/LCL", variable=self.history_limitations_var, value="UCL/LCL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="USL/LSL", variable=self.history_limitations_var, value="USL/LSL").pack(anchor='w', padx=20)
        ttk.Radiobutton(scrollable_frame, text="None", variable=self.history_limitations_var, value="None").pack(anchor='w', padx=20)
        ttk.Button(scrollable_frame, text="CONFIRM", command=lambda: self.show_ltg_history(date_window)).pack(pady=10)

    def on_history_line_change(self, *args):
        if self.history_line_width_var.get() == "OTHERS":
            self.history_line_other_frame.pack(anchor='w', padx=20, pady=5)
            self.history_line_other_entry.focus()
        else:
            self.history_line_other_frame.pack_forget()

    def show_ltg_history(self, date_window):
        if self.all_var.get():
            start_date = '2024-11-01'
            end_date = datetime.now().strftime('%Y-%m-%d')
        else:
            start_date = self.start_date_entry.get()
            end_date = self.end_date_entry.get()
        try:
            start_dt = pd.to_datetime(start_date)
            end_dt = pd.to_datetime(end_date)
        except:
            logging.error("Invalid date format")
            return
        date_window.destroy()
        history_window = tk.Toplevel(self.root)
        history_window.title("LTG History Graphs")
        history_window.geometry(f"{int(self.root.winfo_screenwidth() * 0.8)}x{int(self.root.winfo_screenheight() * 0.8)}")
        main_container = ttk.Frame(history_window)
        main_container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(main_container)
        scrollbar_y = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollbar_x = ttk.Scrollbar(main_container, orient="horizontal", command=canvas.xview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        figs = {}
        axs = {}
        canvases = {}
        df = self.compiledFrame[self.compiledFrame["PASS/NG"] == 1].copy()
        if self.current_model:
            df = df[df['MODEL CODE'] == self.current_model].copy()
        df['DATETIME'] = pd.to_datetime(df['DATETIME'], errors='coerce')
        filtered_df = df[(df['DATETIME'] >= start_dt) & (df['DATETIME'] <= end_dt)]
        filtered_df = filtered_df.sort_values('DATETIME')
        history_line_width_map = {
            "EXTRA EXTRA SMALL": 0.25,
            "EXTRA SMALL": 0.5,
            "SMALL": 1.0,
            "MEDIUM": 2.0,
            "OTHERS": float(self.history_line_other_entry.get()) if self.history_line_width_var.get() == "OTHERS" else 3.0
        }
        history_usl_lsl_map = {
            "VOLTAGE Middle (V)": {"USL": 11.7, "LSL": None},
            "WATTAGE Middle (W)": {"USL": 27.1, "LSL": None},
            "AMPERAGE Middle (A)": {"USL": 3.7, "LSL": None},
            "CLOSED PRESSURE Middle (kPa)": {"USL": 33.2, "LSL": 27.8}
        }
        line_width = history_line_width_map[self.history_line_width_var.get()]
        layout_mode = self.history_layout_mode.get()
        if layout_mode == "HORIZONTAL":
            row1 = ttk.Frame(scrollable_frame)
            row1.pack(fill=tk.BOTH, expand=True, pady=5)
            row2 = ttk.Frame(scrollable_frame)
            row2.pack(fill=tk.BOTH, expand=True, pady=5)
            row_frames = [row1, row2]
            meas_chunks = [self.mid_measurements[:2], self.mid_measurements[2:]]
            for i, row_meas in enumerate(meas_chunks):
                row_frame = row_frames[i]
                for meas in row_meas:
                    sub_frame = ttk.Frame(row_frame)
                    sub_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
                    fig = Figure(figsize=(20, 3))
                    ax = fig.add_subplot(111)
                    fig_canvas = FigureCanvasTkAgg(fig, master=sub_frame)
                    fig_canvas.draw()
                    fig_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                    figs[meas] = fig
                    axs[meas] = ax
                    canvases[meas] = fig_canvas
        else:
            for meas in self.mid_measurements:
                sub_frame = ttk.Frame(scrollable_frame)
                sub_frame.pack(fill=tk.BOTH, expand=True, pady=5)
                fig = Figure(figsize=(20, 3))
                ax = fig.add_subplot(111)
                fig_canvas = FigureCanvasTkAgg(fig, master=sub_frame)
                fig_canvas.draw()
                fig_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                figs[meas] = fig
                axs[meas] = ax
                canvases[meas] = fig_canvas
        if not filtered_df.empty:
            num_ticks = min(10, len(filtered_df))
            mean_val = filtered_df[self.mid_measurements].mean()
            for meas in self.mid_measurements:
                data = filtered_df[meas]
                if pd.isna(data).all():
                    continue
                ax = axs[meas]
                x = np.arange(len(data))
                y = data.values
                if self.history_datapoints_var.get() == "Scatter":
                    if self.history_scatter_color_var.get() == "None":
                        ax.plot(x, y, marker=None, linestyle='-', color='blue', linewidth=line_width)
                    else:
                        ax.plot(x, y, marker=None, linestyle='-', color='blue', linewidth=line_width)
                        colors = np.array(['blue'] * len(y))
                        out_mask = None
                        if self.history_limitations_var.get() == "UCL/LCL":
                            upper_limit = mean_val[meas] * (1 + self.threshold)
                            lower_limit = mean_val[meas] * (1 - self.threshold)
                            out_mask = (y > upper_limit) | (y < lower_limit)
                        elif self.history_limitations_var.get() == "USL/LSL":
                            usl_lsl = history_usl_lsl_map.get(meas, {"USL": None, "LSL": None})
                            upper_limit = usl_lsl["USL"]
                            lower_limit = usl_lsl["LSL"]
                            out_mask = ((y > upper_limit) if upper_limit is not None else False) | ((y < lower_limit) if lower_limit is not None else False)
                        else:
                            out_mask = np.array([False] * len(y))
                        if self.history_scatter_color_var.get() == "Fluctuation Red":
                            colors[out_mask] = 'red'
                        elif self.history_scatter_color_var.get() == "Normal Blue":
                            colors = np.array(['blue'] * len(y))
                        ax.scatter(x, y, color=colors, s=100, zorder=5)
                else:
                    ax.plot(x, y, marker=None, linestyle='-', color='blue', linewidth=line_width)
                ax.axhline(mean_val[meas], color='green', linestyle='--', linewidth=line_width)
                if self.history_limitations_var.get() == "UCL/LCL":
                    upper_limit = mean_val[meas] * (1 + self.threshold)
                    lower_limit = mean_val[meas] * (1 - self.threshold)
                    ax.axhline(upper_limit, color='red', linestyle='--', linewidth=line_width)
                    ax.axhline(lower_limit, color='red', linestyle='--', linewidth=line_width)
                elif self.history_limitations_var.get() == "USL/LSL":
                    usl_lsl = history_usl_lsl_map.get(meas, {"USL": None, "LSL": None})
                    if usl_lsl["USL"] is not None:
                        ax.axhline(usl_lsl["USL"], color='red', linestyle='--', linewidth=line_width)
                    if usl_lsl["LSL"] is not None:
                        ax.axhline(usl_lsl["LSL"], color='red', linestyle='--', linewidth=line_width)
                if self.history_show_title_var.get():
                    ax.set_title(meas)
                else:
                    ax.set_title('')
                if num_ticks > 0:
                    tick_positions = np.linspace(0, len(y)-1, num_ticks, dtype=int)
                    if self.history_x_axis_mode.get() == "Numerical":
                        tick_labels = [str(pos + 1) for pos in tick_positions]
                        ax.set_xlabel('Sample Number')
                    elif self.history_x_axis_mode.get() == "Month/Year":
                        tick_labels = filtered_df['DATETIME'].iloc[tick_positions].dt.strftime('%m/%Y')
                        ax.set_xlabel('Month/Year')
                    elif self.history_x_axis_mode.get() == "Serial No.":
                        tick_labels = filtered_df['S/N'].iloc[tick_positions]
                        ax.set_xlabel('Serial No.')
                    ax.set_xticks(tick_positions)
                    ax.set_xticklabels(tick_labels, rotation=45, ha='right')
                ax.set_ylabel(meas)
                ax.grid(True, linestyle='--', alpha=0.7)
                figs[meas].subplots_adjust(left=0.10, right=0.95, bottom=0.1, top=0.95)
                canvases[meas].draw()
        ttk.Button(history_window, text="Back", command=history_window.destroy).pack(pady=10)

# Start Tkinter main loop
root = tk.Tk()
DatabaseSelection(root)
root.mainloop()
# AI