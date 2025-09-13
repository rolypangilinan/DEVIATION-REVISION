
# CREATE HORIZONTAL SCROLLBAR
# ADJUST THE HISTORICAL DEVIATION VALUES DUE TO THE THRESHOLD (5.0%) IS NOT READABLE
# ERROR - Error processing file: 'MODEL CODE' (BUT STILL WORKING)

import pandas as pd
import numpy as np
import time
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk
from watchdog.observers import Observer
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
# file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledPiMachine\CompiledPIMachine.csv"  #ORIGINAL PATH
file_path = r"C:\Users\ai.pc\OneDrive\Desktop\CompiledPIMachine.csv"  # FALSE TEST
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
        self.root.geometry(f"{int(self.root.winfo_screenwidth() * 0.3)}x{int(self.root.winfo_screenheight() * 0.3)}")
        ttk.Label(root, text="Choose Database Location:", font=('Arial', 12)).pack(pady=10)
        self.location_var = tk.StringVar()
        locations = [
            ("FC1", "FC1"),
            ("TESTING", "TESTING")
        ]
        for text, location in locations:
            ttk.Radiobutton(root, text=text, variable=self.location_var,
                          value=location).pack(anchor='w', padx=40)
        ttk.Label(root, text="Choose Fluctuation Threshold (%):", font=('Arial', 12)).pack(pady=10)
        self.threshold_var = tk.StringVar(value="3")
        threshold_options = [
            ("3%", "3"),
            ("5%", "5"),
            ("Others", "others")
        ]
        for text, val in threshold_options:
            ttk.Radiobutton(root, text=text, variable=self.threshold_var, value=val).pack(anchor='w', padx=40)
        self.other_frame = ttk.Frame(root)
        ttk.Label(self.other_frame, text="Enter %:").pack(side=tk.LEFT)
        self.other_entry = ttk.Entry(self.other_frame, width=5)
        self.other_frame.pack(side=tk.LEFT)
        self.other_entry.insert(0, "3")
        self.threshold_var.trace("w", self.on_threshold_change)
        ttk.Button(root, text="Confirm", command=self.confirm_selection).pack(pady=10)
    def on_threshold_change(self, *args):
        if self.threshold_var.get() == "others":
            self.other_frame.pack(anchor='w', padx=40, pady=5)
            self.other_entry.focus()
        else:
            self.other_frame.pack_forget()
    def confirm_selection(self):
        location = self.location_var.get()
        threshold_str = self.threshold_var.get()
        if threshold_str == "others":
            threshold_str = self.other_entry.get().strip()
        try:
            threshold = float(threshold_str)
        except ValueError:
            threshold = 3.0
        if location:
            self.root.destroy()
            root = tk.Tk()
            app = FluctuationMonitor(root, location, threshold)
            root.protocol("WM_DELETE_WINDOW", app.on_closing)
            root.mainloop()

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, callback):
        self.callback = callback
    def on_modified(self, event):
        if event.src_path.endswith('.csv'):
            self.callback()

class FluctuationMonitor:
    def __init__(self, root, location, threshold=3.0):
        self.root = root
        self.root.title(f"Historical Measurement Trend - {os.path.basename(__file__)}")
        self.root.geometry(f"{int(self.root.winfo_screenwidth() * 0.8)}x{int(self.root.winfo_screenheight() * 0.8)}")
        self.zoom_level = 1.0
        self.threshold = threshold / 100
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
        self.current_model = None
        self.last_date = None
        self.last_good_values = {}
        self.last_good_serial = None
        self.fluctuation_count = 0
        self.last_row_count = 0
        if location == "FC1":
            # self.file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledPiMachine\CompiledPIMachine.csv" # ORIGINAL PATH
            self.file_path = r"C:\Users\ai.pc\OneDrive\Desktop\CompiledPIMachine.csv"  # FALSE TEST
            self.output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_CSV.csv"
        elif location == "TESTING":
            self.file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledPiMachine\CompiledPIMachine.csv"
            self.output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_CSV.csv"
        else:
            self.file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledPiMachine\CompiledPIMachine.csv"
            self.output_path = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\Deviation Original\Deviation_CSV.csv"
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
        if os.path.exists(self.output_path):
            try:
                self.compiledFrame = pd.read_csv(self.output_path, encoding='utf-8-sig', dtype=dtypes, on_bad_lines='skip')
                self.compiledFrame['DATETIME'] = pd.to_datetime(self.compiledFrame['DATETIME'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
            except Exception as e:
                logging.error(f"Error loading existing CSV: {e}")
                self.compiledFrame = pd.DataFrame(columns=emptyColumn)
        else:
            self.compiledFrame = pd.DataFrame(columns=emptyColumn)
        self.main_frame = ttk.Frame(self.scrollable_frame)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X)
        self.left_frame = ttk.Frame(self.top_frame)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(self.left_frame, text="Deviation Status Monitor", 
                 font=('Arial', 14, 'bold')).pack(pady=5)
        ttk.Label(self.left_frame, text=f"Location: {location}", 
                 font=('Arial', 12)).pack(pady=2)
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
        self.details_frame = ttk.Frame(self.top_frame)
        self.details_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        self.create_section("MIN Measurements", [
            "DEV V_MAX PASS FLUCTUATED",
            "DEV WATTAGE MAX (W) FLUCTUATED",
            "DEV CLOSED PRESSURE_MAX (kPa) FLUCTUATED"
        ])
        ttk.Separator(self.details_frame, orient='horizontal').pack(fill='x', pady=5)
        self.create_section("MID Measurements", [
            "DEV VOLTAGE Middle (V) FLUCTUATED",
            "DEV WATTAGE Middle (W) FLUCTUATED",
            "DEV AMPERAGE Middle (A) FLUCTUATED",
            "DEV CLOSED PRESSURE Middle (kPa) FLUCTUATED"
        ])
        ttk.Separator(self.details_frame, orient='horizontal').pack(fill='x', pady=5)
        self.create_section("MAX Measurements", [
            "DEV VOLTAGE MIN (V) FLUCTUATED",
            "DEV WATTAGE MIN (W) FLUCTUATED",
            "DEV CLOSED PRESSURE MIN (kPa) FLUCTUATED"
        ])
        ttk.Separator(self.details_frame, orient='horizontal').pack(fill='x', pady=5)
        self.right_frame = ttk.Frame(self.top_frame)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.create_bar_graph()
        self.line_graph_frame = ttk.Frame(self.main_frame)
        self.line_graph_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.create_line_graph()
        self.combo_frame = ttk.Frame(self.main_frame)
        self.combo_frame.pack(pady=5)
        ttk.Label(self.combo_frame, text="Select Model Code:").pack(side=tk.LEFT)
        self.combo = Combobox(self.combo_frame)
        self.combo.pack(side=tk.LEFT)
        self.combo.bind("<<ComboboxSelected>>", self.on_select)
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=5)
        ttk.Button(button_frame, text="Refresh Values", 
                  command=self.process_and_update).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Reset All", 
                  command=self.reset_all_fluctuations).pack(side=tk.LEFT, padx=2)
        self.last_modified = os.path.getmtime(self.file_path)
        event_handler = FileChangeHandler(self.check_file_update)
        self.observer = Observer()
        self.observer.schedule(event_handler, os.path.dirname(self.file_path))
        self.observer.start()
        self.process_and_update()
        self.root.after(1000, self.periodic_check)
        self.old_width = None
        self.old_height = None
        self.root.bind("<Configure>", self.on_resize)
        self.root.update_idletasks()
        fake_event = type('Event', (), {'width': self.root.winfo_width(), 'height': self.root.winfo_height(), 'widget': self.root})()
        self.on_resize(fake_event)

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
                if isinstance(widget, ttk.Label) and widget.cget('text') in ["MIN Measurements", "MID Measurements", "MAX Measurements"]:
                    new_size = int(12 * self.zoom_level)
                    widget.config(font=('Arial', new_size, 'bold'))
        if hasattr(self, 'fig'):
            self.fig.set_size_inches(6 * self.zoom_level, 4 * self.zoom_level)
            self.canvas_graph.draw()
        if hasattr(self, 'line_fig'):
            window_width = self.root.winfo_width()
            window_height = self.root.winfo_height()
            fig_width = max(8, window_width / 100 * 0.95) * self.zoom_level
            fig_height = max(3, window_height / 100 * 0.4) * self.zoom_level
            self.line_fig.set_size_inches(fig_width, fig_height)
            self.line_canvas.draw()

    def create_bar_graph(self):
        self.fig, self.ax = plt.subplots(figsize=(6, 3.9))
        self.ax.set_title('Current Deviation Values')
        self.ax.set_ylabel('Fluctuation Amount (%)')
        self.ax.grid(True)
        self.fluctuation_measurements = [
            'V_MAX', 'W_MAX', 'CP_MAX',
            'V_MID', 'W_MID', 'A_MID', 'CP_MID',
            'V_MIN', 'W_MIN', 'CP_MIN'
        ]
        self.fluctuation_values = [0] * len(self.fluctuation_measurements)
        self.bars = self.ax.bar(self.fluctuation_measurements, self.fluctuation_values, color='skyblue')
        self.ax.set_xticks(range(len(self.fluctuation_measurements)))
        self.ax.set_xticklabels(self.fluctuation_measurements, rotation=45, ha='right', fontsize=7)
        self.ax.axhline(y=self.threshold * 100, color='r', linestyle='--', linewidth=1)
        self.ax.axhline(y=-self.threshold * 100, color='r', linestyle='--', linewidth=1)
        self.ax.text(1.02, self.threshold * 100, f'Threshold ({self.threshold * 100:.1f}%)', color='r', va='bottom', ha='left', 
                     transform=self.ax.get_yaxis_transform(), fontsize=9, 
                     bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.3', alpha=0.9))
        self.ax.text(1.02, -self.threshold * 100, f'Threshold (-{self.threshold * 100:.1f}%)', color='r', va='top', ha='left', 
                     transform=self.ax.get_yaxis_transform(), fontsize=9, 
                     bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.3', alpha=0.9))
        self.fig.subplots_adjust(left=0.1, right=0.80, top=0.9, bottom=0.25)
        self.canvas_graph = FigureCanvasTkAgg(self.fig, master=self.right_frame)
        self.canvas_graph.draw()
        self.canvas_graph.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def create_line_graph(self):
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        fig_width = max(8, window_width / 100 * 0.95)
        fig_height = max(3, window_height / 100 * 0.9)
        self.line_fig = Figure(figsize=(fig_width, fig_height), dpi=100)
        self.line_ax = self.line_fig.add_subplot(111)
        self.line_ax.set_title('Historical Measurement Trends')
        self.line_ax.set_ylabel('Value')
        self.line_ax.grid(True)
        self.measurements = [
            "VOLTAGE MAX (V)", "WATTAGE MAX (W)", "CLOSED PRESSURE_MAX (kPa)",
            "VOLTAGE Middle (V)", "WATTAGE Middle (W)", "AMPERAGE Middle (A)", "CLOSED PRESSURE Middle (kPa)",
            "VOLTAGE MIN (V)", "WATTAGE MIN (W)", "CLOSED PRESSURE MIN (kPa)"
        ]
        self.colors = [
            'blue', 'green', 'red', 'purple', 'orange',
            'brown', 'pink', 'gray', 'cyan', 'magenta'
        ]
        for idx, measurement in enumerate(self.measurements):
            self.line_ax.plot([], [], label=measurement, color=self.colors[idx], linewidth=1.5)
        legend = self.line_ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
        for line in legend.get_lines():
            line.set_linewidth(4.0)
        self.line_ax.tick_params(axis='x', rotation=45, labelsize=8)
        self.line_fig.subplots_adjust(left=0.1, right=0.75, bottom=0.25, top=0.9)
        self.line_canvas = FigureCanvasTkAgg(self.line_fig, master=self.line_graph_frame)
        self.line_canvas.draw()
        self.line_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def update_line_graph(self):
        global compiledFrame
        try:
            if not hasattr(self, 'compiledFrame') or self.compiledFrame.empty:
                return
            self.line_ax.clear()
            df = self.compiledFrame[self.compiledFrame["PASS/NG"] == 1].copy()
            if self.current_model:
                df = df[df['MODEL CODE'] == self.current_model]
            df = df.tail(200)
            df['DATETIME'] = pd.to_datetime(df['DATETIME'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
            df = df.sort_values('DATETIME')
            if not df.empty:
                x = np.arange(len(df))
                for idx, measurement in enumerate(self.measurements):
                    self.line_ax.plot(x, df[measurement].values, label=measurement, color=self.colors[idx], linewidth=1.5)
                num_ticks = min(10, len(df))
                if num_ticks > 0:
                    tick_positions = np.linspace(0, len(df)-1, num_ticks, dtype=int)
                    tick_labels = df['DATETIME'].iloc[tick_positions].dt.strftime('%Y-%m-%d %H:%M:%S')
                    self.line_ax.set_xticks(tick_positions)
                    self.line_ax.set_xticklabels(tick_labels, rotation=45, ha='right')
            self.line_ax.set_title(f'Historical Measurement Trends for {self.current_model or "All Models"}')
            self.line_ax.set_ylabel('Value')
            self.line_ax.grid(True)
            self.line_ax.tick_params(axis='x', labelsize=8)
            legend = self.line_ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
            for line in legend.get_lines():
                line.set_linewidth(4.0)
            self.line_fig.subplots_adjust(left=0.1, right=0.75, bottom=0.25, top=0.9)
            self.line_canvas.draw()
        except Exception as e:
            logging.error(f"Error updating line graph: {e}")

    def update_bar_graph(self, last_row):
        try:
            self.ax.clear()
            self.fluctuation_values = [
                last_row['DEV V_MAX PASS'] * 100,
                last_row['DEV WATTAGE MAX (W)'] * 100,
                last_row['DEV CLOSED PRESSURE_MAX (kPa)'] * 100,
                last_row['DEV VOLTAGE Middle (V)'] * 100,
                last_row['DEV WATTAGE Middle (W)'] * 100,
                last_row['DEV AMPERAGE Middle (A)'] * 100,
                last_row['DEV CLOSED PRESSURE Middle (kPa)'] * 100,
                last_row['DEV VOLTAGE MIN (V)'] * 100,
                last_row['DEV WATTAGE MIN (W)'] * 100,
                last_row['DEV CLOSED PRESSURE MIN (kPa)'] * 100
            ]
            self.bars = self.ax.bar(self.fluctuation_measurements, self.fluctuation_values, color='skyblue')
            self.ax.set_title('Current Deviation Values')
            self.ax.set_ylabel('Fluctuation Amount (%)')
            self.ax.grid(True)
            self.ax.set_xticks(range(len(self.fluctuation_measurements)))
            self.ax.set_xticklabels(self.fluctuation_measurements, rotation=45, ha='right', fontsize=7)
            max_value = max([abs(v) for v in self.fluctuation_values]) if max([abs(v) for v in self.fluctuation_values]) > 0 else 10
            self.ax.set_ylim(-max_value * 1.1, max_value * 1.1)
            for bar in self.bars:
                height = bar.get_height()
                self.ax.text(bar.get_x() + bar.get_width()/2., height,
                            f'{height:.2f}',
                            ha='center', va='bottom' if height >= 0 else 'top')
            self.ax.axhline(y=self.threshold * 100, color='r', linestyle='--', linewidth=1)
            self.ax.axhline(y=-self.threshold * 100, color='r', linestyle='--', linewidth=1)
            self.ax.text(1.02, self.threshold * 100, f'Threshold ({self.threshold * 100:.1f}%)', color='r', va='bottom', ha='left', 
                         transform=self.ax.get_yaxis_transform(), fontsize=9, 
                         bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.3', alpha=0.9))
            self.ax.text(1.02, -self.threshold * 100, f'Threshold (-{self.threshold * 100:.1f}%)', color='r', va='top', ha='left', 
                         transform=self.ax.get_yaxis_transform(), fontsize=9, 
                         bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.3', alpha=0.9))
            self.fig.subplots_adjust(left=0.1, right=0.80, top=0.9, bottom=0.25)
            self.canvas_graph.draw()
        except Exception as e:
            logging.error(f"Error updating bar graph: {e}")

    def update_status_box(self, has_fluctuation, count=0):
        if has_fluctuation:
            self.status_box.configure(bg="red")
            self.status_box.itemconfig(self.status_text, text="FLUCTUATION DETECTED!")
            self.status_box.itemconfig(self.counter_text, text=f"Fluctuations: {count}/10")
        else:
            self.status_box.configure(bg="green")
            self.status_box.itemconfig(self.status_text, text="NO FLUCTUATION DETECTED")
            self.status_box.itemconfig(self.counter_text, text="Fluctuations: 0/10")

    def periodic_check(self):
        self.check_file_update()
        self.root.after(1000, self.periodic_check)

    def check_file_update(self):
        global stop_dots
        try:
            current_modified = os.path.getmtime(self.file_path)
            if current_modified > self.last_modified:
                with open(self.file_path, 'rb') as f:
                    new_row_count = sum(1 for _ in f) - 1
                self.last_modified = current_modified
                if new_row_count <= self.last_row_count:
                    print("NG PRESSURE OR NO DATA ADDED")
                    return
                self.last_row_count = new_row_count
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
                df = pd.read_csv(self.file_path, encoding='latin1', dtype={'S/N': str, 'MODEL CODE': str})
            except UnicodeDecodeError:
                df = pd.read_csv(self.file_path, encoding='ISO-8859-1', errors='replace', dtype={'S/N': str, 'MODEL CODE': str})
                
            self.last_row_count = len(df)
            df = df[~df["MODEL CODE"].isin(['60CAT0203M', '60CAT0202P', '60CAT0203P', '60FC00000P', 
                                            '60FC00902P', '60FC00903P', '60FC00905P', '60FCXP001P', 
                                            '30FCXP001P'])]
            df = df[df["PASS/NG"].isin([0, 1])]
            df['S/N'] = df['S/N'].astype(str)
            df['MODEL CODE'] = df['MODEL CODE'].astype(str)
            df['DATE'] = pd.to_datetime(df['DATE'], format='%Y/%m/%d', errors='coerce')
            df = df.dropna(subset=['DATE'])
            df = df[df['S/N'].str.len() >= 8]
            df = df[~df['MODEL CODE'].str.contains('M')]
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
            for model, group in df.groupby('MODEL CODE'):
                past_data = group[group['DATE'].dt.date < today.date()]
                if past_data.empty:
                    logging.info(f"Skipping {model}: No past data")
                    continue
                past_data = past_data.sort_values('DATE', ascending=False)
                accumulated_rows = pd.DataFrame()
                count = 0
                for date in past_data['DATE'].dt.date.unique():
                    daily_rows = past_data[past_data['DATE'].dt.date == date]
                    valid_rows = daily_rows[daily_rows["PASS/NG"] == 1]
                    accumulated_rows = pd.concat([accumulated_rows, valid_rows])
                    count += len(valid_rows)
                    if count >= 200:
                        latest_valid_date = date
                        break
                if count < 200:
                    logging.info(f"Skipping {model}: Not enough valid PASS/NG rows")
                    continue
                results.append({
                    'MODEL CODE': model,
                    'LATEST DATE': latest_valid_date,
                    'V-MAX PASS AVG': accumulated_rows["VOLTAGE MAX (V)"].mean(),
                    'WATTAGE MAX AVG': accumulated_rows["WATTAGE MAX (W)"].mean(),
                    'CLOSED PRESSURE_MAX AVG': accumulated_rows["CLOSED PRESSURE_MAX (kPa)"].mean(),
                    'VOLTAGE Middle AVG': accumulated_rows["VOLTAGE Middle (V)"].mean(),
                    'WATTAGE Middle AVG': accumulated_rows["WATTAGE Middle (W)"].mean(),
                    'AMPERAGE Middle AVG': accumulated_rows["AMPERAGE Middle (A)"].mean(),
                    'CLOSED PRESSURE Middle AVG': accumulated_rows["CLOSED PRESSURE Middle (kPa)"].mean(),
                    'VOLTAGE MIN (V) AVG': accumulated_rows["VOLTAGE MIN (V)"].mean(),
                    'WATTAGE MIN AVG': accumulated_rows["WATTAGE MIN (W)"].mean(),
                    'CLOSED PRESSURE MIN AVG': accumulated_rows["CLOSED PRESSURE MIN (kPa)"].mean()
                })
            model_summary = pd.DataFrame(results)
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
            new_df = df[~df['S/N'].isin(existing_sn)].copy()
            if new_df.empty:
                logging.info("No new data to process")
                self.update_combo()
                self.update_display()
                return
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
                    "DATE": row["DATE"],
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
                dataFrame["AVE V_MAX PASS"] = pass_avg_map.get(model_code)
                dataFrame["AVE WATTAGE MAX (W)"] = wattage_avg_map.get(model_code)
                dataFrame["AVE CLOSED PRESSURE_MAX (kPa)"] = closedPressure_avg_map.get(model_code)
                dataFrame["AVE VOLTAGE Middle (V)"] = voltageMiddle_avg_map.get(model_code)
                dataFrame["AVE WATTAGE Middle (W)"] = wattageMiddle_avg.get(model_code)
                dataFrame["AVE AMPERAGE Middle (A)"] = amperageMiddle_avg.get(model_code)
                dataFrame["AVE CLOSED PRESSURE Middle (kPa)"] = closePressureMiddle_avg.get(model_code)
                dataFrame["AVE VOLTAGE MIN (V)"] = voltageMin_avg.get(model_code)
                dataFrame["AVE WATTAGE MIN (W)"] = wattageMin_avg.get(model_code)
                dataFrame["AVE CLOSED PRESSURE MIN (kPa)"] = closePressureMin_avg.get(model_code)
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
                if all(abs(v) <= self.threshold for v in fluctuations.values()):
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
                self.compiledFrame = pd.concat([self.compiledFrame, new_data]).drop_duplicates(subset=['S/N', 'DATETIME'], keep='last').reset_index(drop=True)
                try:
                    self.compiledFrame.to_csv(self.output_path, index=False, encoding='utf-8-sig')
                except Exception as e:
                    logging.error(f"Error saving CSV to {self.output_path}: {e}")
            compiledFrame = self.compiledFrame
            self.update_combo()
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
                    value = last_row[column] * 100
                    status = 1 if abs(value) > self.threshold * 100 else 0
                    self.status_vars[column].set(f"= {status}")
                    color = "red" if status == 1 else "green"
                    self.status_labels[column].configure(foreground=color)
                    if status == 1:
                        has_fluctuation = True
                        fluctuation_count += 1
            self.update_status_box(has_fluctuation, fluctuation_count)
            if "MODEL CODE" in last_row:
                model_code = last_row["MODEL CODE"]
                self.model_display.config(text=f"MODEL CODE: {model_code}")
            self.update_bar_graph(last_row)
            self.update_line_graph()
        except Exception as e:
            logging.error(f"Error updating display: {e}")

    def create_section(self, title, columns):
        frame = ttk.Frame(self.details_frame)
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
            self.status_vars[column_name].set("= 0")
            self.status_labels[column_name].configure(foreground="green")
            last_row = self.compiledFrame.iloc[-1]
            if all(abs(last_row[col]) <= self.threshold for col in self.status_vars if col in last_row):
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
                self.status_vars[column].set("= 0")
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
            try:
                self.compiledFrame.to_csv(self.output_path, index=False, encoding='utf-8-sig')
            except Exception as e:
                logging.error(f"Error saving CSV in reset_all_fluctuations: {e}")
            self.update_status_box(False)
            self.update_display()
        except Exception as e:
            logging.error(f"Error resetting all fluctuations: {e}")

    def on_closing(self):
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
        self.update_line_graph()

    def on_resize(self, event):
        if event.widget == self.root:
            if self.old_width is not None and event.width == self.old_width and event.height == self.old_height:
                return
            self.old_width = event.width
            self.old_height = event.height
            try:
                dpi = self.fig.get_dpi()
                new_width = self.right_frame.winfo_width() / dpi
                new_height = self.right_frame.winfo_height() / dpi
                self.fig.set_size_inches(max(4, new_width * 0.9), max(3, new_height * 0.9))
                self.canvas_graph.draw()
            except:
                pass
            try:
                dpi = self.line_fig.get_dpi()
                new_width = self.line_graph_frame.winfo_width() / dpi
                new_height = self.line_graph_frame.winfo_height() / dpi
                self.line_fig.set_size_inches(max(6, new_width * 0.95), max(4, new_height * 0.95))
                self.line_canvas.draw()
            except:
                pass

# Start Tkinter main loop
root = tk.Tk()
DatabaseSelection(root)
root.mainloop()