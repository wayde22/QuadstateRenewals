import logging
import os
import time
import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk

from .config import (
    get_default_input_file,
    get_default_output_folder,
    load_environment,
)
from .logging_config import configure_logging
from .processor import process_renewals


STATUS_DEFAULT_TEXT_COLOR = ("gray10", "gray90")
STATUS_WARNING_TEXT_COLOR = ("#946200", "#FFD166")


class QuadstateRenewalsApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Quadstate Renewal Processor")
        self.geometry('640x330')
        self.minsize(560, 330)
        self.grid_columnconfigure(0, weight=1)

        logging.debug('Setting default file paths.')
        self.source_var = tk.StringVar(master=self, value=get_default_input_file())
        self.destination_var = tk.StringVar(master=self, value=get_default_output_folder())
        self.status_var = tk.StringVar(master=self, value="Ready - Password loaded from environment")

        self._build_widgets()

    def _build_widgets(self):
        source_label = ctk.CTkLabel(
            self,
            text="Select Source File:",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
            justify="left",
        )
        source_label.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 4))

        source_row = ctk.CTkFrame(self, fg_color="transparent")
        source_row.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 8))
        source_row.grid_columnconfigure(0, weight=1)

        source_entry = ctk.CTkEntry(source_row, textvariable=self.source_var)
        source_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        source_button = ctk.CTkButton(
            source_row,
            text="Browse Source",
            width=140,
            command=self.select_source_file,
        )
        source_button.grid(row=0, column=1)

        destination_label = ctk.CTkLabel(
            self,
            text="Select Destination Folder:",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
            justify="left",
        )
        destination_label.grid(row=2, column=0, sticky="ew", padx=24, pady=(4, 4))

        destination_row = ctk.CTkFrame(self, fg_color="transparent")
        destination_row.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 10))
        destination_row.grid_columnconfigure(0, weight=1)

        destination_entry = ctk.CTkEntry(destination_row, textvariable=self.destination_var)
        destination_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        destination_button = ctk.CTkButton(
            destination_row,
            text="Browse Destination",
            width=160,
            command=self.select_destination_folder,
        )
        destination_button.grid(row=0, column=1)

        self.status_label = ctk.CTkLabel(
            self,
            textvariable=self.status_var,
            text_color=STATUS_DEFAULT_TEXT_COLOR,
            anchor="w",
            justify="left",
        )
        self.status_label.grid(row=4, column=0, sticky="ew", padx=24, pady=(4, 0))

        self.progress_bar = ctk.CTkProgressBar(self, mode='determinate')
        self.progress_bar.set(0)
        self.progress_bar.grid(row=5, column=0, sticky="ew", padx=24, pady=(10, 8))

        self.count_label = ctk.CTkLabel(self, text="Records processed: 0")
        self.count_label.grid(row=6, column=0, pady=(2, 6))

        process_button = ctk.CTkButton(
            self,
            text="Process",
            width=180,
            command=self.process_excel,
        )
        process_button.grid(row=7, column=0, pady=(0, 16))

    def update_count_label(self, record_count, worksheet_row_count):
        self.count_label.configure(
            text=f"Records processed: {record_count} | Excel rows: {worksheet_row_count}"
        )
        logging.info(
            f'Updated count label to: Records processed: {record_count}; Excel rows: {worksheet_row_count}'
        )

    def set_progress(self, percent):
        self.progress_bar.set(percent / 100)
        self.update_idletasks()

    def set_status(self, message):
        self.status_var.set(message)
        if message.startswith("Warning -") or "Source file format changed" in message:
            self.status_label.configure(text_color=STATUS_WARNING_TEXT_COLOR)
        else:
            self.status_label.configure(text_color=STATUS_DEFAULT_TEXT_COLOR)
        self.update_idletasks()

    def process_excel(self):
        result = process_renewals(
            self.source_var.get(),
            self.destination_var.get(),
            on_status=self.set_status,
            on_progress=self.set_progress,
        )

        if result.success:
            self.update_count_label(result.record_count, result.worksheet_row_count)
            self.update_idletasks()
            time.sleep(3)
            self.destroy()

    def select_source_file(self):
        current_source = self.source_var.get()
        initial_dir = (
            os.path.dirname(current_source)
            if current_source and os.path.exists(current_source)
            else os.path.expanduser("~/Downloads")
        )

        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("Excel 2007+ files", "*.xlsx"),
                ("Excel 97-2003 files", "*.xls"),
                ("All files", "*.*"),
            ],
            initialdir=initial_dir,
        )
        if file_path:
            self.source_var.set(file_path)
            logging.debug(f'Source file selected: {file_path}')

    def select_destination_folder(self):
        current_destination = self.destination_var.get()
        initial_dir = (
            current_destination
            if current_destination and os.path.exists(current_destination)
            else os.path.expanduser("~/Desktop")
        )

        folder_path = filedialog.askdirectory(
            title="Select Destination Folder",
            initialdir=initial_dir,
        )
        if folder_path:
            self.destination_var.set(folder_path)
            logging.debug(f'Destination folder selected: {folder_path}')


def run_app():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    app = QuadstateRenewalsApp()
    app.mainloop()


def main():
    load_environment()
    configure_logging()
    run_app()

