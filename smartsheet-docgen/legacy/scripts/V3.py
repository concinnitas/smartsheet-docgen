import os
import subprocess
import sys
import smartsheet
import json
from docx import Document
from tkinter import ttk
from tkinter import filedialog, messagebox
import tkinter as tk
from ttkthemes import ThemedTk
import io
import copy

# Check and install libraries
required_libraries = ["smartsheet", "python-docx", "ttkthemes", "tkinter"]

def install_libraries():
    for lib in required_libraries:
        try:
            globals()[lib] = __import__(lib)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

install_libraries()

# Define paths and global variables
config_file = "config.json"

def save_api_key(api_key):
    config = {"api_key": api_key}
    with open(config_file, 'w') as f:
        json.dump(config, f)

def load_api_key():
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            config = json.load(f)
        return config.get("api_key", "")
    return ""

def lock_api_key_field(entry):
    entry.config(state="disabled")

def save_api_key_to_file(entry):
    api_token = entry.get()
    if api_token:
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, 'w') as file:
                file.write(api_token)
            messagebox.showinfo("Success", "API key saved to file!")
            lock_api_key_field(entry)

def import_config_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("JSON files", "*.json")])
    if file_path:
        with open(file_path, 'r') as file:
            api_key = file.read().strip()
            entry.delete(0, tk.END)
            entry.insert(0, api_key)
            lock_api_key_field(entry)

def fetch_and_populate_sheets(api_token, sheet_id_combobox):
    try:
        ss_client = smartsheet.Smartsheet(api_token)
        response = ss_client.Sheets.list_sheets(include_all=True)
        sheets = response.data
        if sheets:
            sheet_id_combobox['values'] = [sheet.name for sheet in sheets]
            sheet_id_combobox.set(sheets[0].name)
        else:
            messagebox.showerror("Error", "No sheets found.")
    except Exception as e:
        messagebox.showerror("Error", f"Error fetching sheets: {e}")

def browse_directory(entry):
    directory = filedialog.askdirectory()
    if directory:
        entry.delete(0, tk.END)
        entry.insert(0, directory)

def browse_template_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def process_smartsheet_and_generate_doc(api_token, sheet_name, image_dir1, image_dir2, template_file, output_docx, static_cell_mappings, section_cell_mappings):
    try:
        ss_client = smartsheet.Smartsheet(api_token)
        all_sheets = ss_client.Sheets.list_sheets(include_all=True).data
        sheet = next(sheet for sheet in all_sheets if sheet.name == sheet_name)
        sheet_data = ss_client.Sheets.get_sheet(sheet.id)

        doc = Document(template_file)

        template_section = []
        in_template_section = False
        for paragraph in doc.paragraphs:
            if "[SECTION_TEMPLATE]" in paragraph.text:
                in_template_section = True
            if in_template_section:
                template_section.append(paragraph)
                if "End of Section" in paragraph.text:
                    break

        for para in template_section:
            p = para._element
            p.getparent().remove(p)

        # Process static variables
        for row in sheet_data.rows:
            row_cells = {cell.column_id: cell.value for cell in row.cells}

            # Replace static placeholders with values
            for placeholder, (row_idx, col_idx) in static_cell_mappings.items():
                placeholder_tag = f"[{placeholder}]"
                if placeholder_tag in doc.text:
                    value = row_cells.get(col_idx)
                    if value:
                        doc.text = doc.text.replace(placeholder_tag, str(value))

        # Process sectioned data
        for section, (start_row, col_idx) in section_cell_mappings.items():
            for row_idx in range(start_row, start_row + len(sheet_data.rows)):
                row = sheet_data.rows[row_idx]
                section_data = {cell.column_id: cell.value for cell in row.cells}

                for paragraph in doc.paragraphs:
                    for placeholder, col_idx in section_data.items():
                        placeholder_tag = f"[{section}]"
                        if placeholder_tag in paragraph.text:
                            value = section_data.get(col_idx)
                            if value:
                                paragraph.text = paragraph.text.replace(placeholder_tag, str(value))

        doc.save(output_docx)
        messagebox.showinfo("Success", f"Document generated successfully: {output_docx}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def create_ui():
    root = ThemedTk(theme="plastik")
    root.title("SmartSheet Document Generator")
    root.geometry("600x700")

    # Dark mode settings
    root.configure(bg="#2D2D2D")
    default_bg = "#2D2D2D"
    default_fg = "#FFFFFF"
    entry_bg = "#444444"
    entry_fg = "#FFFFFF"
    button_bg = "#555555"
    button_fg = "#FFFFFF"

    style = ttk.Style(root)
    style.configure("TLabel", background=default_bg, foreground=default_fg)
    style.configure("TButton", background=button_bg, foreground=button_fg)
    style.configure("TEntry", fieldbackground=entry_bg, foreground=entry_fg)

    api_token_label = ttk.Label(root, text="API Key").pack(pady=10)
    api_token_entry = ttk.Entry(root, width=40)
    api_token_entry.pack(pady=10)

    # API Key Field Locking (Save, Import)
    api_key_save_button = ttk.Button(root, text="Save API Key", command=lambda: save_api_key_to_file(api_token_entry)).pack(pady=10)
    api_key_import_button = ttk.Button(root, text="Import API Key", command=lambda: import_config_file(api_token_entry)).pack(pady=10)

    sheet_id_label = ttk.Label(root, text="Smartsheet").pack(pady=10)
    sheet_id_combobox = ttk.Combobox(root, width=40, state="readonly")
    sheet_id_combobox.pack(pady=10)

    fetch_sheets_button = ttk.Button(root, text="Fetch Sheets", command=lambda: fetch_and_populate_sheets(api_token_entry.get(), sheet_id_combobox))
    fetch_sheets_button.pack(pady=10)

    # Directory fields for images
    image_dir1_label = ttk.Label(root, text="Image Directory 1").pack(pady=10)
    image_dir1_entry = ttk.Entry(root, width=40)
    image_dir1_entry.pack(pady=10)
    image_dir1_button = ttk.Button(root, text="Browse", command=lambda: browse_directory(image_dir1_entry))
    image_dir1_button.pack(pady=10)

    image_dir2_label = ttk.Label(root, text="Image Directory 2").pack(pady=10)
    image_dir2_entry = ttk.Entry(root, width=40)
    image_dir2_entry.pack(pady=10)
    image_dir2_button = ttk.Button(root, text="Browse", command=lambda: browse_directory(image_dir2_entry))
    image_dir2_button.pack(pady=10)

    # Template file selection
    template_file_label = ttk.Label(root, text="Template File").pack(pady=10)
    template_file_frame = ttk.Frame(root)
    template_file_frame.pack(pady=10)
    template_file_var = tk.StringVar()
    template_file_entry = ttk.Entry(template_file_frame, width=40, textvariable=template_file_var)
    template_file_entry.pack(side=tk.LEFT)
    template_file_button = ttk.Button(template_file_frame, text="Browse", command=lambda: browse_template_file(template_file_entry))
    template_file_button.pack(side=tk.LEFT)

    # Generate document button
    def execute_full_sequence():
        if not api_token_entry.get():
            messagebox.showerror("Error", "API Key is required!")
            return

        if not sheet_id_combobox.get():
            messagebox.showerror("Error", "Smartsheet is required!")
            return

        if not image_dir1_entry.get() or not image_dir2_entry.get():
            messagebox.showerror("Error", "Both image directories are required!")
            return

        if not template_file_var.get():
            messagebox.showerror("Error", "Template file is required!")
            return

        output_directory = filedialog.askdirectory()
        if output_directory:
            # Fetch the project name from Smartsheet data
            project_name = sheet_id_combobox.get()  # Placeholder, adjust as necessary
            output_file = os.path.join(output_directory, f"{project_name}.docx")

            # Static Cell Mappings (row, column indices for each placeholder)
            static_cell_mappings = {
                "PROJECT_NAME": (3, 6),  
                "ADDRESS": (3, 7),  
                "RACKING_TYPE": (3, 8),  
                "RACKING_LOCATION": (3, 9),
                "DC_CAPACITY": (3, 10),
                "AC_CAPACITY": (3, 111),
                "INV_1_COUNT": (3, 12),
                "INV_1_TYPE": (3, 13),
                "INV_2_COUNT": (3, 14),
                "INV_2_TYPE": (3, 15),
                "MOD_COUNT": (3, 16),
                "MOD_TYPE": (3, 17),
                "MLPE_COUNT": (3, 18),
                "MLPE_TYPE": (3, 19),
                "IX_TYPE": (3, 20),
                "IX-VOLTAGE": (3, 21)
            }

            # Section Cell Mappings (section data starts at row 26)
            section_cell_mappings = {
                "Section 1": (26, 2),
                "Section 2": (27, 2),
                "Section 3": (28, 2)
                # Add more sections if needed
            }

            process_smartsheet_and_generate_doc(
                api_token_entry.get(),
                sheet_id_combobox.get(),
                image_dir1_entry.get(),
                image_dir2_entry.get(),
                template_file_var.get(),
                output_file,
                static_cell_mappings,
                section_cell_mappings
            )

    generate_button = ttk.Button(root, text="Generate Document", command=execute_full_sequence)
    generate_button.pack(pady=20)

    root.mainloop()

create_ui()
