import os
import smartsheet
from docx import Document
from docx.shared import Inches
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
from ttkthemes import ThemedTk
import json
import copy
import io

#Namspace Constants
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WORDPROCESSING_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

# Default paths for configuration and mappings
script_dir = os.path.abspath(os.path.dirname(__file__))  # Ensure we use the script directory
config_file = os.path.join(script_dir, "config.json")
mappings_file = os.path.join(script_dir, "mappings.json")

# Load the saved API key from config file
def load_config():
    try:
        if os.path.isfile(config_file):
            with open(config_file, 'r') as file:
                return json.load(file)
        else:
            return {}
    except json.JSONDecodeError as e:
        print(f"Error loading config: {e}")
        return {}

# Save the API key to the config file
def save_config(api_key):
    try:
        config = {"api_key": api_key}
        with open(config_file, 'w') as file:
            json.dump(config, file, indent=4)
        print(f"API key saved successfully to {config_file}")
    except Exception as e:
        print(f"Error saving config: {e}")

# Load JSON mappings
def load_mappings():
    if not os.path.isfile(mappings_file):
        return None
    with open(mappings_file, 'r') as file:
        try:
            return json.load(file)
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Error parsing mapping file: {e}")
            return None

# Save a new JSON mappings file
def save_new_mappings(content=None):
    if not content:
        content = {
            "static_cell_mappings": {},
            "dynamic_section_mappings": {
                "header_text": "[SECTION_START]",
                "fields": {},
                "starting_row": 0,
                "end_row": 0
            }
        }
    with open(mappings_file, 'w') as file:
        json.dump(content, file, indent=4)

# Replace static placeholders in the Word document
def replace_static_placeholders(doc, sheet_data, static_mappings):
    for placeholder, details in static_mappings.items():
        row_index, col_index = details.get("location", [0, 0])
        fallback = details.get("fallback", "")
        try:
            value = sheet_data.rows[row_index].cells[col_index].value
            value = value if value else fallback  # Use fallback if value is None or empty
            for paragraph in doc.paragraphs:
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, str(value))
        except IndexError:
            print(f"Invalid index for placeholder {placeholder}: Row {row_index}, Column {col_index}")
            continue

# Main processing function
def process_smartsheet_and_generate_doc(api_token, sheet_name, template_file, output_file, mappings):
    try:
        # Load Smartsheet data
        ss_client = smartsheet.Smartsheet(api_token)
        all_sheets = ss_client.Sheets.list_sheets(include_all=True).data
        sheet = next((sheet for sheet in all_sheets if sheet.name == sheet_name), None)

        if sheet is None:
            raise ValueError(f"Sheet with name '{sheet_name}' not found in Smartsheet.")

        sheet_data = ss_client.Sheets.get_sheet(sheet.id)

        # Attempt to load the Word template
        try:
            doc = Document(template_file)
        except Exception as e:
            if "Package not found" in str(e):
                raise IOError(f"The template document '{template_file}' is open in another program. Please close it and try again.")
            else:
                raise e  # Re-raise any unexpected exceptions

        # Debugging: Print the total number of paragraphs in the document
        print(f"Total number of paragraphs in the document: {len(doc.paragraphs)}")

        # Replace static placeholders
        replace_static_placeholders(doc, sheet_data, mappings.get("static_cell_mappings", {}))

        # Save the output document
        doc.save(output_file)
        messagebox.showinfo("Success", f"Document generated successfully: {output_file}")

    except IOError as e:
        messagebox.showerror("Error", str(e))  # Display the I/O error message
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

# GUI Creation
def create_ui():
    global current_theme

    root = ThemedTk(theme="plastik")  # Default to dark mode
    root.title("Smartsheet Document Generator")
    root.geometry("800x600")

    # Load saved configuration
    config = load_config()

    # Filepath entries
    global api_key_var, smartsheet_dropdown
    api_key_var = tk.StringVar(value=config.get("api_key", ""))
    sheet_name_var = tk.StringVar()
    template_file_var = tk.StringVar()
    output_dir_var = tk.StringVar()
    image_dir1_var = tk.StringVar()
    image_dir2_var = tk.StringVar()

    # Load mappings from JSON or prompt user
    mappings = load_mappings()
    if not mappings:
        if messagebox.askyesno("Mappings Missing", "Mappings file not found. Do you want to locate it?"):
            mappings_file_location = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
            if mappings_file_location:
                mappings = load_mappings()
        if not mappings:
            save_new_mappings()
            mappings = load_mappings()

    # GUI Elements
    ttk.Label(root, text="API Key").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    api_key_entry = ttk.Entry(root, textvariable=api_key_var, width=50)
    api_key_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")

    # Save API Key Button
    def save_api_key_to_file():
        api_key = api_key_var.get()
        if api_key:
            file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
            if file_path:
                with open(file_path, 'w') as file:
                    file.write(api_key)
                messagebox.showinfo("Success", "API key saved to file.")

    ttk.Button(root, text="Save Key", command=save_api_key_to_file).grid(row=0, column=2, padx=5, pady=10)

    # Import API Key Button
    def import_api_key_from_file():
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, 'r') as file:
                api_key = file.read().strip()
                api_key_var.set(api_key)
                messagebox.showinfo("Success", "API key imported from file.")
            fetch_smartsheets()

    ttk.Button(root, text="Import Key", command=import_api_key_from_file).grid(row=0, column=3, padx=5, pady=10)

    # Smartsheet Dropdown
    ttk.Label(root, text="Smartsheet Name").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    smartsheet_dropdown = ttk.Combobox(root, textvariable=sheet_name_var, width=47)
    smartsheet_dropdown.grid(row=1, column=1, padx=10, pady=10, sticky="w")

    # Fetch Smartsheets Button
    ttk.Button(root, text="Fetch Smartsheets", command=fetch_smartsheets).grid(row=1, column=2, padx=5, pady=10)

    # Template File
    ttk.Label(root, text="Template File").grid(row=2, column=0, padx=10, pady=10, sticky="e")
    template_entry = ttk.Entry(root, textvariable=template_file_var, width=50)
    template_entry.grid(row=2, column=1, padx=10, pady=10, sticky="w")
    ttk.Button(root, text="Browse", command=lambda: template_file_var.set(filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")]))).grid(row=2, column=2, padx=5, pady=10)

    # Output Directory
    ttk.Label(root, text="Output Directory").grid(row=3, column=0, padx=10, pady=10, sticky="e")
    output_dir_entry = ttk.Entry(root, textvariable=output_dir_var, width=50)
    output_dir_entry.grid(row=3, column=1, padx=10, pady=10, sticky="w")
    ttk.Button(root, text="Browse", command=lambda: output_dir_var.set(filedialog.askdirectory())).grid(row=3, column=2, padx=5, pady=10)

    # Generate Button
    def execute():
        if not (api_key_var.get() and sheet_name_var.get() and template_file_var.get() and output_dir_var.get()):
            messagebox.showerror("Error", "All fields are required!")
            return
        output_file = os.path.join(output_dir_var.get(), f"{sheet_name_var.get()}_output.docx")
        
        # Save the API key for future use
        save_config(api_key_var.get())

        # Call document generation function here
        process_smartsheet_and_generate_doc(
            api_key_var.get(),
            sheet_name_var.get(),
            template_file_var.get(),
            output_file,
            mappings
        )

    ttk.Button(root, text="Generate Document", command=execute).grid(row=6, column=1, pady=20)

    # Automatically fetch smartsheets if the API key is loaded
    if api_key_var.get():
        fetch_smartsheets()

    root.mainloop()

# Fetch smartsheets function
def fetch_smartsheets():
    try:
        api_key = api_key_var.get()
        ss_client = smartsheet.Smartsheet(api_key)
        all_sheets = ss_client.Sheets.list_sheets(include_all=True).data
        smartsheet_dropdown['values'] = [sheet.name for sheet in all_sheets]
        if all_sheets:
            smartsheet_dropdown.current(0)  # Select the first sheet by default
    except Exception as e:
        messagebox.showerror("Error", f"Failed to fetch Smartsheets: {e}")

create_ui()

