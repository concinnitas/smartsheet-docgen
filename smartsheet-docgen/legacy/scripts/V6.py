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

# Track the current theme
#current_theme = "dark"  # Set default theme to dark

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

# Handle dynamic section duplication
def duplicate_sections(doc, sheet_data, dynamic_mappings, image_dir1, image_dir2):
    try:
        # Validate dynamic mappings
        if "header_text" not in dynamic_mappings or "fields" not in dynamic_mappings:
            raise KeyError("Missing required keys in dynamic mappings ('header_text' or 'fields')")

        section_header_text = dynamic_mappings.get("header_text", "[SECTION_START]")
        fields = dynamic_mappings.get("fields", {})
        starting_row = dynamic_mappings.get("starting_row", 27)
        end_row = dynamic_mappings.get("end_row", len(sheet_data.rows) - 1)

        if not section_header_text:
            raise ValueError("The 'header_text' in the mappings cannot be empty.")

        # Locate template section
        template_section = locate_template_section(doc, section_header_text)

        # Remove the original template section
        for para in template_section:
            p = para._element
            p.getparent().remove(p)

        # Duplicate sections based on the row range
        for i, row_index in enumerate(range(starting_row, end_row + 1)):
            try:
                row = sheet_data.rows[row_index]
                for template_para in template_section:
                    new_para = copy.deepcopy(template_para)
                    doc._body._element.append(new_para._element)

                    # Replace placeholders in the duplicated section
                    for placeholder, details in fields.items():
                        col_index = details.get("column", 0)
                        fallback = details.get("fallback", "")
                        try:
                            cell_value = row.cells[col_index].value
                            cell_value = cell_value if cell_value else fallback  # Use fallback if value is None or empty
                            for run in new_para.runs:
                                if placeholder in run.text:
                                    print(f"Replacing placeholder '{placeholder}' with value '{cell_value}' in run: {run.text}")
                                    run.text = run.text.replace(placeholder, str(cell_value))
                        except IndexError:
                            print(f"Invalid column index {col_index} for row {row_index}")

                    # Debugging for [SECTION_START] and [SECTION_END]
                    print(f"Before cleaning tags, paragraph text: {''.join(run.text for run in new_para.runs)}")

                    # Combine all runs' text, process, and reassign
                    combined_text = "".join(run.text for run in new_para.runs)
                    cleaned_text = combined_text.replace("[SECTION_START]", "").replace("[SECTION_END]", "").strip()

                    # Update all runs with the cleaned text
                    if new_para.runs:
                        new_para.runs[0].text = cleaned_text  # Assign cleaned text to the first run
                        for run in new_para.runs[1:]:
                            run.clear()  # Clear other runs
                    else:
                        print(f"No runs available to clean tags in paragraph: {new_para.text}")

                    # After cleaning
                    print(f"After cleaning tags, paragraph text: {''.join(run.text for run in new_para.runs)}")

                   # Replace images based on section index 
                    def resolve_image_path(alt_text, index, image_dir1, image_dir2):
                        """Resolve image path dynamically based on alt text and directory."""
                        image_mapping = {
                            "fig1": f"{index + 1}_fig1",
                            "fig2": f"{index + 1}_fig2"
                        }
                        if alt_text in image_mapping:
                            base_name = image_mapping[alt_text]
                            for ext in ["jpg", "jpeg"]:
                                potential_path = os.path.join(image_dir1 if "fig1" in alt_text else image_dir2, f"{base_name}.{ext}")
                                if os.path.isfile(potential_path):
                                    return potential_path
                        return None

                    # Main processing loop
                    for i, run in enumerate(new_para.runs):
                        try:
                            # Debug: Print run element XML for inspection
                            run_xml = run._element.xml
                            print(f"Run {i} XML: {run_xml}")

                            # Locate both inline and anchored image elements
                            blip_elements = run._element.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
                            anchor_elements = run._element.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor")

                            if not blip_elements and not anchor_elements:
                                print(f"No image placeholder or anchor found in run {i}. Skipping.")
                                continue

                            # Handle inline <a:blip>
                            for blip in blip_elements:
                                alt_text = blip.get(f"{{{REL_NS}}}descr", None)
                                if not alt_text:
                                    print(f"No alt text found in inline image in run {i}. Skipping.")
                                    continue

                                print(f"Alt text detected in inline image in run {i}: {alt_text}")
                                image_path = resolve_image_path(alt_text, i, image_dir1, image_dir2)

                                if not image_path:
                                    print(f"No valid image found for alt text '{alt_text}' in inline image in run {i}. Clearing run.")
                                    run.clear()
                                    continue

                                print(f"Inserting inline image from: {image_path}")
                                with open(image_path, "rb") as img_file:
                                    image_data = io.BytesIO(img_file.read())
                                    run.clear()
                                    try:
                                        run.add_picture(image_data, width=Inches(2))
                                    except Exception as e:
                                        print(f"Error adding inline image to run {i}: {e}")

                            # Handle floating images in <wp:anchor>
                            for anchor in anchor_elements:
                                blip = anchor.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
                                if blip is None:
                                    print(f"No <a:blip> found in anchor for run {i}. Skipping.")
                                    continue

                                alt_text = blip.get(f"{{{REL_NS}}}descr", None)
                                if not alt_text:
                                    print(f"No alt text found in anchor image in run {i}. Skipping.")
                                    continue

                                print(f"Alt text detected in anchor image in run {i}: {alt_text}")
                                image_path = resolve_image_path(alt_text, i, image_dir1, image_dir2)

                                if not image_path:
                                    print(f"No valid image found for alt text '{alt_text}' in anchor image in run {i}. Clearing run.")
                                    run.clear()
                                    continue

                                print(f"Inserting floating image from: {image_path}")
                                with open(image_path, "rb") as img_file:
                                    image_data = io.BytesIO(img_file.read())
                                    run.clear()
                                    try:
                                        run.add_picture(image_data, width=Inches(2))
                                    except Exception as e:
                                        print(f"Error adding floating image to run {i}: {e}")
                        except Exception as e:
                            print(f"Unexpected error processing run {i}: {e}")

                    # Debugging: Print run element XML for further inspection if needed
                    for run in new_para.runs:
                        print(run._element.xml)

            except IndexError:
                print(f"Row {row_index} is out of range in Smartsheet.")
    except KeyError as e:
        messagebox.showerror("Mapping Error", f"Mapping Error: {e}")
    except ValueError as e:
        messagebox.showerror("Value Error", f"Value Error: {e}")
        
# Locate template section using the updated "[SECTION_START]" without specific style considerations
def locate_template_section(doc, header_text):
    template_section = []
    in_section = False

    # Convert the expected header text to lowercase and strip whitespace for comparison
    expected_header = header_text.strip().lower()

    # Iterate over paragraphs in the document
    for idx, paragraph in enumerate(doc.paragraphs):
        # Collect the full text of the paragraph, including all runs, to ensure nothing is missed
        full_text = "".join(run.text for run in paragraph.runs).strip().lower()

        # Debugging: Print each paragraph's full text for inspection
        print(f"Inspecting Paragraph {idx}: '{full_text}'")

        # Check if paragraph text matches the header text (ignoring case and whitespace)
        if expected_header in full_text:
            in_section = True
            print("Found section start.")

        if in_section:
            template_section.append(paragraph)
            # End when [SECTION_END] is found (ignoring case and whitespace)
            if "[SECTION_END]" in full_text:
                print("End of section found.")
                break

    if not template_section:
        # Print all paragraphs for further inspection
        print("No template section found. Here are all paragraphs in the document:")
        for idx, paragraph in enumerate(doc.paragraphs):
            full_text = "".join(run.text for run in paragraph.runs).strip()
            print(f"Paragraph {idx}: '{full_text}'")

        raise ValueError(f"No template section found in the document with header '{header_text}'. Please verify that the text '[SECTION_START]' exists as expected in the document.")

    return template_section

# Main processing function
def process_smartsheet_and_generate_doc(api_token, sheet_name, template_file, output_file, mappings, image_dir1, image_dir2):
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

        # Handle dynamic sections
        duplicate_sections(doc, sheet_data, mappings.get("dynamic_section_mappings", {}), image_dir1, image_dir2)

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

    # Light/Dark Mode Toggle
   # def toggle_mode():
   #     global current_theme
   #     if current_theme == "dark":  # Dark mode is active, switch to light
   #         root.set_theme("clam")  # Switch to light mode
   #         update_colors(light_mode=True)
   #         current_theme = "light"
   #     else:
   #         root.set_theme("plastik")  # Switch to dark mode
   #         update_colors(light_mode=False)
   #         current_theme = "dark"

    # Update color scheme based on the mode
   # def update_colors(light_mode=False):
   #     text_color = "black" if light_mode else "white"
   #     api_key_entry.config(foreground=text_color)
   #     smartsheet_dropdown.config(foreground=text_color)
   #     template_entry.config(foreground=text_color)
   #     output_dir_entry.config(foreground=text_color)
   #     image_dir1_entry.config(foreground=text_color)
   #     image_dir2_entry.config(foreground=text_color)

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

    # Image Directories
    ttk.Label(root, text="Image Directory 1").grid(row=4, column=0, padx=10, pady=10, sticky="e")
    image_dir1_entry = ttk.Entry(root, textvariable=image_dir1_var, width=50)
    image_dir1_entry.grid(row=4, column=1, padx=10, pady=10, sticky="w")
    ttk.Button(root, text="Browse", command=lambda: image_dir1_var.set(filedialog.askdirectory())).grid(row=4, column=2, padx=5, pady=10)

    ttk.Label(root, text="Image Directory 2").grid(row=5, column=0, padx=10, pady=10, sticky="e")
    image_dir2_entry = ttk.Entry(root, textvariable=image_dir2_var, width=50)
    image_dir2_entry.grid(row=5, column=1, padx=10, pady=10, sticky="w")
    ttk.Button(root, text="Browse", command=lambda: image_dir2_var.set(filedialog.askdirectory())).grid(row=5, column=2, padx=5, pady=10)

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
            mappings,
            image_dir1_var.get(),
            image_dir2_var.get()
        )

    ttk.Button(root, text="Generate Document", command=execute).grid(row=6, column=1, pady=20)

    # Light/Dark Mode Toggle Button
    #toggle_button = ttk.Button(root, text="Toggle Light/Dark Mode", command=toggle_mode)
    #toggle_button.grid(row=7, column=1, pady=10)

    # Automatically fetch smartsheets if the API key is loaded
    if api_key_var.get():
        fetch_smartsheets()

    # Start with updating colors based on the default theme
    #update_colors(light_mode=False)

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

