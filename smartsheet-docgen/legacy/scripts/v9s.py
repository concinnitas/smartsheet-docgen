import json
import re
import zipfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import smartsheet
from docxtpl import DocxTemplate
from ttkthemes import ThemedTk


APP_TITLE = "Smartsheet Report Generator"
DEFAULT_TEMPLATE_NAME = "FCI AUTOMATION TEMPLATE.docx"
LEGACY_TOKEN_RE = re.compile(r"\[([A-Z][A-Z0-9_\-]*)\]")
DOCXTPL_MARKERS = ("{{", "{%", "{#")

SCRIPT_DIR = Path(__file__).resolve().parent
SEARCH_DIRS = (SCRIPT_DIR, SCRIPT_DIR.parent)


def resolve_data_path(filename):
    for base_dir in SEARCH_DIRS:
        candidate = base_dir / filename
        if candidate.exists():
            return candidate
    return SCRIPT_DIR / filename


CONFIG_FILE = resolve_data_path("config.json")
MAPPINGS_FILE = resolve_data_path("mappings.json")
DEFAULT_TEMPLATE_FILE = resolve_data_path(DEFAULT_TEMPLATE_NAME)
TEMPLATE_GUIDE_FILE = SCRIPT_DIR / "TEMPLATE_MIGRATION.md"


def default_mappings():
    return {
        "static_cell_mappings": {},
        "dynamic_section_mappings": {
            "group_by_column": "Name",
            "fields": {},
        },
    }


def load_json_file(path, default_value):
    if not path.is_file():
        return default_value
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def load_config():
    try:
        return load_json_file(CONFIG_FILE, {})
    except json.JSONDecodeError:
        return {}


def save_config(api_key):
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with CONFIG_FILE.open("w", encoding="utf-8") as handle:
        json.dump({"api_key": api_key.strip()}, handle, indent=4)


def load_mappings():
    try:
        return load_json_file(MAPPINGS_FILE, default_mappings())
    except json.JSONDecodeError as exc:
        raise ValueError(f"Mappings file is not valid JSON: {exc}") from exc


def save_new_mappings(content=None):
    if content is None:
        content = default_mappings()
    MAPPINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with MAPPINGS_FILE.open("w", encoding="utf-8") as handle:
        json.dump(content, handle, indent=4)


def normalize_key(value):
    return str(value).strip().casefold()


def placeholder_name(token):
    return token.strip("[]").replace("-", "_")


def iter_docx_xml_parts(template_path):
    with zipfile.ZipFile(template_path) as archive:
        for part_name in archive.namelist():
            if part_name.startswith("word/") and part_name.endswith(".xml"):
                yield archive.read(part_name).decode("utf-8", errors="ignore")


def template_contains_docxtpl_tags(template_path):
    return any(
        marker in xml_text
        for xml_text in iter_docx_xml_parts(template_path)
        for marker in DOCXTPL_MARKERS
    )


def find_legacy_placeholders(template_path):
    tokens = set()
    for xml_text in iter_docx_xml_parts(template_path):
        tokens.update(LEGACY_TOKEN_RE.findall(xml_text))
    return sorted(tokens)


def validate_template_file(template_path):
    if not template_path.is_file():
        raise ValueError(f"Template not found: {template_path}")

    try:
        has_docxtpl_tags = template_contains_docxtpl_tags(template_path)
        legacy_tokens = find_legacy_placeholders(template_path)
    except zipfile.BadZipFile as exc:
        raise ValueError("The selected template is not a valid .docx file.") from exc

    if legacy_tokens and not has_docxtpl_tags:
        sample = ", ".join(f"[{token}]" for token in legacy_tokens[:8])
        raise ValueError(
            "The selected template still uses legacy bracket placeholders "
            f"such as {sample}. Convert it to docxtpl syntax first. "
            f"See {TEMPLATE_GUIDE_FILE.name} for the required format."
        )

    if not has_docxtpl_tags:
        raise ValueError(
            "No docxtpl tags were detected in the template. Add tags like "
            "{{ PROJECT_NAME }} and loop blocks like {% for row in rows %}."
        )


def create_client(api_key):
    client = smartsheet.Smartsheet(api_key.strip())
    client.errors_as_exceptions(True)
    return client


def build_column_lookup(columns):
    return {normalize_key(column.title): column for column in columns}


def build_row_cell_lookup(row):
    return {cell.column_id: cell for cell in getattr(row, "cells", [])}


def get_cell_value(cell):
    if cell is None:
        return None

    for attr_name in ("display_value", "value"):
        value = getattr(cell, attr_name, None)
        if value is None:
            continue
        if isinstance(value, str):
            value = value.strip()
            if value == "":
                continue
        return value
    return None


def value_or_fallback(value, fallback):
    return fallback if value in (None, "") else value


def resolve_column(mapping_details, column_lookup, sheet_columns):
    column_ref = None
    for key in ("column_name", "column_title", "column"):
        if key in mapping_details and mapping_details[key] not in (None, ""):
            column_ref = mapping_details[key]
            break

    if column_ref is None:
        return None

    if isinstance(column_ref, int):
        if 0 <= column_ref < len(sheet_columns):
            return sheet_columns[column_ref]
        return None

    return column_lookup.get(normalize_key(column_ref))


def available_columns_text(sheet_columns, limit=12):
    column_titles = [str(column.title) for column in sheet_columns]
    if len(column_titles) > limit:
        return ", ".join(column_titles[:limit]) + ", ..."
    return ", ".join(column_titles)


def build_rows_by_group(sheet, row_cell_maps, mappings, column_lookup):
    group_by_column = (
        mappings.get("dynamic_section_mappings", {}).get("group_by_column") or "Name"
    )
    group_column = column_lookup.get(normalize_key(group_by_column))
    if group_column is None:
        raise ValueError(
            f"Group-by column '{group_by_column}' was not found in the Smartsheet. "
            f"Available columns: {available_columns_text(sheet.columns)}"
        )

    rows_by_group = {}
    for row, row_cells in zip(sheet.rows, row_cell_maps):
        group_value = get_cell_value(row_cells.get(group_column.id))
        if group_value in (None, ""):
            continue
        key = str(group_value).strip()
        rows_by_group.setdefault(key, []).append({"row": row, "cells": row_cells})

    return group_by_column, rows_by_group


def build_static_context(sheet, row_cell_maps, mappings, column_lookup):
    context = {}
    warnings = []

    for placeholder, details in mappings.get("static_cell_mappings", {}).items():
        location = details.get("location", {})
        fallback = details.get("fallback", "")
        row_index = None
        column = None
        value = None

        if isinstance(location, dict):
            row_number = location.get("row", 1)
            try:
                row_index = int(row_number) - 1
            except (TypeError, ValueError):
                row_index = None
            column = resolve_column(location, column_lookup, sheet.columns)
        elif isinstance(location, (list, tuple)) and len(location) == 2:
            row_index = location[0]
            column = resolve_column({"column": location[1]}, column_lookup, sheet.columns)

        if row_index is None or not (0 <= row_index < len(sheet.rows)):
            warnings.append(
                f"{placeholder} used its fallback because row '{location}' is invalid."
            )
        elif column is None:
            warnings.append(
                f"{placeholder} used its fallback because its mapped column could not be found."
            )
        else:
            value = get_cell_value(row_cell_maps[row_index].get(column.id))
            if value in (None, ""):
                warnings.append(
                    f"{placeholder} used its fallback because the Smartsheet cell was blank."
                )

        context[placeholder_name(placeholder)] = value_or_fallback(value, fallback)

    return context, warnings


def build_dynamic_rows(sheet, grouped_rows, mappings, column_lookup):
    rows_context = []
    warnings = []
    dynamic_fields = mappings.get("dynamic_section_mappings", {}).get("fields", {})

    for row_number, row_entry in enumerate(grouped_rows, start=1):
        row_context = {"INDEX": row_number}
        for placeholder, details in dynamic_fields.items():
            fallback = details.get("fallback", "")
            column = resolve_column(details, column_lookup, sheet.columns)
            value = None

            if column is None:
                warnings.append(
                    f"{placeholder} used its fallback because its mapped column could not be found."
                )
            else:
                value = get_cell_value(row_entry["cells"].get(column.id))
                if value in (None, ""):
                    warnings.append(
                        f"{placeholder} used its fallback because a grouped row cell was blank."
                    )

            row_context[placeholder_name(placeholder)] = value_or_fallback(value, fallback)

        rows_context.append(row_context)

    return rows_context, warnings


def build_report_context(sheet, row_cell_maps, selected_group, mappings, column_lookup, rows_by_group):
    static_context, static_warnings = build_static_context(
        sheet, row_cell_maps, mappings, column_lookup
    )
    grouped_rows = rows_by_group.get(selected_group, [])
    if not grouped_rows:
        raise ValueError(f"No rows were found for '{selected_group}'.")

    rows_context, dynamic_warnings = build_dynamic_rows(
        sheet, grouped_rows, mappings, column_lookup
    )
    context = {
        **static_context,
        "rows": rows_context,
        "items": rows_context,
        "SELECTED_NAME": selected_group,
        "ITEM_COUNT": len(rows_context),
        "GENERATED_AT": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    return context, static_warnings + dynamic_warnings


def template_help_text(mappings):
    static_tokens = sorted(
        placeholder_name(token)
        for token in mappings.get("static_cell_mappings", {}).keys()
    )
    dynamic_tokens = sorted(
        placeholder_name(token)
        for token in mappings.get("dynamic_section_mappings", {}).get("fields", {}).keys()
    )
    static_line = ", ".join(static_tokens) if static_tokens else "None"
    dynamic_line = ", ".join(dynamic_tokens) if dynamic_tokens else "None"

    return (
        "This app now expects a docxtpl-compatible Word template.\n\n"
        "Static tags should look like:\n"
        "  {{ PROJECT_NAME }}\n"
        "  {{ ADDRESS }}\n\n"
        "Repeated issue rows should be wrapped in a loop like:\n"
        "  {% for row in rows %}\n"
        "  {{ row.ISSUE_ID }}\n"
        "  {{ row.LOCATION }}\n"
        "  {% endfor %}\n\n"
        "Current static fields:\n"
        f"  {static_line}\n\n"
        "Current repeated row fields:\n"
        f"  {dynamic_line}\n\n"
        f"Detailed migration notes: {TEMPLATE_GUIDE_FILE}"
    )


def create_ui():
    root = ThemedTk(theme="plastik")
    root.title(APP_TITLE)
    root.geometry("760x430")
    root.columnconfigure(1, weight=1)

    config = load_config()
    try:
        mappings = load_mappings()
    except ValueError as exc:
        messagebox.showerror("Error", str(exc))
        root.destroy()
        return

    state = {
        "sheet_lookup": {},
        "sheet_data": None,
        "row_cell_maps": [],
        "column_lookup": {},
        "rows_by_group": {},
        "group_by_column": "Name",
    }

    api_key_var = tk.StringVar(value=config.get("api_key", ""))
    sheet_name_var = tk.StringVar()
    template_var = tk.StringVar(
        value=str(DEFAULT_TEMPLATE_FILE) if DEFAULT_TEMPLATE_FILE.exists() else ""
    )
    out_dir_var = tk.StringVar(value=str(SCRIPT_DIR))
    out_file_var = tk.StringVar()
    selected_name = tk.StringVar()
    status_var = tk.StringVar(
        value="Enter an API key, fetch sheets, and choose a group to render."
    )

    def set_status(message):
        status_var.set(message)
        root.update_idletasks()

    def show_template_help():
        messagebox.showinfo("Template Format", template_help_text(mappings))

    def fetch_sheets():
        api_key = api_key_var.get().strip()
        if not api_key:
            messagebox.showerror("Error", "API key is required.")
            return

        try:
            set_status("Fetching available Smartsheet sheets...")
            client = create_client(api_key)
            sheets = client.Sheets.list_sheets(include_all=True).data
            state["sheet_lookup"] = {sheet.name: sheet for sheet in sheets}
            sheet_names = sorted(state["sheet_lookup"].keys())
            sheet_combo.config(values=sheet_names, state="readonly")
            if not sheet_names:
                set_status("No Smartsheet sheets were returned for this API key.")
                messagebox.showwarning(
                    "No Sheets",
                    "No sheets were returned for this API key.",
                )
                return

            if sheet_name_var.get() not in state["sheet_lookup"]:
                sheet_name_var.set(sheet_names[0])
            save_config(api_key)
            set_status(f"Loaded {len(sheet_names)} sheet(s). Loading selected sheet...")
            load_rows()
        except Exception as exc:
            set_status("Failed to fetch sheets.")
            messagebox.showerror("Error", f"Failed to fetch sheets:\n\n{exc}")

    def load_rows(event=None):
        del event

        api_key = api_key_var.get().strip()
        sheet_name = sheet_name_var.get().strip()
        if not api_key or not sheet_name:
            return

        if sheet_name not in state["sheet_lookup"]:
            messagebox.showerror(
                "Error",
                "Please click 'Fetch Sheets' before loading a Smartsheet.",
            )
            return

        try:
            set_status(f"Loading rows from '{sheet_name}'...")
            client = create_client(api_key)
            sheet_summary = state["sheet_lookup"][sheet_name]
            sheet_data = client.Sheets.get_sheet(sheet_summary.id)
            row_cell_maps = [build_row_cell_lookup(row) for row in sheet_data.rows]
            column_lookup = build_column_lookup(sheet_data.columns)
            group_by_column, rows_by_group = build_rows_by_group(
                sheet_data, row_cell_maps, mappings, column_lookup
            )

            state["sheet_data"] = sheet_data
            state["row_cell_maps"] = row_cell_maps
            state["column_lookup"] = column_lookup
            state["rows_by_group"] = rows_by_group
            state["group_by_column"] = group_by_column

            group_names = sorted(rows_by_group.keys())
            name_combo.config(values=group_names, state="readonly")
            if not group_names:
                selected_name.set("")
                set_status(
                    f"Loaded '{sheet_name}', but no grouped rows were found using "
                    f"'{group_by_column}'."
                )
                return

            if selected_name.get() not in rows_by_group:
                selected_name.set(group_names[0])

            set_status(
                f"Loaded '{sheet_name}' with {len(sheet_data.rows)} row(s) across "
                f"{len(group_names)} value(s) in '{group_by_column}'."
            )
        except Exception as exc:
            set_status("Failed to load rows from Smartsheet.")
            messagebox.showerror("Error", f"Failed to load rows:\n\n{exc}")

    def pick_template():
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            template_var.set(path)

    def pick_out_dir():
        chosen_dir = filedialog.askdirectory()
        if chosen_dir:
            out_dir_var.set(chosen_dir)

    def generate_report():
        api_key = api_key_var.get().strip()
        template_path = Path(template_var.get().strip()).expanduser()
        output_dir = Path(out_dir_var.get().strip()).expanduser()
        output_name = out_file_var.get().strip()
        selected_group = selected_name.get().strip()

        if not api_key:
            messagebox.showerror("Error", "API key is required.")
            return
        if state["sheet_data"] is None:
            messagebox.showerror("Error", "Fetch sheets and load a Smartsheet first.")
            return
        if not output_name:
            messagebox.showerror("Error", "Output filename is required.")
            return
        if not selected_group:
            messagebox.showerror("Error", "Pick a group value before generating.")
            return

        try:
            validate_template_file(template_path)
            output_dir.mkdir(parents=True, exist_ok=True)

            if not output_name.lower().endswith(".docx"):
                output_name += ".docx"
            output_path = output_dir / output_name

            set_status(f"Rendering report for '{selected_group}'...")
            context, warnings = build_report_context(
                state["sheet_data"],
                state["row_cell_maps"],
                selected_group,
                mappings,
                state["column_lookup"],
                state["rows_by_group"],
            )

            template = DocxTemplate(str(template_path))
            template.render(context)
            template.save(str(output_path))

            save_config(api_key)
            set_status(
                f"Generated '{output_path.name}' with {context['ITEM_COUNT']} grouped row(s). "
                f"Fallbacks applied: {len(warnings)}."
            )
            messagebox.showinfo(
                "Success",
                f"Generated:\n{output_path}\n\n"
                f"Grouped row count: {context['ITEM_COUNT']}\n"
                f"Fallback values used: {len(warnings)}",
            )
        except Exception as exc:
            set_status("Report generation failed.")
            messagebox.showerror("Error", f"Failed to generate report:\n\n{exc}")

    ttk.Label(root, text="API Key").grid(row=0, column=0, sticky="e", padx=8, pady=6)
    ttk.Entry(root, textvariable=api_key_var, width=50).grid(
        row=0, column=1, sticky="ew", padx=8, pady=6
    )
    ttk.Button(root, text="Fetch Sheets", command=fetch_sheets).grid(
        row=0, column=2, padx=8, pady=6
    )

    ttk.Label(root, text="Smartsheet Name").grid(
        row=1, column=0, sticky="e", padx=8, pady=6
    )
    sheet_combo = ttk.Combobox(
        root,
        textvariable=sheet_name_var,
        width=48,
        state="readonly",
    )
    sheet_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=6)
    sheet_combo.bind("<<ComboboxSelected>>", load_rows)

    ttk.Label(root, text="Group Value").grid(
        row=2, column=0, sticky="e", padx=8, pady=6
    )
    name_combo = ttk.Combobox(
        root,
        textvariable=selected_name,
        width=48,
        state="disabled",
    )
    name_combo.grid(row=2, column=1, sticky="ew", padx=8, pady=6)

    ttk.Label(root, text="Template").grid(row=3, column=0, sticky="e", padx=8, pady=6)
    ttk.Entry(root, textvariable=template_var, width=50).grid(
        row=3, column=1, sticky="ew", padx=8, pady=6
    )
    ttk.Button(root, text="Browse", command=pick_template).grid(
        row=3, column=2, padx=8, pady=6
    )
    ttk.Button(root, text="Template Help", command=show_template_help).grid(
        row=3, column=3, padx=8, pady=6
    )

    ttk.Label(root, text="Output Dir").grid(
        row=4, column=0, sticky="e", padx=8, pady=6
    )
    ttk.Entry(root, textvariable=out_dir_var, width=50).grid(
        row=4, column=1, sticky="ew", padx=8, pady=6
    )
    ttk.Button(root, text="Browse", command=pick_out_dir).grid(
        row=4, column=2, padx=8, pady=6
    )

    ttk.Label(root, text="Output Filename").grid(
        row=5, column=0, sticky="e", padx=8, pady=6
    )
    ttk.Entry(root, textvariable=out_file_var, width=50).grid(
        row=5, column=1, sticky="ew", padx=8, pady=6
    )

    ttk.Button(root, text="Generate", command=generate_report).grid(
        row=6, column=1, sticky="w", padx=8, pady=20
    )

    ttk.Label(root, textvariable=status_var, wraplength=700).grid(
        row=7, column=0, columnspan=4, sticky="w", padx=8, pady=(6, 10)
    )

    root.mainloop()


if __name__ == "__main__":
    create_ui()
