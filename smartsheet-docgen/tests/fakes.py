"""Minimal stand-ins for Smartsheet SDK objects."""
from types import SimpleNamespace as NS


def make_sheet(columns, rows):
    """columns: list of titles; rows: list of lists of values."""
    cols = [NS(id=100 + i, title=t) for i, t in enumerate(columns)]
    sheet_rows = [
        NS(id=1000 + r, cells=[NS(column_id=c.id, value=v, display_value=None) for c, v in zip(cols, vals)])
        for r, vals in enumerate(rows)
    ]
    return NS(columns=cols, rows=sheet_rows)


COLUMNS = ["Name", "PROJECT NAME", "ADDRESS", "DC CAPACITY", "ISSUE ID", "LOCATION",
           "INSPECTION COMMENT", "STATUS", "CONTRACTOR RESPONSE", "AGI COMMENT"]
ROWS = [
    ["Sunny Acres", "Sunny Acres Solar", "123 Farm Rd", "1,250", 1, "Inverter pad", "Strap conduit", "Closed", "Done", "Accepted"],
    ["Sunny Acres", "Sunny Acres Solar", "123 Farm Rd", "1,250", 2, "Combiner 3", "Missing label", "Open", "", "Pending"],
    ["Other Site", "Other", "9 Elm", "500", 1, "Array", "Torque marks", "Open", "", ""],
]
