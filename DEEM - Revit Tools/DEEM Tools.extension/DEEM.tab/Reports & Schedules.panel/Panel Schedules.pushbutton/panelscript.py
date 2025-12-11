# -*- coding: utf-8 -*-
__title__ = "Export\nPanel\nLegends"
__author__ = "Christopher + GPT-5"
__doc__ = "Exports DEEM-formatted panel legends to Excel with logo and color mapping."

import os
from datetime import datetime
import xlsxwriter
from pyrevit import forms, revit
from Autodesk.Revit.DB import *

# ---------------------------------------------------------------------------
# CONFIGURATION (Python dictionary, IronPython-compatible)
# ---------------------------------------------------------------------------
CONFIG = {
    "paths": {
        "logo": r"C:\Users\CBerndt\OneDrive - DEEM, LLC\Desktop\PyRevit Stuff\DEEM - Revit Tools\DEEM Tools.extension\lib\DEEMlogo.png",
        "export_folder": ""
    },
    "panel_settings": {
        "default": {
            "slots": 42,
            "phase_colors": {
                "A": "#8B4513",   # Brown
                "B": "#FFA500",   # Orange
                "C": "#FFFF00",   # Yellow
                "N": "#BFBFBF",   # Grey
                "G": "#008000"    # Green
            }
        },
        "Main Panel": {
            "slots": 48,
            "phase_colors": {
                "A": "#8B4513",   # Brown
                "B": "#FFA500",   # Orange
                "C": "#FFFF00",   # Yellow
                "N": "#BFBFBF",   # Grey
                "G": "#008000"    # Green
            }
        },
        "Sub Panel": {
            "slots": 36,
            "phase_colors": {
                    "A": "#8B4513",   # Brown
                    "B": "#FFA500",   # Orange
                    "C": "#FFFF00",   # Yellow
                    "N": "#BFBFBF",   # Grey
                    "G": "#008000"    # Green
            }
        }
    },
    "excel": {
        "title_format": {"bold": True, "align": "center", "valign": "vcenter", "font_size": 14},
        # make header bold with strong accent color to match reference
        "header_format": {"bold": True, "bg_color": "#CC0000", "font_color": "#FFFFFF", "align": "center", "border": 1},
        "row_format": {"border": 1, "align": "center"},
        "phase_format": {"border": 1, "align": "center", "bg_color_default": "#FFFFFF"},
        "logo_scale": {"x": 0.3, "y": 0.3},
        "footer_text": "DEEM – Electrical Panel Legend",
        "footer_format": {"italic": True, "font_size": 10}
    },
    "user_prompt": {
        "select_mode": True,
        "select_folder": True,
        "alert_on_complete": True
    }
}

LOGO_PATH = CONFIG['paths']['logo']
EXPORT_FOLDER_DEFAULT = CONFIG['paths'].get('export_folder', '')
EXCEL_FORMAT = CONFIG['excel']
USER_PROMPT = CONFIG['user_prompt']
PANEL_SETTINGS = CONFIG['panel_settings']
DEFAULT_PANEL_SETTINGS = PANEL_SETTINGS['default']

# ---------------------------------------------------------------------------
# HELPER FUNCTION: Get panel-specific settings
# ---------------------------------------------------------------------------
def get_panel_config(panel_name):
    """Return slot count and phase colors for the panel, default if not found."""
    for key, settings in PANEL_SETTINGS.items():
        if key != "default" and key.lower() in panel_name.lower():
            return settings
    return DEFAULT_PANEL_SETTINGS

# ---------------------------------------------------------------------------
# HELPER FUNCTIONS: Revit data extraction
# ---------------------------------------------------------------------------
def get_all_panels(doc):
    """Return all electrical equipment panels in project."""
    panels = []
    fec = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType()
    for p in fec:
        try:
            if hasattr(p, "ElectricalSystemType") or "Panel" in p.Name:
                panels.append(p)
        except:
            pass
    return panels

def get_panel_circuits(panel):
    """Return circuit data for a given panel."""
    circuits_data = []
    try:
        circuits = panel.ElectricalSystems
        for c in circuits:
            try:
                circ_num = c.CircuitNumber
                load_name = c.LoadName
                breaker_size = c.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_BREAKER_PARAM).AsValueString()
                phase = ""
                if c.get_Parameter(BuiltInParameter.RBS_ELEC_PHASE_PARAM):
                    phase = c.get_Parameter(BuiltInParameter.RBS_ELEC_PHASE_PARAM).AsString()
                load_val = ""
                if c.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD_PARAM):
                    load_val = c.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD_PARAM).AsValueString()

                circuits_data.append({
                    "Circuit": circ_num,
                    "Load": load_name,
                    "Breaker": breaker_size,
                    "Phase": phase,
                    "LoadValue": load_val
                })
            except:
                continue
    except:
        pass
    return circuits_data

# ---------------------------------------------------------------------------
# HELPER FUNCTION: Write panel to Excel
# ---------------------------------------------------------------------------
def write_panel_to_excel(ws, panel_name, circuits, start_row, workbook):
    """Write DEEM-style panel schedule (2 circuits per row, numbered CIR columns)."""
    settings = get_panel_config(panel_name)
    slots = settings['slots']  # total number of slots
    row_fmt = workbook.add_format(EXCEL_FORMAT['row_format'])
    header_fmt = workbook.add_format(EXCEL_FORMAT['header_format'])
    title_fmt = workbook.add_format(EXCEL_FORMAT['title_format'])
    fed_fmt = workbook.add_format({'bold': True, 'align': 'left'})
    phase_fmt = workbook.add_format({'bold': True, 'align': 'left'})

    # Build phase formats from settings so we can color-code circuits
    phase_colors = settings.get('phase_colors', {})
    phase_formats = {}
    for ph, col_hex in phase_colors.items():
        try:
            phase_formats[ph.upper()] = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': col_hex})
        except:
            phase_formats[ph.upper()] = row_fmt

    # -----------------------
    # Panel Title
    # -----------------------
    ws.merge_range(start_row, 0, start_row, 5, "Panel: {}".format(panel_name), title_fmt)
    start_row += 2

    # -----------------------
    # FED FROM
    # -----------------------
    ws.write(start_row, 0, "FED FROM:", fed_fmt)
    ws.write(start_row, 2, "120/208 V - 3 Ph - 4 W", fed_fmt)
    start_row += 2

    # -----------------------
    # Headers
    # -----------------------
    headers = ["LOAD SERVED", "BRK", "CIR", "CIR", "BRK", "LOAD SERVED"]
    for col, h in enumerate(headers):
        ws.write(start_row, col, h, header_fmt)
        # set column widths with better defaults per column
        if col in (0, 5):
            ws.set_column(col, col, 30)
        elif col in (1, 4):
            ws.set_column(col, col, 8)
        else:
            ws.set_column(col, col, 6)
    start_row += 1

    # -----------------------
    # Circuit rows (2 per row)
    # -----------------------
    total_rows = slots // 2
    for i in range(total_rows):
        left_index = i * 2
        right_index = left_index + 1

        # Left circuit
        if left_index < len(circuits):
            c1 = circuits[left_index]
            # apply phase color format if available
            ph = (c1.get("Phase") or "").upper()
            fmt_left = phase_formats.get(ph, row_fmt)
            ws.write(start_row, 0, c1["Load"], fmt_left)
            ws.write(start_row, 1, c1["Breaker"], fmt_left)
        else:
            ws.write_blank(start_row, 0, None, row_fmt)
            ws.write_blank(start_row, 1, None, row_fmt)
        ws.write(start_row, 2, left_index + 1, row_fmt)  # CIR number

        # Right circuit
        if right_index < len(circuits):
            c2 = circuits[right_index]
            ph2 = (c2.get("Phase") or "").upper()
            fmt_right = phase_formats.get(ph2, row_fmt)
            ws.write(start_row, 4, c2["Breaker"], fmt_right)
            ws.write(start_row, 5, c2["Load"], fmt_right)
        else:
            ws.write_blank(start_row, 4, None, row_fmt)
            ws.write_blank(start_row, 5, None, row_fmt)
        ws.write(start_row, 3, right_index + 1, row_fmt)  # CIR number

        start_row += 1

    start_row += 1

    # -----------------------
    # Phase legend (colored boxes + labels)
    # -----------------------
    legend_row = start_row
    col_box = 0
    col_label = 1
    ws.set_column(col_box, col_box, 4)
    ws.set_column(col_label, col_label, 28)

    legend_items = [
        ("A", "Phase A - Brown"),
        ("B", "Phase B - Orange"),
        ("C", "Phase C - Yellow"),
        ("N", "Neutral - Grey"),
        ("G", "Ground - Green"),
    ]

    for key, label in legend_items:
        # get a format for the color box; fall back to a default boxed format
        box_fmt = phase_formats.get(key, workbook.add_format({'bg_color': EXCEL_FORMAT['phase_format'].get('bg_color_default', '#FFFFFF'), 'border': 1}))
        # draw a small colored box (merge a single cell so it looks filled)
        ws.merge_range(legend_row, col_box, legend_row, col_box, "", box_fmt)
        # write the label next to it
        ws.write(legend_row, col_label, label, phase_fmt)
        legend_row += 1

    start_row = legend_row + 1

    # -----------------------
    # Logo + footer
    # -----------------------
    try:
        ws.insert_image(start_row, 0, LOGO_PATH, {'x_scale': EXCEL_FORMAT['logo_scale']['x'], 'y_scale': EXCEL_FORMAT['logo_scale']['y']})
    except:
        pass
    ws.write(start_row + 3, 0, EXCEL_FORMAT['footer_text'], workbook.add_format(EXCEL_FORMAT['footer_format']))

    return start_row + 8




# ---------------------------------------------------------------------------
# MAIN EXECUTION
# ---------------------------------------------------------------------------
doc = revit.doc

# Choose single or batch panel export
if USER_PROMPT['select_mode']:
    mode_choice = forms.SelectFromList.show(
        ["Single Panel", "All Panels"],
        title="Choose export mode",
        button_name="Export"
    )
    if not mode_choice:
        raise Exception("Export canceled by user.")
    mode = (mode_choice == "Single Panel")
else:
    mode = False

if mode:
    selection = revit.get_selection()
    if not selection or len(selection) == 0:
        forms.alert("Please select a panel first.")
        raise Exception("No panel selected.")
    panels = [selection[0]]
else:
    panels = get_all_panels(doc)

if not panels:
    forms.alert("No panels found in project.")
    raise Exception("No panels found.")

# Choose export folder
if USER_PROMPT['select_folder'] or not EXPORT_FOLDER_DEFAULT:
    export_folder = forms.pick_folder(title="Select Export Folder")
else:
    export_folder = EXPORT_FOLDER_DEFAULT

if not export_folder:
    raise Exception("Export canceled.")

timestamp = datetime.now().strftime("%Y%m%d_%H%M")
file_path = os.path.join(export_folder, "DEEM_Panel_Legends_{}.xlsx".format(timestamp))

workbook = xlsxwriter.Workbook(file_path)
ws = workbook.add_worksheet("Panel Legends")

current_row = 0
for p in panels:
    try:
        circuits = get_panel_circuits(p)
        panel_name = p.Name
        current_row = write_panel_to_excel(ws, panel_name, circuits, current_row, workbook)
    except Exception as e:
        print("Error processing panel {}: {}".format(p.Name, e))
        continue

workbook.close()

if USER_PROMPT['alert_on_complete']:
    forms.alert("Export complete!\n\nFile saved to:\n{}".format(file_path))
