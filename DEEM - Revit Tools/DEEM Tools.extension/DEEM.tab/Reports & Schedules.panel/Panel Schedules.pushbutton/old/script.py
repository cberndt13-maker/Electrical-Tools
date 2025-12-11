# -*- coding: utf-8 -*-
__title__ = "Export\nPanel Legends"
__author__ = "Christopher + GPT-5"
__doc__ = "Exports DEEM-formatted panel legends to Excel with logo and color mapping."

import os
from datetime import datetime
import xlsxwriter
import yaml
from pyrevit import forms, revit
from Autodesk.Revit.DB import *

# ---------------------------------------------------------------------------
# LOAD YAML CONFIG
# ---------------------------------------------------------------------------
CONFIG_PATH = r"C:\Users\CBerndt\OneDrive - DEEM, LLC\Desktop\PyRevit Stuff\DEEM - Revit Tools\DEEM Tools.extension\DEEM.tab\Reports & Schedules.panel\Panel Schedules.pushbutton\bundle.yaml"  # <-- update path
with open(CONFIG_PATH, 'r') as f:
    config = yaml.safe_load(f)

LOGO_PATH = config['paths']['logo']
EXPORT_FOLDER_DEFAULT = config['paths'].get('export_folder', '')
EXCEL_FORMAT = config['excel']
USER_PROMPT = config['user_prompt']
PANEL_SETTINGS = config['panel_settings']
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
    """Write panel data formatted into Excel sheet using YAML settings."""
    settings = get_panel_config(panel_name)
    slots = settings['slots']
    phase_colors = settings['phase_colors']

    # Title
    ws.merge_range(
        start_row, 0, start_row, 5,
        "Panel: {}".format(panel_name),
        workbook.add_format(EXCEL_FORMAT['title_format'])
    )
    start_row += 2

    # Headers
    headers = ["Circuit", "Load Served", "Breaker", "Phase", "Load (VA)"]
    header_fmt = workbook.add_format(EXCEL_FORMAT['header_format'])
    for col, h in enumerate(headers):
        ws.write(start_row, col, h, header_fmt)
    start_row += 1

    # Circuit rows
    row_fmt = workbook.add_format(EXCEL_FORMAT['row_format'])
    for i in range(slots):
        if i < len(circuits):
            c = circuits[i]
            ws.write(start_row + i, 0, c["Circuit"], row_fmt)
            ws.write(start_row + i, 1, c["Load"], row_fmt)
            ws.write(start_row + i, 2, c["Breaker"], row_fmt)
            phase_fmt = workbook.add_format(EXCEL_FORMAT['phase_format'])
            phase_fmt.set_bg_color(phase_colors.get(c["Phase"], EXCEL_FORMAT['phase_format']['bg_color_default']))
            ws.write(start_row + i, 3, c["Phase"], phase_fmt)
            ws.write(start_row + i, 4, c["LoadValue"], row_fmt)
        else:
            for col in range(5):
                ws.write_blank(start_row + i, col, None, row_fmt)
    start_row += slots + 2

    # Insert logo
    try:
        ws.insert_image(
            start_row, 0, LOGO_PATH,
            {'x_scale': EXCEL_FORMAT['logo_scale']['x'], 'y_scale': EXCEL_FORMAT['logo_scale']['y']}
        )
    except:
        pass

    # Footer text
    ws.write(start_row + 3, 0, EXCEL_FORMAT['footer_text'], workbook.add_format(EXCEL_FORMAT['footer_format']))

    return start_row + 8

# ---------------------------------------------------------------------------
# MAIN EXECUTION
# ---------------------------------------------------------------------------
doc = revit.doc

# Choose single or batch panel export
if USER_PROMPT['select_mode']:
    mode = forms.alert(
        "Export panel legends for one panel or all panels?",
        ok=False,
        yes="Selected Panel",
        no="All Panels"
    )
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
