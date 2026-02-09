# -*- coding: utf-8 -*-
__title__ = "Count\nElectrical\nFixtures"
__author__ = "Christopher Berndt"
__doc__ = "Counts all Electrical Fixtures in the host and linked models, grouped by Family + Type, and exports to CSV or Excel."

from pyrevit import revit, DB, script, forms
import os
import csv
from datetime import datetime

output = script.get_output()
doc = revit.doc


# -----------------------------
# Helper: Collect Electrical Fixtures
# -----------------------------
def get_electrical_fixtures_from_doc(target_doc):
    try:
        return (
            DB.FilteredElementCollector(target_doc)
            .OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []


# -----------------------------
# Collect from Host and Linked Models
# -----------------------------
host_fixtures = get_electrical_fixtures_from_doc(doc)
linked_fixtures = []

for link_instance in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
    link_doc = link_instance.GetLinkDocument()
    if not link_doc:
        continue
    linked_fixtures.extend(get_electrical_fixtures_from_doc(link_doc))


# -----------------------------
# Count Fixtures by Family : Type
# -----------------------------
def count_fixtures(fixtures):
    counts = {}
    for fixture in fixtures:
        try:
            fam = fixture.Symbol.Family.Name
            typ = fixture.Name
            key = "{} : {}".format(fam, typ)
            counts[key] = counts.get(key, 0) + 1
        except Exception:
            continue
    return counts


host_counts = count_fixtures(host_fixtures)
link_counts = count_fixtures(linked_fixtures)

combined_counts = {}
for k, v in host_counts.items():
    combined_counts[k] = combined_counts.get(k, 0) + v
for k, v in link_counts.items():
    combined_counts[k] = combined_counts.get(k, 0) + v


# -----------------------------
# Prepare Data Table
# -----------------------------
if not combined_counts:
    forms.alert("⚠️ No electrical fixtures found in host or linked models.", exitscript=True)

data = [["Family : Type", "Host Count", "Linked Count", "Total"]]
total_host = total_linked = 0

for key in sorted(combined_counts.keys()):
    host_val = host_counts.get(key, 0)
    link_val = link_counts.get(key, 0)
    total = host_val + link_val
    total_host += host_val
    total_linked += link_val
    data.append([key, str(host_val), str(link_val), str(total)])

grand_total = total_host + total_linked

output.print_table(
    table_data=data,
    title="⚡ Total Fixtures — Host: {} | Linked: {} | Combined: {}".format(
        total_host, total_linked, grand_total
    ),
)


# -----------------------------
# Export Options
# -----------------------------
project_name = os.path.splitext(os.path.basename(doc.PathName))[0] or "Unnamed_Project"
timestamp = datetime.now().strftime("%Y-%m-%d")

# Ask for export format
format_choice = forms.SelectFromList.show(
    ["CSV (.csv)", "Excel (.xlsx)"],
    title="Choose Export Format",
    button_name="Next"
)
if not format_choice:
    forms.alert("❌ Export canceled by user.", exitscript=True)

file_ext = "csv" if format_choice == "CSV (.csv)" else "xlsx"
default_filename = "Fixture_Counts_{}_{}.{}".format(project_name, timestamp, file_ext)

# Ask for save path
save_path = forms.save_file(
    file_ext=file_ext,
    title="Save Fixture Counts As...",
    default_name=default_filename
)
if not save_path:
    forms.alert("❌ Export canceled by user.", exitscript=True)


# -----------------------------
# Export: CSV (IronPython Safe)
# -----------------------------
if file_ext == "csv":
    try:
        f = open(save_path, "wb")  # binary mode for IronPython
        writer = csv.writer(f)
        for row in data:
            safe_row = []
            for item in row:
                if isinstance(item, unicode):
                    safe_row.append(item.encode("utf-8"))
                else:
                    safe_row.append(str(item))
            writer.writerow(safe_row)
        f.close()
        forms.alert("✅ CSV exported successfully:\n{}".format(save_path))
    except Exception as e:
        forms.alert("❌ CSV export failed:\n{}".format(e))


# -----------------------------
# Export: Excel (via XlsxWriter)
# -----------------------------
elif file_ext == "xlsx":
    try:
        import xlsxwriter

        workbook = xlsxwriter.Workbook(save_path)
        worksheet = workbook.add_worksheet("Electrical Fixtures")

        # Define styles
        header_format = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1})
        normal_format = workbook.add_format({'border': 1})
        total_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC'})

        # Write header
        for col_num, header in enumerate(data[0]):
            worksheet.write(0, col_num, header, header_format)

        # Write data rows
        for row_num, row_data in enumerate(data[1:], start=1):
            fmt = total_format if "Total" in row_data[0] else normal_format
            for col_num, value in enumerate(row_data):
                worksheet.write(row_num, col_num, value, fmt)

        # Autofit columns
        worksheet.set_column(0, 0, 45)
        worksheet.set_column(1, 3, 15)

        workbook.close()
        forms.alert("✅ Excel exported successfully:\n{}".format(save_path))

    except ImportError:
        forms.alert("❌ XlsxWriter not installed. Try exporting as CSV instead.")
    except Exception as e:
        forms.alert("❌ Excel export failed:\n{}".format(e))
