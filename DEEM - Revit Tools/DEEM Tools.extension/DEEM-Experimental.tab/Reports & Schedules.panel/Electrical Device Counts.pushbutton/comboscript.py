# -*- coding: utf-8 -*-
__title__ = "Count\nElectrical\nFixtures"
__author__ = "Christopher Berndt"
__doc__ = "Counts all Electrical Fixtures in the host and linked models, grouped by Family + Type, and exports to CSV or Excel."

from pyrevit import revit, DB, script, forms
import os
import csv
from datetime import datetime

# WPF imports for custom window
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows import Window, Application
from System.Windows.Controls import Button, StackPanel, TextBlock, DataGrid, DataGridTextColumn, Grid, RowDefinition, ComboBox, ComboBoxItem, ScrollViewer, CheckBox
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment
from System.Windows.Data import Binding
from System.Collections.ObjectModel import ObservableCollection

output = script.get_output()
doc = revit.doc


# -----------------------------
# Helper: Natural Sort
# -----------------------------
def natural_sort_key(text):
    """Generate a sort key for natural/alphanumeric sorting"""
    import re
    def atoi(text):
        return int(text) if text.isdigit() else text.lower()
    return [atoi(c) for c in re.split(r'(\d+)', text)]


# -----------------------------
# Data Row Classes
# -----------------------------
class FamilyHeaderRow:
    """Class to represent a family header row"""
    def __init__(self, family_name, total_count):
        self.IsHeader = True
        self.FamilyName = family_name
        self.TypeName = family_name  # Show family name in Type column for headers
        self.HostCount = ""
        self.LinkedCount = ""
        self.Total = total_count


class FixtureCountRow:
    """Class to represent a fixture type row in the table"""
    def __init__(self, family_name, type_name, category, host_count, linked_count, total):
        self.IsHeader = False
        self.FamilyName = family_name
        self.TypeName = type_name
        self.Category = category  # "Lighting" or "Electrical"
        self.HostCount = host_count
        self.LinkedCount = linked_count
        self.Total = total
        self._base_host_count = host_count  # Store original counts
        self._base_linked_count = linked_count


# -----------------------------
# Fixture Count Results Window
# -----------------------------
class FixtureCountWindow(Window):
    """WPF Window to display fixture counts with export options"""

    def __init__(self, all_data_rows, total_host, total_linked, total_combined, project_name):
        self.all_data_rows = all_data_rows  # All rows (includes both headers and data)
        self.project_name = project_name
        self.total_host = total_host
        self.total_linked = total_linked
        self.total_combined = total_combined

        self.Title = "Electrical & Lighting Fixture Counts"
        self.Width = 800
        self.Height = 650
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        # Main layout using Grid
        main_grid = Grid()
        main_grid.Margin = Thickness(15)

        # Define rows: Title, Filters, DataGrid (flexible), Buttons
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[0].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[1].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[2].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)  # Flexible

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[3].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        # Title with totals
        self.title_text = TextBlock()
        self.title_text.Text = "Total Fixtures - Host: {} | Linked: {} | Combined: {}".format(
            total_host, total_linked, total_combined
        )
        self.title_text.FontSize = 14
        self.title_text.FontWeight = System.Windows.FontWeights.Bold
        self.title_text.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.title_text, 0)
        main_grid.Children.Add(self.title_text)

        # Filter panel
        filter_panel = StackPanel()
        filter_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        filter_panel.Margin = Thickness(0, 0, 0, 10)

        # Category filter
        category_label = TextBlock()
        category_label.Text = "Filter by Category: "
        category_label.VerticalAlignment = VerticalAlignment.Center
        category_label.Margin = Thickness(0, 0, 10, 0)
        filter_panel.Children.Add(category_label)

        self.category_combo = ComboBox()
        self.category_combo.Width = 200
        self.category_combo.Margin = Thickness(0, 0, 30, 0)

        cat_all = ComboBoxItem()
        cat_all.Content = "All Elements"
        cat_all.Tag = "All"
        self.category_combo.Items.Add(cat_all)

        cat_lighting = ComboBoxItem()
        cat_lighting.Content = "Lighting Fixtures"
        cat_lighting.Tag = "Lighting"
        self.category_combo.Items.Add(cat_lighting)

        cat_electrical = ComboBoxItem()
        cat_electrical.Content = "Electrical Fixtures"
        cat_electrical.Tag = "Electrical"
        self.category_combo.Items.Add(cat_electrical)

        self.category_combo.SelectedIndex = 0
        self.category_combo.SelectionChanged += self.filter_changed
        filter_panel.Children.Add(self.category_combo)

        # Count visibility checkboxes
        show_label = TextBlock()
        show_label.Text = "Show Counts: "
        show_label.VerticalAlignment = VerticalAlignment.Center
        show_label.Margin = Thickness(0, 0, 10, 0)
        filter_panel.Children.Add(show_label)

        self.show_host_check = CheckBox()
        self.show_host_check.Content = "Host"
        self.show_host_check.IsChecked = True
        self.show_host_check.VerticalAlignment = VerticalAlignment.Center
        self.show_host_check.Margin = Thickness(0, 0, 20, 0)
        self.show_host_check.Checked += self.count_visibility_changed
        self.show_host_check.Unchecked += self.count_visibility_changed
        filter_panel.Children.Add(self.show_host_check)

        self.show_linked_check = CheckBox()
        self.show_linked_check.Content = "Linked"
        self.show_linked_check.IsChecked = True
        self.show_linked_check.VerticalAlignment = VerticalAlignment.Center
        self.show_linked_check.Checked += self.count_visibility_changed
        self.show_linked_check.Unchecked += self.count_visibility_changed
        filter_panel.Children.Add(self.show_linked_check)

        Grid.SetRow(filter_panel, 1)
        main_grid.Children.Add(filter_panel)

        # DataGrid for fixture counts
        self.data_grid = DataGrid()
        self.data_grid.IsReadOnly = True
        self.data_grid.AutoGenerateColumns = False
        self.data_grid.CanUserSortColumns = False  # Disable sorting to maintain hierarchy
        self.data_grid.CanUserResizeColumns = True
        self.data_grid.GridLinesVisibility = System.Windows.Controls.DataGridGridLinesVisibility.All
        self.data_grid.HeadersVisibility = System.Windows.Controls.DataGridHeadersVisibility.Column
        self.data_grid.RowHeight = 25

        # Define columns
        col1 = DataGridTextColumn()
        col1.Header = "Type"
        col1.Binding = Binding("TypeName")
        col1.Width = System.Windows.Controls.DataGridLength(350)
        self.data_grid.Columns.Add(col1)

        col2 = DataGridTextColumn()
        col2.Header = "Host Count"
        col2.Binding = Binding("HostCount")
        col2.Width = System.Windows.Controls.DataGridLength(120)
        self.data_grid.Columns.Add(col2)

        col3 = DataGridTextColumn()
        col3.Header = "Linked Count"
        col3.Binding = Binding("LinkedCount")
        col3.Width = System.Windows.Controls.DataGridLength(120)
        self.data_grid.Columns.Add(col3)

        col4 = DataGridTextColumn()
        col4.Header = "Total"
        col4.Binding = Binding("Total")
        col4.Width = System.Windows.Controls.DataGridLength(100)
        self.data_grid.Columns.Add(col4)

        # Create observable collection for data binding
        self.items = ObservableCollection[object]()
        self.data_grid.ItemsSource = self.items

        # Add row style to handle header formatting
        self.data_grid.LoadingRow += self.on_loading_row

        # Initial population
        self.populate_grid()

        Grid.SetRow(self.data_grid, 2)
        main_grid.Children.Add(self.data_grid)

        # Buttons
        button_panel = StackPanel()
        button_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Right
        button_panel.Margin = Thickness(0, 10, 0, 0)

        csv_btn = Button()
        csv_btn.Content = "Export to CSV"
        csv_btn.Width = 120
        csv_btn.Height = 30
        csv_btn.Margin = Thickness(0, 0, 5, 0)
        csv_btn.Click += self.export_csv_click
        button_panel.Children.Add(csv_btn)

        excel_btn = Button()
        excel_btn.Content = "Export to Excel"
        excel_btn.Width = 120
        excel_btn.Height = 30
        excel_btn.Margin = Thickness(0, 0, 5, 0)
        excel_btn.Click += self.export_excel_click
        button_panel.Children.Add(excel_btn)

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Height = 30
        cancel_btn.Click += self.cancel_click
        button_panel.Children.Add(cancel_btn)

        Grid.SetRow(button_panel, 3)
        main_grid.Children.Add(button_panel)

        self.Content = main_grid

    def on_loading_row(self, sender, args):
        """Style rows differently for headers vs data"""
        row_data = args.Row.Item
        if hasattr(row_data, 'IsHeader') and row_data.IsHeader:
            # Header row - bold and gray background
            args.Row.FontWeight = System.Windows.FontWeights.Bold
            args.Row.Background = System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(230, 230, 230)
            )
        else:
            # Normal row
            args.Row.FontWeight = System.Windows.FontWeights.Normal
            args.Row.Background = System.Windows.Media.Brushes.White

    def populate_grid(self):
        """Populate grid based on current filters"""
        self.items.Clear()

        # Get selected category filter
        category_item = self.category_combo.SelectedItem
        if category_item is None:
            return

        selected_category = category_item.Tag

        # Get show/hide settings for counts
        show_host = self.show_host_check.IsChecked
        show_linked = self.show_linked_check.IsChecked

        # Show all data
        rows_to_show = self.all_data_rows

        # Filter by category and update totals
        visible_families = {}  # Track which families have visible children

        for row in rows_to_show:
            if row.IsHeader:
                # Mark family headers for later - we'll add them if they have visible children
                visible_families[row.FamilyName] = {'header': row, 'has_children': False, 'total': 0}
            else:
                # Check category filter
                if selected_category != "All" and row.Category != selected_category:
                    continue

                # Update total based on visible counts
                new_total = 0
                if show_host:
                    new_total += row._base_host_count
                if show_linked:
                    new_total += row._base_linked_count

                row.Total = new_total

                # Update displayed counts
                row.HostCount = row._base_host_count if show_host else 0
                row.LinkedCount = row._base_linked_count if show_linked else 0

                # Mark that this family has visible children
                if row.FamilyName in visible_families:
                    visible_families[row.FamilyName]['has_children'] = True
                    visible_families[row.FamilyName]['total'] += new_total

        # Add rows to display
        for row in rows_to_show:
            if row.IsHeader:
                # Only add family header if it has visible children
                if row.FamilyName in visible_families and visible_families[row.FamilyName]['has_children']:
                    # Update header total
                    row.Total = visible_families[row.FamilyName]['total']
                    self.items.Add(row)
            else:
                # Check category filter
                if selected_category != "All" and row.Category != selected_category:
                    continue
                self.items.Add(row)

        # Update title with current totals
        self.update_title()

    def update_title(self):
        """Update title with current visible totals"""
        show_host = self.show_host_check.IsChecked
        show_linked = self.show_linked_check.IsChecked

        current_host = 0
        current_linked = 0

        for row in self.items:
            if not row.IsHeader:
                if show_host:
                    current_host += row._base_host_count
                if show_linked:
                    current_linked += row._base_linked_count

        current_total = current_host + current_linked

        # Update the title TextBlock
        self.title_text.Text = "Total Fixtures - Host: {} | Linked: {} | Combined: {}".format(
            current_host if show_host else 0,
            current_linked if show_linked else 0,
            current_total
        )

    def count_visibility_changed(self, sender, args):
        """Handle count visibility checkbox changes"""
        self.populate_grid()

    def filter_changed(self, sender, args):
        """Handle filter change"""
        self.populate_grid()

    def export_csv_click(self, sender, args):
        """Export data to CSV"""
        timestamp = datetime.now().strftime("%Y-%m-%d")
        default_filename = "Fixture_Counts_{}_{}.csv".format(self.project_name, timestamp)

        save_path = forms.save_file(
            file_ext="csv",
            title="Save Fixture Counts As CSV",
            default_name=default_filename
        )

        if not save_path:
            return

        try:
            f = open(save_path, "wb")  # binary mode for IronPython
            writer = csv.writer(f)

            # Write header
            writer.writerow(["Type", "Host Count", "Linked Count", "Total"])

            # Write all visible rows (includes headers and data)
            for row in self.items:
                if row.IsHeader:
                    # Family header row
                    safe_row = [
                        self._safe_encode(row.FamilyName),
                        "",  # No host count for header
                        "",  # No linked count for header
                        self._safe_encode(str(row.Total))
                    ]
                else:
                    # Data row
                    safe_row = [
                        self._safe_encode(row.TypeName),
                        self._safe_encode(str(row.HostCount)),
                        self._safe_encode(str(row.LinkedCount)),
                        self._safe_encode(str(row.Total))
                    ]
                writer.writerow(safe_row)

            # Write totals row
            writer.writerow(["TOTAL", str(self.total_host), str(self.total_linked), str(self.total_combined)])

            f.close()
            forms.alert("CSV exported successfully:\n{}".format(save_path), title="Export Successful")
        except Exception as e:
            forms.alert("CSV export failed:\n{}".format(e), title="Export Failed")

    def _safe_encode(self, item):
        """Safely encode string for CSV export"""
        try:
            if hasattr(item, 'encode'):
                return item.encode("utf-8")
            else:
                return str(item)
        except:
            return str(item)

    def export_excel_click(self, sender, args):
        """Export data to Excel"""
        timestamp = datetime.now().strftime("%Y-%m-%d")
        default_filename = "Fixture_Counts_{}_{}.xlsx".format(self.project_name, timestamp)

        save_path = forms.save_file(
            file_ext="xlsx",
            title="Save Fixture Counts As Excel",
            default_name=default_filename
        )

        if not save_path:
            return

        try:
            import xlsxwriter

            workbook = xlsxwriter.Workbook(save_path)
            worksheet = workbook.add_worksheet("Electrical Fixtures")

            # Define styles
            header_format = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1})
            normal_format = workbook.add_format({'border': 1})
            family_header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#E6E6E6'})
            total_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC'})

            # Write header
            headers = ["Type", "Host Count", "Linked Count", "Total"]
            for col_num, header in enumerate(headers):
                worksheet.write(0, col_num, header, header_format)

            # Write all visible rows
            row_num = 1
            for row in self.items:
                if row.IsHeader:
                    # Family header row
                    worksheet.write(row_num, 0, row.FamilyName, family_header_format)
                    worksheet.write(row_num, 1, "", family_header_format)
                    worksheet.write(row_num, 2, "", family_header_format)
                    worksheet.write(row_num, 3, row.Total, family_header_format)
                else:
                    # Data row
                    worksheet.write(row_num, 0, row.TypeName, normal_format)
                    worksheet.write(row_num, 1, row.HostCount, normal_format)
                    worksheet.write(row_num, 2, row.LinkedCount, normal_format)
                    worksheet.write(row_num, 3, row.Total, normal_format)
                row_num += 1

            # Write totals row
            worksheet.write(row_num, 0, "TOTAL", total_format)
            worksheet.write(row_num, 1, self.total_host, total_format)
            worksheet.write(row_num, 2, self.total_linked, total_format)
            worksheet.write(row_num, 3, self.total_combined, total_format)

            # Autofit columns
            worksheet.set_column(0, 0, 40)  # Type
            worksheet.set_column(1, 3, 15)  # Counts

            workbook.close()
            forms.alert("Excel exported successfully:\n{}".format(save_path), title="Export Successful")

        except ImportError:
            forms.alert("XlsxWriter not installed. Try exporting as CSV instead.", title="Export Failed")
        except Exception as e:
            forms.alert("Excel export failed:\n{}".format(e), title="Export Failed")

    def cancel_click(self, sender, args):
        """Close window without exporting"""
        self.Close()


# -----------------------------
# Helper: Collect Fixtures (Electrical & Lighting)
# -----------------------------
def get_fixtures_from_doc(target_doc):
    """Collect both electrical fixtures and lighting fixtures with category info"""
    fixtures_with_category = []
    try:
        electrical = DB.FilteredElementCollector(target_doc)\
            .OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures)\
            .WhereElementIsNotElementType()\
            .ToElements()

        for fixture in electrical:
            fixtures_with_category.append((fixture, "Electrical"))

        lighting = DB.FilteredElementCollector(target_doc)\
            .OfCategory(DB.BuiltInCategory.OST_LightingFixtures)\
            .WhereElementIsNotElementType()\
            .ToElements()

        for fixture in lighting:
            fixtures_with_category.append((fixture, "Lighting"))

        return fixtures_with_category
    except Exception:
        return []


def get_fixture_system(fixture):
    """Get the system name for a fixture"""
    try:
        # Try to get MEP System parameter
        system_param = fixture.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
        if system_param and system_param.HasValue:
            system_name = system_param.AsString()
            if system_name:
                return system_name

        # Try alternate method - get from electrical systems
        try:
            electrical_systems = fixture.MEPModel.ElectricalSystems
            if electrical_systems and electrical_systems.Size > 0:
                return electrical_systems[0].Name
        except:
            pass

        return "No System"
    except:
        return "No System"


# -----------------------------
# Collect from Host and Linked Models
# -----------------------------
host_fixtures = get_fixtures_from_doc(doc)
linked_fixtures = []

for link_instance in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
    link_doc = link_instance.GetLinkDocument()
    if not link_doc:
        continue
    linked_fixtures.extend(get_fixtures_from_doc(link_doc))


# -----------------------------
# Count Fixtures by Family + Type + System
# -----------------------------
def count_fixtures_detailed(fixtures_with_category):
    """Count fixtures and track category info"""
    # counts[family][type][category] = count
    counts = {}
    for fixture, category in fixtures_with_category:
        try:
            family_name = fixture.Symbol.Family.Name
            type_name = fixture.Name

            if family_name not in counts:
                counts[family_name] = {}
            if type_name not in counts[family_name]:
                counts[family_name][type_name] = {}
            if category not in counts[family_name][type_name]:
                counts[family_name][type_name][category] = 0

            counts[family_name][type_name][category] += 1
        except Exception:
            continue
    return counts


host_counts = count_fixtures_detailed(host_fixtures)
link_counts = count_fixtures_detailed(linked_fixtures)


# -----------------------------
# Combine Host and Linked Counts
# -----------------------------
# Merge host_counts and link_counts into a single structure
all_families = set(host_counts.keys()) | set(link_counts.keys())

if not all_families:
    forms.alert("No electrical or lighting fixtures found in host or linked models.", exitscript=True)


# -----------------------------
# Prepare Hierarchical Data
# -----------------------------
all_data_rows = []  # All rows including headers
total_host = 0
total_linked = 0

# Sort families
sorted_families = sorted(all_families, key=natural_sort_key)

for family_name in sorted_families:
    host_family = host_counts.get(family_name, {})
    link_family = link_counts.get(family_name, {})

    # Get all types for this family
    all_types = set(host_family.keys()) | set(link_family.keys())
    sorted_types = sorted(all_types, key=natural_sort_key)

    # Calculate family total
    family_total = 0
    family_rows = []

    for type_name in sorted_types:
        host_type = host_family.get(type_name, {})
        link_type = link_family.get(type_name, {})

        # Get all categories for this type
        all_categories = set(host_type.keys()) | set(link_type.keys())

        for category in all_categories:
            host_val = host_type.get(category, 0)
            link_val = link_type.get(category, 0)
            total = host_val + link_val

            # Update totals
            total_host += host_val
            total_linked += link_val
            family_total += total

            # Create data row
            row = FixtureCountRow(family_name, type_name, category, host_val, link_val, total)
            family_rows.append(row)

    # Add family header
    header = FamilyHeaderRow(family_name, family_total)
    all_data_rows.append(header)

    # Add family's type rows
    all_data_rows.extend(family_rows)

grand_total = total_host + total_linked

# Get project name for export filenames
project_name = os.path.splitext(os.path.basename(doc.PathName))[0] or "Unnamed_Project"

# Show window with results
window = FixtureCountWindow(all_data_rows, total_host, total_linked, grand_total, project_name)
window.ShowDialog()
