# -*- coding: utf-8 -*-

#######################

# Title : Light Fixture Flip

pyRevit_tool_name = "Light\nFixture\nFlip"

# Description : This tool will flip the work plane of selected light fixtures by type. 

# Author : Chris Berndt

# Date : 01-15-26

#######################

#    _      _____ ____  _____            _____  _____ ______  _____
#   | |    |_   _|  _ \|  __ \     /\   |  __ \|_   _|  ____|/ ____|
#   | |      | | | |_) | |__) |   /  \  | |__) | | | | |__  | (___
#   | |      | | |  _ <|  _  /   / /\ \ |  _  /  | | |  __|  \___ \
#   | |____ _| |_| |_) | | \ \  / ____ \| | \ \ _| |_| |____ ____) |
#   |______|_____|____/|_|  \_\/_/    \_\_|  \_\_____|______|_____/
#

#######################

import sys
import os
import time

# Autodesk Revit Database
from Autodesk.Revit import DB, UI

# Used for creation of ICollection List
from System.Collections.Generic import List

# pyRevit Library
from pyrevit import script
output = script.get_output()

from pyrevit import revit, forms

# Error info
import traceback

# WPF imports for custom window
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows import Window
from System.Windows.Controls import Button, StackPanel, TextBlock, ScrollViewer, Border, Grid, RowDefinition, CheckBox, Expander
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment
from System.Windows.Media import Brushes, SolidColorBrush, Color
from System.Windows.Input import ModifierKeys
import re

#######################

#     _____ _                  ______  _____
#    / ____| |          /\    / _____|/ ____|
#   | |    | |         /  \   \___ \ | (___
#   | |    | |        / /\ \   ___) | \___ \
#   | |____| |____   / ____ \ |____  | ____) |
#    \_____|______| /_/    \_\_____/  |_____/
#

#######################

def natural_sort_key(text):
    """
    Generate a sort key for natural/alphanumeric sorting.
    Converts '1', '2', '10' to sort as 1, 2, 10 instead of 1, 10, 2
    """
    def atoi(text):
        return int(text) if text.isdigit() else text.lower()
    return [atoi(c) for c in re.split(r'(\d+)', text)]


class SelectionMethodWindow(Window):
    """Simple WPF Window for choosing selection method"""

    def __init__(self):
        self.selection_method = None
        self.Title = "Light Fixture Flip"
        self.Width = 400
        self.Height = 200
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.NoResize

        # Main layout
        main_stack = StackPanel()
        main_stack.Margin = Thickness(20)

        # Message
        message = TextBlock()
        message.Text = "How would you like to select light fixtures?"
        message.FontSize = 14
        message.Margin = Thickness(0, 0, 0, 20)
        message.TextWrapping = System.Windows.TextWrapping.Wrap
        message.HorizontalAlignment = HorizontalAlignment.Center
        main_stack.Children.Add(message)

        # Buttons
        button_panel = StackPanel()
        button_panel.HorizontalAlignment = HorizontalAlignment.Center

        # Select by Type button
        type_btn = Button()
        type_btn.Content = "Select by Type"
        type_btn.Width = 200
        type_btn.Height = 35
        type_btn.Margin = Thickness(0, 0, 0, 10)
        type_btn.FontSize = 12
        type_btn.Click += self.select_by_type_click
        button_panel.Children.Add(type_btn)

        # Pick in View button
        pick_btn = Button()
        pick_btn.Content = "Pick in View"
        pick_btn.Width = 200
        pick_btn.Height = 35
        pick_btn.Margin = Thickness(0, 0, 0, 10)
        pick_btn.FontSize = 12
        pick_btn.Click += self.pick_in_view_click
        button_panel.Children.Add(pick_btn)

        # Cancel button
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 200
        cancel_btn.Height = 35
        cancel_btn.FontSize = 12
        cancel_btn.Click += self.cancel_click
        button_panel.Children.Add(cancel_btn)

        main_stack.Children.Add(button_panel)
        self.Content = main_stack

    def select_by_type_click(self, sender, args):
        """Select by Type option"""
        self.selection_method = 'Select by Type'
        self.DialogResult = True
        self.Close()

    def pick_in_view_click(self, sender, args):
        """Pick in View option"""
        self.selection_method = 'Pick in View'
        self.DialogResult = True
        self.Close()

    def cancel_click(self, sender, args):
        """Cancel option"""
        self.selection_method = None
        self.DialogResult = False
        self.Close()


class FixtureTypeData:
    """Class to store fixture type selection state"""
    def __init__(self, family_name, type_name, full_name, fixtures):
        self.family_name = family_name
        self.type_name = type_name
        self.full_name = full_name  # "Family : Type"
        self.fixtures = fixtures
        self.is_selected = False
        self.row_border = None
        self.is_visible = True  # For filtering
        self.is_header = False  # True for family headers, False for type rows


class FamilyHeaderData:
    """Class to store family header information"""
    def __init__(self, family_name):
        self.family_name = family_name
        self.is_header = True
        self.is_visible = True
        self.row_border = None


class FixtureTypeSelectionWindow(Window):
    """WPF Window for selecting fixture types with row-based selection"""

    def __init__(self, type_dict, host_dict):
        self.type_dict = type_dict
        self.host_dict = host_dict  # Dictionary mapping host names to host info
        self.selected_fixtures = []
        self.Title = "Select Light Fixture Types to Flip Work Plane"
        self.Width = 600
        self.Height = 650
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        # Track selection state
        self.last_selected_index = -1
        self.is_dragging = False
        self.drag_start_selected = False

        # Track host checkboxes
        self.host_checkboxes = {}  # Dictionary mapping host names to checkbox controls

        # Organize data by family
        families = {}
        for key, data in type_dict.items():
            family_name = data['family']
            type_name = data['type']
            full_name = data['full_name']
            fixtures = data['fixtures']

            if family_name not in families:
                families[family_name] = []
            families[family_name].append(FixtureTypeData(family_name, type_name, full_name, fixtures))

        # Create display list with family headers and indented types
        self.all_display_items = []  # Mix of FamilyHeaderData and FixtureTypeData
        self.fixture_type_data = []  # Only FixtureTypeData items (for selection logic)

        # Sort families
        sorted_families = sorted(families.keys(), key=natural_sort_key)

        for family_name in sorted_families:
            # Add family header
            header = FamilyHeaderData(family_name)
            self.all_display_items.append(header)

            # Sort types within family
            types_in_family = sorted(families[family_name], key=lambda x: natural_sort_key(x.type_name))

            # Add types under family
            for type_data in types_in_family:
                self.all_display_items.append(type_data)
                self.fixture_type_data.append(type_data)

        # Main layout using Grid for dynamic resizing
        main_grid = Grid()
        main_grid.Margin = Thickness(15)

        # Define rows: Filter, Instructions, ScrollViewer (flexible), Buttons
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[0].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[1].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[2].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)  # Flexible

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[3].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        # Host filter - collapsible expander
        host_expander = Expander()
        host_expander.Header = "Filter by Host (select multiple)"
        host_expander.IsExpanded = False  # Start collapsed
        host_expander.Margin = Thickness(0, 0, 0, 10)
        host_expander.FontWeight = System.Windows.FontWeights.Bold

        # Create scrollable container for host checkboxes
        host_scroll = ScrollViewer()
        host_scroll.MaxHeight = 120
        host_scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        host_scroll.HorizontalAlignment = HorizontalAlignment.Left
        host_scroll.Margin = Thickness(20, 5, 0, 5)  # Indent content slightly

        host_stack = StackPanel()
        host_scroll.Content = host_stack

        # Add "All Hosts" checkbox
        all_checkbox = CheckBox()
        all_checkbox.Content = "All Hosts"
        all_checkbox.IsChecked = True
        all_checkbox.Margin = Thickness(0, 2, 0, 2)
        all_checkbox.FontWeight = System.Windows.FontWeights.Normal  # Normal weight for checkbox content
        all_checkbox.Tag = None
        all_checkbox.Checked += self.host_filter_changed
        all_checkbox.Unchecked += self.host_filter_changed
        host_stack.Children.Add(all_checkbox)
        self.host_checkboxes[None] = all_checkbox

        # Add individual host checkboxes sorted naturally
        sorted_host_names = sorted(host_dict.keys(), key=natural_sort_key)
        for host_name in sorted_host_names:
            checkbox = CheckBox()
            checkbox.Content = host_name
            checkbox.IsChecked = False
            checkbox.Margin = Thickness(20, 2, 0, 2)  # Indent individual hosts
            checkbox.FontWeight = System.Windows.FontWeights.Normal  # Normal weight for checkbox content
            checkbox.Tag = host_name
            checkbox.Checked += self.host_filter_changed
            checkbox.Unchecked += self.host_filter_changed
            host_stack.Children.Add(checkbox)
            self.host_checkboxes[host_name] = checkbox

        host_expander.Content = host_scroll

        Grid.SetRow(host_expander, 0)
        main_grid.Children.Add(host_expander)

        # Instructions
        instructions = TextBlock()
        instructions.Text = "Click to select fixture types. Use Ctrl+Click to add/remove, Shift+Click to select range, or drag to select multiple."
        instructions.Margin = Thickness(0, 0, 0, 10)
        instructions.TextWrapping = System.Windows.TextWrapping.Wrap
        instructions.FontSize = 11
        Grid.SetRow(instructions, 1)
        main_grid.Children.Add(instructions)

        # Scrollable list of fixture types (dynamic height)
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto

        self.type_panel = StackPanel()
        scroll.Content = self.type_panel

        # Create all rows (headers + types)
        for item in self.all_display_items:
            if item.is_header:
                row = self.create_family_header_row(item)
            else:
                row = self.create_type_row(item)
            self.type_panel.Children.Add(row)

        Grid.SetRow(scroll, 2)
        main_grid.Children.Add(scroll)

        # Buttons
        button_panel = StackPanel()
        button_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Right
        button_panel.Margin = Thickness(0, 10, 0, 0)

        select_all_btn = Button()
        select_all_btn.Content = "Select All"
        select_all_btn.Width = 80
        select_all_btn.Margin = Thickness(0, 0, 5, 0)
        select_all_btn.Click += self.select_all_click
        button_panel.Children.Add(select_all_btn)

        deselect_all_btn = Button()
        deselect_all_btn.Content = "Deselect All"
        deselect_all_btn.Width = 90
        deselect_all_btn.Margin = Thickness(0, 0, 5, 0)
        deselect_all_btn.Click += self.deselect_all_click
        button_panel.Children.Add(deselect_all_btn)

        flip_btn = Button()
        flip_btn.Content = "Flip Selected"
        flip_btn.Width = 100
        flip_btn.Margin = Thickness(0, 0, 5, 0)
        flip_btn.Click += self.flip_click
        button_panel.Children.Add(flip_btn)

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Click += self.cancel_click
        button_panel.Children.Add(cancel_btn)

        Grid.SetRow(button_panel, 3)
        main_grid.Children.Add(button_panel)

        # Add mouse event handlers for drag selection
        self.MouseMove += self.window_mouse_move
        self.MouseLeftButtonUp += self.window_mouse_up

        self.Content = main_grid

    def host_filter_changed(self, sender, args):
        """Handle host filter checkbox changes"""
        changed_checkbox = sender
        changed_host = changed_checkbox.Tag

        # Special handling for "All Hosts" checkbox
        if changed_host is None:  # "All Hosts" checkbox
            if changed_checkbox.IsChecked:
                # When "All Hosts" is checked, uncheck all other checkboxes
                for host_name, checkbox in self.host_checkboxes.items():
                    if host_name is not None:  # Not the "All Hosts" checkbox
                        checkbox.IsChecked = False
        else:
            # When a specific host is checked, uncheck "All Hosts"
            if changed_checkbox.IsChecked:
                all_hosts_checkbox = self.host_checkboxes.get(None)
                if all_hosts_checkbox:
                    all_hosts_checkbox.IsChecked = False

        # Get all selected hosts
        selected_hosts = [
            host_name for host_name, checkbox in self.host_checkboxes.items()
            if checkbox.IsChecked
        ]

        # If no hosts are selected, check "All Hosts" automatically
        if not selected_hosts:
            all_hosts_checkbox = self.host_checkboxes.get(None)
            if all_hosts_checkbox:
                all_hosts_checkbox.IsChecked = True
            selected_hosts = [None]

        # Update visibility for all type data
        for type_data in self.fixture_type_data:
            if None in selected_hosts:
                # "All Hosts" is selected - show all
                type_data.is_visible = True
            else:
                # Check if any fixture in this type has any of the selected hosts
                type_data.is_visible = any(
                    self.get_fixture_host_name(fixture) in selected_hosts
                    for fixture in type_data.fixtures
                )

        # Update visibility for family headers (show if any child type is visible)
        for item in self.all_display_items:
            if item.is_header:
                # Check if any types under this family are visible
                family_name = item.family_name
                item.is_visible = any(
                    td.is_visible and td.family_name == family_name
                    for td in self.fixture_type_data
                )

        # Rebuild panel with visible items
        self.type_panel.Children.Clear()
        for item in self.all_display_items:
            if item.is_visible:
                if item.is_header:
                    row = self.create_family_header_row(item)
                else:
                    row = self.create_type_row(item)
                self.type_panel.Children.Add(row)

    def get_fixture_host_name(self, fixture):
        """Get the host name for a fixture"""
        try:
            host = fixture.Host
            if host:
                # Get host element name and category
                host_category = host.Category.Name if host.Category else "Unknown"
                host_name = DB.Element.Name.GetValue(host) if hasattr(host, 'Name') else str(host.Id.IntegerValue)
                return "{}: {}".format(host_category, host_name)
        except:
            pass
        return "No Host"

    def create_family_header_row(self, header_data):
        """Create a non-selectable header row for a family"""
        row_border = Border()
        row_border.BorderThickness = Thickness(0, 1, 0, 0)
        row_border.BorderBrush = Brushes.Gray
        row_border.Padding = Thickness(5, 8, 5, 5)
        row_border.Margin = Thickness(0, 5, 0, 2)
        row_border.Background = SolidColorBrush(Color.FromRgb(240, 240, 240))  # Light gray

        # Store reference for updates
        header_data.row_border = row_border

        # Family name text (bold)
        text_block = TextBlock()
        text_block.Text = header_data.family_name
        text_block.FontWeight = System.Windows.FontWeights.Bold
        text_block.FontSize = 12
        text_block.VerticalAlignment = VerticalAlignment.Center
        row_border.Child = text_block

        return row_border

    def create_type_row(self, type_data):
        """Create a selectable row for a fixture type (indented under family)"""
        row_border = Border()
        row_border.BorderThickness = Thickness(0)
        row_border.Padding = Thickness(25, 5, 5, 5)  # Indent from left
        row_border.Margin = Thickness(0, 1, 0, 1)
        row_border.Tag = type_data
        row_border.MouseLeftButtonDown += self.row_click

        # Store reference for updates
        type_data.row_border = row_border

        # Update appearance based on selection
        self.update_row_appearance(type_data)

        # Type name text - show only type name (not full name) with count
        text_block = TextBlock()
        text_block.Text = "{} ({} fixtures)".format(type_data.type_name, len(type_data.fixtures))
        text_block.VerticalAlignment = VerticalAlignment.Center
        row_border.Child = text_block

        return row_border

    def row_click(self, sender, args):
        """Handle row click with Shift/Ctrl support and start drag selection"""
        clicked_row = sender
        clicked_type_data = clicked_row.Tag

        # Only handle type rows (not family headers)
        if not isinstance(clicked_type_data, FixtureTypeData):
            return

        # Get list of currently visible type items (not headers)
        visible_items = [td for td in self.fixture_type_data if td.is_visible]

        # Get the index of clicked row in visible list
        clicked_index = visible_items.index(clicked_type_data)

        # Get modifier keys
        modifiers = System.Windows.Input.Keyboard.Modifiers

        if modifiers == ModifierKeys.Control:
            # Ctrl+Click: Toggle selection of clicked row
            clicked_type_data.is_selected = not clicked_type_data.is_selected
            self.last_selected_index = clicked_index
            self.update_row_appearance(clicked_type_data)
            # Start drag in toggle mode
            self.is_dragging = True
            self.drag_start_selected = clicked_type_data.is_selected

        elif modifiers == ModifierKeys.Shift and self.last_selected_index >= 0:
            # Shift+Click: Select range from last selected to current in visible items
            start_index = min(self.last_selected_index, clicked_index)
            end_index = max(self.last_selected_index, clicked_index)

            for i in range(start_index, end_index + 1):
                visible_items[i].is_selected = True
                self.update_row_appearance(visible_items[i])

        else:
            # Normal click: Select only this row, deselect others
            for type_data in self.fixture_type_data:
                type_data.is_selected = False
                self.update_row_appearance(type_data)

            clicked_type_data.is_selected = True
            self.last_selected_index = clicked_index
            self.update_row_appearance(clicked_type_data)
            # Start drag in select mode
            self.is_dragging = True
            self.drag_start_selected = True

        args.Handled = True

    def window_mouse_move(self, sender, args):
        """Handle mouse move for drag selection"""
        if not self.is_dragging:
            return

        # Get the element under the mouse
        element = self.InputHitTest(args.GetPosition(self))
        if element is None:
            return

        # Find the row border by traversing up the visual tree
        current = element
        while current is not None:
            if isinstance(current, Border) and hasattr(current, 'Tag'):
                # Only handle type rows, not headers
                if isinstance(current.Tag, FixtureTypeData):
                    type_data = current.Tag
                    type_data.is_selected = self.drag_start_selected
                    self.update_row_appearance(type_data)
                break
            # Move up the visual tree
            try:
                current = System.Windows.Media.VisualTreeHelper.GetParent(current)
            except:
                break

    def window_mouse_up(self, sender, args):
        """Handle mouse up to end drag selection"""
        self.is_dragging = False

    def update_row_appearance(self, type_data):
        """Update row background based on selection state"""
        if type_data.row_border is None:
            return

        if type_data.is_selected:
            # Selected row - highlight with blue
            type_data.row_border.Background = SolidColorBrush(Color.FromRgb(173, 216, 230))  # Light blue
        else:
            # Not selected - white
            type_data.row_border.Background = Brushes.White

    def select_all_click(self, sender, args):
        """Select all fixture types"""
        for type_data in self.fixture_type_data:
            type_data.is_selected = True
            self.update_row_appearance(type_data)

    def deselect_all_click(self, sender, args):
        """Deselect all fixture types"""
        for type_data in self.fixture_type_data:
            type_data.is_selected = False
            self.update_row_appearance(type_data)

    def flip_click(self, sender, args):
        """Flip selected fixture types"""
        self.selected_fixtures = []
        for type_data in self.fixture_type_data:
            if type_data.is_selected:
                self.selected_fixtures.extend(type_data.fixtures)

        if self.selected_fixtures:
            self.DialogResult = True
            self.Close()

    def cancel_click(self, sender, args):
        """Cancel selection"""
        self.selected_fixtures = []
        self.DialogResult = False
        self.Close()


#######################

#    ______ _    _ _   _  _____ _______ _____ ____  _   _  _____
#   |  ____| |  | | \ | |/ ____|__   __|_   _/ __ \| \ | |/ ____|
#   | |__  | |  | |  \| | |       | |    | || |  | |  \| | (___
#   |  __| | |  | | . ` | |       | |    | || |  | | . ` |\___ \
#   | |    | |__| | |\  | |____   | |   _| || |__| | |\  |____) |
#   |_|     \____/|_| \_|\_____|  |_|  |_____\____/|_| \_|_____/
#

#######################

# Popup to show errors
def show_error_dialog(title="Error", message="", details=""):
    dialog = UI.TaskDialog(title)
    dialog.MainIcon = UI.TaskDialogIcon.TaskDialogIconError
    dialog.MainInstruction = message
    dialog.MainContent = "Check details below."
    dialog.ExpandedContent = details
    dialog.Show()

# Get all light fixtures in the model
def get_all_light_fixtures(doc):
    collector = DB.FilteredElementCollector(doc)\
        .OfCategory(DB.BuiltInCategory.OST_LightingFixtures)\
        .WhereElementIsNotElementType()
    return list(collector)

# Get light fixture types from elements and collect hosts
def get_fixture_types_and_hosts(fixtures):
    """
    Returns a tuple of (type_dict, host_dict)
    type_dict: dictionary with keys as unique type IDs, values as dict with 'family', 'type', 'full_name', 'fixtures'
    host_dict: dictionary of host names to host elements
    """
    type_dict = {}
    host_dict = {}

    for fixture in fixtures:
        # Get type information
        type_id = fixture.GetTypeId()
        if type_id != DB.ElementId.InvalidElementId:
            fixture_type = fixture.Document.GetElement(type_id)
            if fixture_type:
                # Get family name directly from the family
                family_name = fixture_type.FamilyName

                # Get type name from the type element
                type_name = DB.Element.Name.GetValue(fixture_type)

                # Create full name for display/reference
                full_name = "{} : {}".format(family_name, type_name)

                # Use a unique key (type element ID) to avoid duplicates
                unique_key = type_id.IntegerValue

                # Add to type dictionary
                if unique_key not in type_dict:
                    type_dict[unique_key] = {
                        'family': family_name,
                        'type': type_name,
                        'full_name': full_name,
                        'fixtures': []
                    }
                type_dict[unique_key]['fixtures'].append(fixture)

        # Collect host information
        try:
            host = fixture.Host
            if host:
                # Get host element name and category
                host_category = host.Category.Name if host.Category else "Unknown"
                host_name = DB.Element.Name.GetValue(host) if hasattr(host, 'Name') else str(host.Id.IntegerValue)
                full_host_name = "{}: {}".format(host_category, host_name)

                if full_host_name not in host_dict:
                    host_dict[full_host_name] = host
            else:
                # No host - add to dictionary
                if "No Host" not in host_dict:
                    host_dict["No Host"] = None
        except:
            # Error getting host - add to "No Host" category
            if "No Host" not in host_dict:
                host_dict["No Host"] = None

    return type_dict, host_dict

# Flip work plane for a single element
def flip_work_plane(element):
    """
    Flips the work plane of a family instance.
    Toggles the IsWorkPlaneFlipped property.
    """
    if isinstance(element, DB.FamilyInstance):
        # Check if the element can be flipped
        if element.CanFlipWorkPlane:
            # Toggle the work plane flip state
            element.IsWorkPlaneFlipped = not element.IsWorkPlaneFlipped
            return True
    return False

#######################

#   __      __     _____  _____ ____  _      ______  _____
#   \ \    / /\   |  __ \|_   _|  _ \| |    |  ____|/ ____|
#    \ \  / /  \  | |__) | | | | |_) | |    | |__  | (___
#     \ \/ / /\ \ |  _  /  | | |  _ <| |    |  __|  \___ \
#      \  / ____ \| | \ \ _| |_| |_) | |____| |____ ____) |
#       \/_/    \_\_|  \_\_____|____/|______|______|_____/
#

#######################

start_time = time.time()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#######################

#     _____ ____  _____  ______
#    / ____/ __ \|  __ \|  ____|
#   | |   | |  | | |  | | |__
#   | |   | |  | | |  | |  __|
#   | |___| |__| | |__| | |____
#    \_____\____/|_____/|______|
#

#######################

try:
    fixtures_to_flip = []

    # Check for pre-selected light fixtures
    current_selection = uidoc.Selection.GetElementIds()
    if current_selection.Count > 0:
        for elem_id in current_selection:
            element = doc.GetElement(elem_id)
            if element and element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_LightingFixtures):
                fixtures_to_flip.append(element)

    # If no light fixtures pre-selected, show selection options
    if not fixtures_to_flip:
        # Show custom selection method window
        method_window = SelectionMethodWindow()
        result = method_window.ShowDialog()

        # If user cancelled or closed window, exit
        if not result or method_window.selection_method is None:
            pass  # Exit gracefully
        elif method_window.selection_method == 'Select by Type':
            # Get all light fixtures in the model
            all_fixtures = get_all_light_fixtures(doc)

            if not all_fixtures:
                forms.alert("No light fixtures found in the model.", title="Light Fixture Flip")
            else:
                # Get fixture types dictionary and hosts
                type_dict, host_dict = get_fixture_types_and_hosts(all_fixtures)

                if not type_dict:
                    forms.alert("No valid light fixture types found.", title="Light Fixture Flip")
                else:
                    # Show custom fixture type selection window
                    type_selection_window = FixtureTypeSelectionWindow(type_dict, host_dict)
                    result = type_selection_window.ShowDialog()

                    if result and type_selection_window.selected_fixtures:
                        fixtures_to_flip = type_selection_window.selected_fixtures

        elif method_window.selection_method == 'Pick in View':
            # Let user pick fixtures in the view
            try:
                # Create a selection filter for lighting fixtures
                selection_filter = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_LightingFixtures)

                selected_refs = uidoc.Selection.PickObjects(
                    UI.Selection.ObjectType.Element,
                    'Select light fixtures to flip (press Finish when done)'
                )

                for ref in selected_refs:
                    element = doc.GetElement(ref.ElementId)
                    if element and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_LightingFixtures):
                        fixtures_to_flip.append(element)

                if not fixtures_to_flip:
                    forms.alert("No light fixtures were selected.", title="Light Fixture Flip")

            except Exception as pick_ex:
                # User cancelled the selection
                if "cancelled" in str(pick_ex).lower() or "canceled" in str(pick_ex).lower():
                    pass  # User cancelled, just exit gracefully
                else:
                    raise pick_ex

    # Flip the fixtures if any were selected
    if fixtures_to_flip:
        # Count for results
        flipped_count = 0
        skipped_count = 0

        # Start Transaction
        t = DB.Transaction(doc, 'Flip Light Fixture Work Planes')
        t.Start()

        for fixture in fixtures_to_flip:
            if flip_work_plane(fixture):
                flipped_count += 1
            else:
                skipped_count += 1

        # End Transaction
        t.Commit()

        # Only show dialog if no fixtures were flipped
        if flipped_count == 0:
            forms.alert("No fixtures could be flipped (work plane flip not supported).", title="Light Fixture Flip")

except Exception as ex:
    details = traceback.format_exc()
    show_error_dialog(title="Light Fixture Flip Error", message="An unexpected error occurred.", details=details)

#######################

#     ____  _    _ _______ _____  _    _ _______  ___      ____   _____
#    / __ \| |  | |__   __|  __ \| |  | |__   __|/ / |    / __ \ / ____|
#   | |  | | |  | |  | |  | |__) | |  | |  | |  / /| |   | |  | | |  __
#   | |  | | |  | |  | |  |  ___/| |  | |  | | / / | |   | |  | | | |_ |
#   | |__| | |__| |  | |  | |    | |__| |  | |/ /  | |___| |__| | |__| |
#    \____/ \____/   |_|  |_|     \____/   |_/_/   |______\____/ \_____|
#

#######################
