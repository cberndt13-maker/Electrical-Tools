# -*- coding: utf-8 -*-
"""
Grid Bubble Control
Control visibility of grid bubbles for individual gridlines in the active view.
"""

__title__ = "Gridline\nBubbles"
__author__ = "DEEM"

from pyrevit import revit, DB, UI, forms, script
from System.Collections.Generic import List
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows import Window, Application
from System.Windows.Controls import CheckBox, Button, StackPanel, WrapPanel, TextBlock, ScrollViewer, Grid, Border, ComboBox
from System.Windows.Controls import RowDefinition, ColumnDefinition
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment
from System.Windows.Media import Brushes, SolidColorBrush, Color
from System.Windows.Input import ModifierKeys
from System.Windows.Threading import DispatcherTimer
from System import Func, TimeSpan
import re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView


def natural_sort_key(text):
    """
    Generate a sort key for natural/alphanumeric sorting.
    Converts '1', '2', '10' to sort as 1, 2, 10 instead of 1, 10, 2
    """
    def atoi(text):
        return int(text) if text.isdigit() else text.lower()

    return [atoi(c) for c in re.split(r'(\d+)', text)]


class GridBubbleData:
    """Class to store grid and its bubble visibility settings"""
    def __init__(self, grid, view):
        self.grid = grid
        self.view = view
        self.name = grid.Name

        # Get current bubble visibility for both ends
        try:
            # Check if bubbles are visible at each end
            self.end0_visible = grid.IsBubbleVisibleInView(DB.DatumEnds.End0, view)
            self.end1_visible = grid.IsBubbleVisibleInView(DB.DatumEnds.End1, view)
        except:
            # Fallback to checking extent type
            datum_refs_2d = grid.GetDatumExtentTypeInView(DB.DatumEnds.End0, view)
            self.end0_visible = datum_refs_2d == DB.DatumExtentType.ViewSpecific

            datum_refs_2d = grid.GetDatumExtentTypeInView(DB.DatumEnds.End1, view)
            self.end1_visible = datum_refs_2d == DB.DatumExtentType.ViewSpecific

        # Determine orientation (vertical or horizontal)
        curve = grid.Curve
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)

        dx = abs(end.X - start.X)
        dy = abs(end.Y - start.Y)

        self.is_vertical = dx < dy

        # Store checkboxes references for all directions
        self.check_left = None
        self.check_right = None
        self.check_top = None
        self.check_bottom = None

        # Row selection state (no checkbox needed)
        self.is_selected = False

        # Row grid reference for appearance updates
        self.row_grid = None


class ViewData:
    """Class to store view selection state"""
    def __init__(self, view_name, view):
        self.view_name = view_name
        self.view = view
        self.is_selected = False
        self.row_border = None  # Reference to the row UI element


class ViewSelectionWindow(Window):
    """WPF Window for selecting views with type filtering"""

    def __init__(self, views_dict):
        self.views_dict = views_dict  # Dict of {view_name: view}
        self.selected_views = []
        self.Title = "Select Views to Apply Grid Bubble Settings"
        self.Width = 500
        self.Height = 600
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        # Track selection state
        self.last_selected_index = -1
        self.is_dragging = False
        self.drag_start_selected = False

        # Create view data objects
        self.all_view_data = []
        for view_name, view in views_dict.items():
            self.all_view_data.append(ViewData(view_name, view))

        # Organize views by type
        self.views_by_type = {}
        for view_data in self.all_view_data:
            view_type = str(view_data.view.ViewType)
            if view_type not in self.views_by_type:
                self.views_by_type[view_type] = []
            self.views_by_type[view_type].append(view_data)

        # Main layout
        main_grid = Grid()
        main_grid.Margin = Thickness(15)

        # Define rows
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[0].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[1].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[2].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[3].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        # Instructions
        instructions = TextBlock()
        instructions.Text = "Click to select views. Use Ctrl+Click to add/remove, Shift+Click to select range, or drag to select multiple."
        instructions.Margin = Thickness(0, 0, 0, 10)
        instructions.TextWrapping = System.Windows.TextWrapping.Wrap
        instructions.FontSize = 11
        Grid.SetRow(instructions, 0)
        main_grid.Children.Add(instructions)

        # Filter dropdown
        filter_panel = StackPanel()
        filter_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        filter_panel.Margin = Thickness(0, 0, 0, 10)

        filter_label = TextBlock()
        filter_label.Text = "Filter by View Type: "
        filter_label.VerticalAlignment = VerticalAlignment.Center
        filter_label.Margin = Thickness(0, 0, 10, 0)
        filter_panel.Children.Add(filter_label)

        self.filter_combo = ComboBox()
        self.filter_combo.Width = 200
        self.filter_combo.Items.Add("All View Types")
        for view_type in sorted(self.views_by_type.keys()):
            self.filter_combo.Items.Add(view_type)
        self.filter_combo.SelectedIndex = 0
        self.filter_combo.SelectionChanged += self.filter_changed
        filter_panel.Children.Add(self.filter_combo)

        Grid.SetRow(filter_panel, 1)
        main_grid.Children.Add(filter_panel)

        # Scrollable list of views
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto

        self.view_panel = StackPanel()
        scroll.Content = self.view_panel

        Grid.SetRow(scroll, 2)
        main_grid.Children.Add(scroll)

        # Buttons
        button_panel = WrapPanel()
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

        apply_btn = Button()
        apply_btn.Content = "Apply to Selected"
        apply_btn.Width = 110
        apply_btn.Margin = Thickness(0, 0, 5, 0)
        apply_btn.Click += self.apply_click
        button_panel.Children.Add(apply_btn)

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

        # Initial population
        self.populate_view_list()

    def populate_view_list(self):
        """Populate the view list based on current filter"""
        self.view_panel.Children.Clear()
        self.current_view_data = []

        selected_filter = self.filter_combo.SelectedItem

        # Get views to display
        if selected_filter == "All View Types":
            self.current_view_data = self.all_view_data[:]
        else:
            self.current_view_data = self.views_by_type.get(selected_filter, [])[:]

        # Sort by name
        self.current_view_data = sorted(self.current_view_data, key=lambda x: natural_sort_key(x.view_name))

        # Create view rows
        for view_data in self.current_view_data:
            row = self.create_view_row(view_data)
            self.view_panel.Children.Add(row)

    def create_view_row(self, view_data):
        """Create a selectable row for a view"""
        row_border = Border()
        row_border.BorderThickness = Thickness(0)
        row_border.Padding = Thickness(5)
        row_border.Margin = Thickness(0, 1, 0, 1)
        row_border.Tag = view_data
        row_border.MouseLeftButtonDown += self.row_click

        # Store reference for updates
        view_data.row_border = row_border

        # Update appearance based on selection
        self.update_row_appearance(view_data)

        # View name text
        text_block = TextBlock()
        text_block.Text = view_data.view_name
        text_block.VerticalAlignment = VerticalAlignment.Center
        row_border.Child = text_block

        return row_border

    def row_click(self, sender, args):
        """Handle row click with Shift/Ctrl support and start drag selection"""
        clicked_row = sender
        clicked_view_data = clicked_row.Tag

        # Get the index of clicked row in the current list
        clicked_index = self.current_view_data.index(clicked_view_data)

        # Get modifier keys
        modifiers = System.Windows.Input.Keyboard.Modifiers

        if modifiers == ModifierKeys.Control:
            # Ctrl+Click: Toggle selection of clicked row
            clicked_view_data.is_selected = not clicked_view_data.is_selected
            self.last_selected_index = clicked_index
            self.update_row_appearance(clicked_view_data)
            # Start drag in toggle mode
            self.is_dragging = True
            self.drag_start_selected = clicked_view_data.is_selected

        elif modifiers == ModifierKeys.Shift and self.last_selected_index >= 0:
            # Shift+Click: Select range from last selected to current
            start_index = min(self.last_selected_index, clicked_index)
            end_index = max(self.last_selected_index, clicked_index)

            for i in range(start_index, end_index + 1):
                self.current_view_data[i].is_selected = True
                self.update_row_appearance(self.current_view_data[i])

        else:
            # Normal click: Select only this row, deselect others
            for view_data in self.current_view_data:
                view_data.is_selected = False
                self.update_row_appearance(view_data)

            clicked_view_data.is_selected = True
            self.last_selected_index = clicked_index
            self.update_row_appearance(clicked_view_data)
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
            if isinstance(current, Border) and hasattr(current, 'Tag') and isinstance(current.Tag, ViewData):
                # Found a row border
                view_data = current.Tag
                view_data.is_selected = self.drag_start_selected
                self.update_row_appearance(view_data)
                break
            # Move up the visual tree
            try:
                current = System.Windows.Media.VisualTreeHelper.GetParent(current)
            except:
                break

    def window_mouse_up(self, sender, args):
        """Handle mouse up to end drag selection"""
        self.is_dragging = False

    def update_row_appearance(self, view_data):
        """Update row background based on selection state"""
        if view_data.row_border is None:
            return

        if view_data.is_selected:
            # Selected row - highlight with blue
            view_data.row_border.Background = SolidColorBrush(Color.FromRgb(173, 216, 230))  # Light blue
        else:
            # Not selected - white
            view_data.row_border.Background = Brushes.White

    def filter_changed(self, sender, args):
        """Handle filter selection change"""
        self.populate_view_list()

    def select_all_click(self, sender, args):
        """Select all visible views"""
        for view_data in self.current_view_data:
            view_data.is_selected = True
            self.update_row_appearance(view_data)

    def deselect_all_click(self, sender, args):
        """Deselect all visible views"""
        for view_data in self.current_view_data:
            view_data.is_selected = False
            self.update_row_appearance(view_data)

    def apply_click(self, sender, args):
        """Apply to selected views"""
        self.selected_views = []
        for view_data in self.all_view_data:
            if view_data.is_selected:
                self.selected_views.append(view_data.view)

        if self.selected_views:
            self.DialogResult = True
            self.Close()

    def cancel_click(self, sender, args):
        """Cancel selection"""
        self.selected_views = []
        self.DialogResult = False
        self.Close()


class GridBubbleWindow(Window):
    """WPF Window for grid bubble control with table layout"""

    def __init__(self, grid_data_list, view):
        self.grid_data_list = grid_data_list
        self.view = view
        self.Title = "Grid Bubble Control"
        self.Width = 650
        self.Height = 600
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        # Track last selected index for Shift+click range selection
        self.last_selected_index = -1

        # Track drag selection state
        self.is_dragging = False
        self.drag_start_selected = False

        # Create sorted list for display (used for selection logic)
        self.sorted_grid_data_list = sorted(grid_data_list, key=lambda x: natural_sort_key(x.name))

        # Main grid for dynamic layout
        main_grid = Grid()
        main_grid.Margin = Thickness(15)

        # Define rows: Instructions, Header, Table (flexible), Buttons
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[0].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[1].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[2].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)  # Flexible

        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[3].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        # Instructions
        instructions = TextBlock()
        instructions.Text = "Click rows to select. Use Ctrl+Click to add/remove, Shift+Click to select range. Checking/unchecking any checkbox applies to all selected rows. Check to show bubbles, Uncheck to hide. Grayed checkboxes don't apply to grid orientation."
        instructions.Margin = Thickness(0, 0, 0, 10)
        instructions.TextWrapping = System.Windows.TextWrapping.Wrap
        instructions.FontSize = 11
        Grid.SetRow(instructions, 0)
        main_grid.Children.Add(instructions)

        # Create sticky header (outside scroll viewer)
        header_grid = self.create_header()
        Grid.SetRow(header_grid, 1)
        main_grid.Children.Add(header_grid)

        # Scroll viewer for data rows only (flexible height)
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto

        # Create the data rows
        data_panel = self.create_data_rows()
        scroll.Content = data_panel
        Grid.SetRow(scroll, 2)
        main_grid.Children.Add(scroll)

        # Buttons panel - using Grid to position left and right groups
        button_grid = Grid()
        button_grid.Margin = Thickness(0, 10, 0, 0)

        # Two columns: left (for Apply to Views) and right (for other buttons)
        button_grid.ColumnDefinitions.Add(ColumnDefinition())
        button_grid.ColumnDefinitions[0].Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        button_grid.ColumnDefinitions.Add(ColumnDefinition())
        button_grid.ColumnDefinitions[1].Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)

        # Left side - Apply to Views button
        left_panel = StackPanel()
        left_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        left_panel.HorizontalAlignment = HorizontalAlignment.Left

        apply_views_btn = Button()
        apply_views_btn.Content = "Apply to Views"
        apply_views_btn.Width = 120
        apply_views_btn.Click += self.apply_to_views_click
        left_panel.Children.Add(apply_views_btn)

        Grid.SetColumn(left_panel, 0)
        button_grid.Children.Add(left_panel)

        # Right side - other buttons
        right_panel = WrapPanel()
        right_panel.HorizontalAlignment = HorizontalAlignment.Right

        # Toggle Selected button (swap left/right or top/bottom)
        toggle_selected_btn = Button()
        toggle_selected_btn.Content = "Toggle Selected"
        toggle_selected_btn.Width = 110
        toggle_selected_btn.Margin = Thickness(0, 0, 5, 0)
        toggle_selected_btn.Click += self.toggle_selected_click
        right_panel.Children.Add(toggle_selected_btn)

        # Uncheck All button
        uncheck_all_btn = Button()
        uncheck_all_btn.Content = "Uncheck All"
        uncheck_all_btn.Width = 100
        uncheck_all_btn.Margin = Thickness(0, 0, 5, 0)
        uncheck_all_btn.Click += self.uncheck_all_click
        right_panel.Children.Add(uncheck_all_btn)

        # Apply button (applies but doesn't close)
        apply_btn = Button()
        apply_btn.Content = "Apply"
        apply_btn.Width = 80
        apply_btn.Margin = Thickness(0, 0, 5, 0)
        apply_btn.Click += self.apply_click
        right_panel.Children.Add(apply_btn)

        # OK button (closes window)
        ok_btn = Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 80
        ok_btn.Margin = Thickness(0, 0, 5, 0)
        ok_btn.Click += self.ok_click
        right_panel.Children.Add(ok_btn)

        # Cancel button
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Click += self.cancel_click
        right_panel.Children.Add(cancel_btn)

        Grid.SetColumn(right_panel, 1)
        button_grid.Children.Add(right_panel)

        Grid.SetRow(button_grid, 3)
        main_grid.Children.Add(button_grid)

        # Add mouse event handlers for drag selection
        self.MouseMove += self.window_mouse_move
        self.MouseLeftButtonUp += self.window_mouse_up

        self.Content = main_grid

    def create_header(self):
        """Create sticky header row"""
        header_grid = Grid()
        header_grid.Background = Brushes.LightGray
        header_grid.Margin = Thickness(0, 0, 0, 2)

        # Define columns (removed Select column)
        header_grid.ColumnDefinitions.Add(ColumnDefinition())  # Grid Name
        header_grid.ColumnDefinitions[0].Width = System.Windows.GridLength(120)

        header_grid.ColumnDefinitions.Add(ColumnDefinition())  # Left
        header_grid.ColumnDefinitions[1].Width = System.Windows.GridLength(80)

        header_grid.ColumnDefinitions.Add(ColumnDefinition())  # Right
        header_grid.ColumnDefinitions[2].Width = System.Windows.GridLength(80)

        header_grid.ColumnDefinitions.Add(ColumnDefinition())  # Top
        header_grid.ColumnDefinitions[3].Width = System.Windows.GridLength(80)

        header_grid.ColumnDefinitions.Add(ColumnDefinition())  # Bottom
        header_grid.ColumnDefinitions[4].Width = System.Windows.GridLength(80)

        # Header labels
        headers = ["Grid Name", "Left", "Right", "Top", "Bottom"]
        for i, header_text in enumerate(headers):
            header_label = TextBlock()
            header_label.Text = header_text
            header_label.FontWeight = System.Windows.FontWeights.Bold
            header_label.Margin = Thickness(5)
            header_label.HorizontalAlignment = HorizontalAlignment.Center if i != 0 else HorizontalAlignment.Left
            Grid.SetColumn(header_label, i)
            header_grid.Children.Add(header_label)

        return header_grid

    def create_data_rows(self):
        """Create data rows panel (scrollable)"""
        data_panel = StackPanel()

        # Data rows - using pre-sorted list
        for grid_data in self.sorted_grid_data_list:
            row = self.create_grid_row(grid_data)
            data_panel.Children.Add(row)

        return data_panel

    def create_grid_row(self, grid_data):
        """Create a table row for a single grid"""
        row_grid = Grid()
        row_grid.Margin = Thickness(0, 1, 0, 1)
        row_grid.Tag = grid_data  # Store reference to grid_data

        # Solid background color (no alternating)
        row_grid.Background = Brushes.White

        # Store the row reference in grid_data for later updates
        grid_data.row_grid = row_grid

        # Make row clickable for selection
        row_grid.MouseLeftButtonDown += self.row_click

        # Define columns (same widths as header, removed Select column)
        row_grid.ColumnDefinitions.Add(ColumnDefinition())
        row_grid.ColumnDefinitions[0].Width = System.Windows.GridLength(120)

        row_grid.ColumnDefinitions.Add(ColumnDefinition())
        row_grid.ColumnDefinitions[1].Width = System.Windows.GridLength(80)

        row_grid.ColumnDefinitions.Add(ColumnDefinition())
        row_grid.ColumnDefinitions[2].Width = System.Windows.GridLength(80)

        row_grid.ColumnDefinitions.Add(ColumnDefinition())
        row_grid.ColumnDefinitions[3].Width = System.Windows.GridLength(80)

        row_grid.ColumnDefinitions.Add(ColumnDefinition())
        row_grid.ColumnDefinitions[4].Width = System.Windows.GridLength(80)

        # Grid name
        name_label = TextBlock()
        name_label.Text = grid_data.name
        name_label.Margin = Thickness(5)
        name_label.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(name_label, 0)
        row_grid.Children.Add(name_label)

        # Create checkboxes for each direction
        # For horizontal grids: Left (End0) and Right (End1) are enabled
        # For vertical grids: Bottom (End0) and Top (End1) are enabled

        # Left checkbox
        check_left = CheckBox()
        check_left.HorizontalAlignment = HorizontalAlignment.Center
        check_left.VerticalAlignment = VerticalAlignment.Center
        check_left.Margin = Thickness(5)
        check_left.Tag = grid_data  # Store reference for click handler
        if not grid_data.is_vertical:
            check_left.IsEnabled = True
            check_left.IsChecked = grid_data.end1_visible  # Swapped: Left = End1
            check_left.Click += self.checkbox_left_click
        else:
            check_left.IsEnabled = False
            check_left.Opacity = 0.3
        grid_data.check_left = check_left
        Grid.SetColumn(check_left, 1)
        row_grid.Children.Add(check_left)

        # Right checkbox
        check_right = CheckBox()
        check_right.HorizontalAlignment = HorizontalAlignment.Center
        check_right.VerticalAlignment = VerticalAlignment.Center
        check_right.Margin = Thickness(5)
        check_right.Tag = grid_data  # Store reference for click handler
        if not grid_data.is_vertical:
            check_right.IsEnabled = True
            check_right.IsChecked = grid_data.end0_visible  # Swapped: Right = End0
            check_right.Click += self.checkbox_right_click
        else:
            check_right.IsEnabled = False
            check_right.Opacity = 0.3
        grid_data.check_right = check_right
        Grid.SetColumn(check_right, 2)
        row_grid.Children.Add(check_right)

        # Top checkbox
        check_top = CheckBox()
        check_top.HorizontalAlignment = HorizontalAlignment.Center
        check_top.VerticalAlignment = VerticalAlignment.Center
        check_top.Margin = Thickness(5)
        check_top.Tag = grid_data  # Store reference for click handler
        if grid_data.is_vertical:
            check_top.IsEnabled = True
            check_top.IsChecked = grid_data.end0_visible  # Swapped: Top = End0
            check_top.Click += self.checkbox_top_click
        else:
            check_top.IsEnabled = False
            check_top.Opacity = 0.3
        grid_data.check_top = check_top
        Grid.SetColumn(check_top, 3)
        row_grid.Children.Add(check_top)

        # Bottom checkbox
        check_bottom = CheckBox()
        check_bottom.HorizontalAlignment = HorizontalAlignment.Center
        check_bottom.VerticalAlignment = VerticalAlignment.Center
        check_bottom.Margin = Thickness(5)
        check_bottom.Tag = grid_data  # Store reference for click handler
        if grid_data.is_vertical:
            check_bottom.IsEnabled = True
            check_bottom.IsChecked = grid_data.end1_visible  # Swapped: Bottom = End1
            check_bottom.Click += self.checkbox_bottom_click
        else:
            check_bottom.IsEnabled = False
            check_bottom.Opacity = 0.3
        grid_data.check_bottom = check_bottom
        Grid.SetColumn(check_bottom, 4)
        row_grid.Children.Add(check_bottom)

        return row_grid

    def row_click(self, sender, args):
        """Handle row click with Shift/Ctrl support and start drag selection"""
        # Check if click originated from a checkbox - if so, don't handle row selection
        original_source = args.OriginalSource
        source = args.Source

        # Don't interfere with checkbox clicks
        if isinstance(original_source, CheckBox) or isinstance(source, CheckBox):
            return

        clicked_row = sender
        clicked_grid_data = clicked_row.Tag

        # Get the index of clicked row in the SORTED list (display order)
        clicked_index = self.sorted_grid_data_list.index(clicked_grid_data)

        # Get modifier keys
        modifiers = System.Windows.Input.Keyboard.Modifiers

        if modifiers == ModifierKeys.Control:
            # Ctrl+Click: Toggle selection of clicked row
            clicked_grid_data.is_selected = not clicked_grid_data.is_selected
            self.last_selected_index = clicked_index
            self.update_row_appearance(clicked_grid_data)
            # Start drag in toggle mode
            self.is_dragging = True
            self.drag_start_selected = clicked_grid_data.is_selected

        elif modifiers == ModifierKeys.Shift and self.last_selected_index >= 0:
            # Shift+Click: Select range from last selected to current (in sorted order)
            start_index = min(self.last_selected_index, clicked_index)
            end_index = max(self.last_selected_index, clicked_index)

            for i in range(start_index, end_index + 1):
                self.sorted_grid_data_list[i].is_selected = True
                self.update_row_appearance(self.sorted_grid_data_list[i])
            # Don't start drag for Shift+click

        else:
            # Normal click: Select only this row, deselect others
            for grid_data in self.sorted_grid_data_list:
                grid_data.is_selected = False
                self.update_row_appearance(grid_data)

            clicked_grid_data.is_selected = True
            self.last_selected_index = clicked_index
            self.update_row_appearance(clicked_grid_data)
            # Start drag in select mode
            self.is_dragging = True
            self.drag_start_selected = True

        # Mark event as handled
        args.Handled = True

    def window_mouse_move(self, sender, args):
        """Handle mouse move for drag selection"""
        if not self.is_dragging:
            return

        # Get the element under the mouse
        element = self.InputHitTest(args.GetPosition(self))
        if element is None:
            return

        # Find the row grid by traversing up the visual tree
        current = element
        while current is not None:
            if isinstance(current, Grid) and hasattr(current, 'Tag') and isinstance(current.Tag, GridBubbleData):
                # Found a row grid
                grid_data = current.Tag
                grid_data.is_selected = self.drag_start_selected
                self.update_row_appearance(grid_data)
                break
            # Move up the visual tree
            try:
                current = System.Windows.Media.VisualTreeHelper.GetParent(current)
            except:
                break

    def window_mouse_up(self, sender, args):
        """Handle mouse up to end drag selection"""
        self.is_dragging = False

    def update_row_appearance(self, grid_data):
        """Update row background based on selection state"""
        row_grid = grid_data.row_grid

        if grid_data.is_selected:
            # Selected row - highlight with blue
            row_grid.Background = SolidColorBrush(Color.FromRgb(173, 216, 230))  # Light blue
        else:
            # Not selected - solid white
            row_grid.Background = Brushes.White

    def checkbox_left_click(self, sender, args):
        """Handle left checkbox click - apply to all selected rows if clicked row is selected"""
        clicked_checkbox = sender
        clicked_grid_data = clicked_checkbox.Tag
        new_state = clicked_checkbox.IsChecked

        # If clicked row is selected, apply to all selected rows
        if clicked_grid_data.is_selected:
            for grid_data in self.grid_data_list:
                if grid_data.is_selected and grid_data.check_left.IsEnabled:
                    grid_data.check_left.IsChecked = new_state
        # Otherwise, checkbox already toggled itself (single row change)

    def checkbox_right_click(self, sender, args):
        """Handle right checkbox click - apply to all selected rows if clicked row is selected"""
        clicked_checkbox = sender
        clicked_grid_data = clicked_checkbox.Tag
        new_state = clicked_checkbox.IsChecked

        # If clicked row is selected, apply to all selected rows
        if clicked_grid_data.is_selected:
            for grid_data in self.grid_data_list:
                if grid_data.is_selected and grid_data.check_right.IsEnabled:
                    grid_data.check_right.IsChecked = new_state
        # Otherwise, checkbox already toggled itself (single row change)

    def checkbox_top_click(self, sender, args):
        """Handle top checkbox click - apply to all selected rows if clicked row is selected"""
        clicked_checkbox = sender
        clicked_grid_data = clicked_checkbox.Tag
        new_state = clicked_checkbox.IsChecked

        # If clicked row is selected, apply to all selected rows
        if clicked_grid_data.is_selected:
            for grid_data in self.grid_data_list:
                if grid_data.is_selected and grid_data.check_top.IsEnabled:
                    grid_data.check_top.IsChecked = new_state
        # Otherwise, checkbox already toggled itself (single row change)

    def checkbox_bottom_click(self, sender, args):
        """Handle bottom checkbox click - apply to all selected rows if clicked row is selected"""
        clicked_checkbox = sender
        clicked_grid_data = clicked_checkbox.Tag
        new_state = clicked_checkbox.IsChecked

        # If clicked row is selected, apply to all selected rows
        if clicked_grid_data.is_selected:
            for grid_data in self.grid_data_list:
                if grid_data.is_selected and grid_data.check_bottom.IsEnabled:
                    grid_data.check_bottom.IsChecked = new_state
        # Otherwise, checkbox already toggled itself (single row change)

    def toggle_selected_click(self, sender, args):
        """Toggle/swap bubble positions for selected rows (Left<->Right for horizontal, Top<->Bottom for vertical)"""
        for grid_data in self.grid_data_list:
            if grid_data.is_selected:
                if grid_data.is_vertical:
                    # Vertical grid: swap Top and Bottom
                    if grid_data.check_top.IsEnabled and grid_data.check_bottom.IsEnabled:
                        temp = grid_data.check_top.IsChecked
                        grid_data.check_top.IsChecked = grid_data.check_bottom.IsChecked
                        grid_data.check_bottom.IsChecked = temp
                else:
                    # Horizontal grid: swap Left and Right
                    if grid_data.check_left.IsEnabled and grid_data.check_right.IsEnabled:
                        temp = grid_data.check_left.IsChecked
                        grid_data.check_left.IsChecked = grid_data.check_right.IsChecked
                        grid_data.check_right.IsChecked = temp

    def uncheck_all_click(self, sender, args):
        """Uncheck all checkboxes to reset everything"""
        for grid_data in self.grid_data_list:
            if grid_data.check_left.IsEnabled:
                grid_data.check_left.IsChecked = False
            if grid_data.check_right.IsEnabled:
                grid_data.check_right.IsChecked = False
            if grid_data.check_top.IsEnabled:
                grid_data.check_top.IsChecked = False
            if grid_data.check_bottom.IsEnabled:
                grid_data.check_bottom.IsChecked = False

    def apply_click(self, sender, args):
        """Apply changes but keep window open"""
        try:
            apply_bubble_changes(self.grid_data_list, self.view)
            # Visual feedback that changes were applied
            original_content = sender.Content
            sender.Content = "Applied ✓"
            sender.IsEnabled = False

            # Create and start timer to reset button
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromSeconds(1)

            def on_timer_tick(s, e):
                sender.Content = original_content
                sender.IsEnabled = True
                timer.Stop()

            timer.Tick += on_timer_tick
            timer.Start()

        except Exception as e:
            forms.alert("Error applying changes: {}".format(str(e)), title="Error")

    def ok_click(self, sender, args):
        """Close window"""
        self.DialogResult = True
        self.Close()

    def apply_to_views_click(self, sender, args):
        """Apply current settings to other views"""
        # Get all applicable views in the project
        all_views = DB.FilteredElementCollector(doc)\
                      .OfClass(DB.View)\
                      .WhereElementIsNotElementType()\
                      .ToElements()

        # Filter to valid view types
        valid_view_types = [
            DB.ViewType.FloorPlan,
            DB.ViewType.CeilingPlan,
            DB.ViewType.EngineeringPlan,
            DB.ViewType.AreaPlan,
            DB.ViewType.Elevation,
            DB.ViewType.Section
        ]

        applicable_views = []
        for v in all_views:
            if v.ViewType in valid_view_types and not v.IsTemplate:
                # Exclude the current view
                if v.Id != self.view.Id:
                    applicable_views.append(v)

        if not applicable_views:
            forms.alert("No other applicable views found.", title="No Views")
            return

        # Create view selection dict
        view_dict = {"{} - {}".format(v.ViewType, v.Name): v for v in applicable_views}

        # Show custom view selection window with filter
        view_selection_window = ViewSelectionWindow(view_dict)
        result = view_selection_window.ShowDialog()

        # If user cancels or doesn't select anything, exit without applying
        if not result or not view_selection_window.selected_views:
            return

        # Apply settings to selected views
        selected_views = view_selection_window.selected_views

        try:
            # Apply the current checkbox states to selected views
            for target_view in selected_views:
                apply_bubble_changes(self.grid_data_list, target_view)

            # Update button text temporarily to show success (no popup)
            original_content = sender.Content
            sender.Content = "Applied to {} views ✓".format(len(selected_views))
            sender.IsEnabled = False

            # Reset button after delay
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromSeconds(2)

            def on_timer_tick(s, e):
                sender.Content = original_content
                sender.IsEnabled = True
                timer.Stop()

            timer.Tick += on_timer_tick
            timer.Start()

        except Exception as e:
            forms.alert("Error applying to views: {}".format(str(e)), title="Error")

    def cancel_click(self, sender, args):
        """Cancel without applying"""
        self.DialogResult = False
        self.Close()


def get_grids_in_view(view):
    """Get all grid elements visible in the view"""
    collector = DB.FilteredElementCollector(doc, view.Id)\
                  .OfClass(DB.Grid)\
                  .WhereElementIsNotElementType()
    return list(collector)


def apply_bubble_changes(grid_data_list, view):
    """Apply bubble visibility changes to grids"""
    with revit.Transaction("Update Grid Bubbles"):
        for grid_data in grid_data_list:
            grid = grid_data.grid

            # Determine which checkboxes to use based on orientation
            if grid_data.is_vertical:
                # Vertical grid: Top = End0, Bottom = End1 (fixed mapping)
                end0_checked = grid_data.check_top.IsChecked
                end1_checked = grid_data.check_bottom.IsChecked
            else:
                # Horizontal grid: Left = End1, Right = End0 (swapped mapping)
                end0_checked = grid_data.check_right.IsChecked
                end1_checked = grid_data.check_left.IsChecked

            # Update End 0
            try:
                if end0_checked:
                    grid.ShowBubbleInView(DB.DatumEnds.End0, view)
                else:
                    grid.HideBubbleInView(DB.DatumEnds.End0, view)
            except Exception:
                pass  # Silently ignore errors (e.g., grid not visible in view)

            # Update End 1
            try:
                if end1_checked:
                    grid.ShowBubbleInView(DB.DatumEnds.End1, view)
                else:
                    grid.HideBubbleInView(DB.DatumEnds.End1, view)
            except Exception:
                pass  # Silently ignore errors (e.g., grid not visible in view)


def main():
    """Main execution function"""
    # Check if we're in a valid view
    if not active_view:
        forms.alert("No active view found.", exitscript=True)

    # Check if view type supports grids
    view_type = active_view.ViewType
    valid_view_types = [
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.EngineeringPlan,
        DB.ViewType.AreaPlan,
        DB.ViewType.Elevation,
        DB.ViewType.Section
    ]

    if view_type not in valid_view_types:
        forms.alert("This tool only works in floor plans, ceiling plans, elevations, and sections.",
                   exitscript=True)

    # Get all grids in the view
    grids = get_grids_in_view(active_view)

    if not grids:
        forms.alert("No grids found in the active view.", exitscript=True)

    # Create grid data objects
    grid_data_list = []
    for grid in grids:
        try:
            grid_data = GridBubbleData(grid, active_view)
            grid_data_list.append(grid_data)
        except Exception as e:
            print("Error processing grid {}: {}".format(grid.Name, str(e)))

    if not grid_data_list:
        forms.alert("Could not process any grids in the view.", exitscript=True)

    # Show window
    window = GridBubbleWindow(grid_data_list, active_view)
    window.ShowDialog()


if __name__ == "__main__":
    main()
