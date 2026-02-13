"""Center Light Fixtures in DB.Ceiling Grid"""
__title__ = "Center Lights\nIn Grid"
__author__ = "Christopher Berndt"
__doc__ = "Centers light fixtures in ACT ceiling grid from linked architectural model."

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ObjectType
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows import Window
from System.Windows.Controls import Button, StackPanel, TextBlock, Border, ScrollViewer, Grid as WPFGrid
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment
from System.Windows.Media import Brushes, SolidColorBrush, Color
from System.Windows.Controls import RowDefinition

# Get current document and UI document
doc = revit.doc
uidoc = revit.uidoc

# Global grid offset dictionary - keyed by ceiling ID
# Value: (offset_x, offset_y, spacing_x, spacing_y)
CEILING_GRID_OFFSETS = {}


class LinkSelectionWindow(Window):
    """WPF Window for selecting a linked file"""

    def __init__(self, link_dict):
        self.link_dict = link_dict
        self.selected_link = None
        self.Title = "Select Linked File with Ceiling Grid"
        self.Width = 500
        self.Height = 400
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Topmost = True  # Keep window on top

        # Main layout
        main_grid = WPFGrid()
        main_grid.Margin = Thickness(15)

        # Define rows
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[0].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[1].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[2].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)

        # Instructions
        instructions = TextBlock()
        instructions.Text = "Select the linked file containing the ACT ceiling grid:"
        instructions.Margin = Thickness(0, 0, 0, 10)
        instructions.FontSize = 12
        instructions.FontWeight = System.Windows.FontWeights.Bold
        WPFGrid.SetRow(instructions, 0)
        main_grid.Children.Add(instructions)

        # Separator line
        separator = Border()
        separator.BorderThickness = Thickness(0, 0, 0, 1)
        separator.BorderBrush = SolidColorBrush(Color.FromRgb(200, 200, 200))
        separator.Margin = Thickness(0, 0, 0, 10)
        WPFGrid.SetRow(separator, 0)
        WPFGrid.SetColumn(separator, 0)
        main_grid.Children.Add(separator)

        # Scrollable list of links
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto

        self.link_panel = StackPanel()
        scroll.Content = self.link_panel

        WPFGrid.SetRow(scroll, 1)
        main_grid.Children.Add(scroll)

        # Create link rows
        for link_name in sorted(link_dict.keys()):
            row = self.create_link_row(link_name)
            self.link_panel.Children.Add(row)

        # Buttons
        button_panel = StackPanel()
        button_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Right
        button_panel.Margin = Thickness(0, 10, 0, 0)

        select_btn = Button()
        select_btn.Content = "Select"
        select_btn.Width = 80
        select_btn.Margin = Thickness(0, 0, 5, 0)
        select_btn.Click += self.select_click
        button_panel.Children.Add(select_btn)

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Click += self.cancel_click
        button_panel.Children.Add(cancel_btn)

        WPFGrid.SetRow(button_panel, 2)
        main_grid.Children.Add(button_panel)

        self.Content = main_grid

    def create_link_row(self, link_name):
        """Create a selectable row for a link"""
        row_border = Border()
        row_border.BorderThickness = Thickness(0)
        row_border.Padding = Thickness(5)
        row_border.Margin = Thickness(0, 1, 0, 1)
        row_border.Background = Brushes.White
        row_border.Tag = link_name
        row_border.MouseLeftButtonDown += self.row_click

        # Link name text
        text_block = TextBlock()
        text_block.Text = link_name
        text_block.VerticalAlignment = VerticalAlignment.Center
        row_border.Child = text_block

        return row_border

    def row_click(self, sender, args):
        """Handle row click"""
        # Deselect all rows
        for child in self.link_panel.Children:
            if isinstance(child, Border):
                child.Background = Brushes.White

        # Select clicked row
        sender.Background = SolidColorBrush(Color.FromRgb(173, 216, 230))  # Light blue
        self.selected_link = self.link_dict[sender.Tag]

    def select_click(self, sender, args):
        """Apply selection"""
        if self.selected_link:
            self.DialogResult = True
            self.Close()
        else:
            forms.alert("Please select a linked file.")

    def cancel_click(self, sender, args):
        """Cancel selection"""
        self.selected_link = None
        self.DialogResult = False
        self.Close()


def is_act_ceiling(ceiling):
    """Check if ceiling is an ACT (Acoustic Ceiling Tile) with grid"""
    try:
        ceiling_type_name = ceiling.CeilingType.FamilyName.lower()
        type_name = ceiling.CeilingType.Name.lower()

        # Check for keywords that indicate ACT/grid ceilings
        act_keywords = ['act', 'acoustic', 'grid', 'tile', '2x2', '2x4', '2 x 2', '2 x 4']

        combined_name = ceiling_type_name + " " + type_name

        for keyword in act_keywords:
            if keyword in combined_name:
                return True
    except:
        pass

    return False


def get_ceiling_at_point(point, link_instance):
    """Get ceiling element at a given point from the specified linked file
    Returns: (ceiling, transform) tuple"""

    try:
        link_doc = link_instance.GetLinkDocument()
        if not link_doc:
            return None, None

        # Get transform from link instance
        transform = link_instance.GetTransform()

        # Transform the point to link coordinates
        inverse_transform = transform.Inverse
        link_point = inverse_transform.OfPoint(point)

        # Get ceilings from linked document
        link_ceilings = DB.FilteredElementCollector(link_doc).OfClass(DB.Ceiling).ToElements()

        for ceiling in link_ceilings:
            try:
                bbox = ceiling.get_BoundingBox(None)
                if not bbox:
                    continue

                if (bbox.Min.X <= link_point.X <= bbox.Max.X and
                    bbox.Min.Y <= link_point.Y <= bbox.Max.Y):
                    return ceiling, transform
            except:
                continue
    except:
        pass

    return None, None


def get_grid_spacing(ceiling, use_default=True):
    """Extract grid spacing from ceiling element or use default
    Returns (spacing_x, spacing_y) in feet (Revit internal units)"""

    # Try to get grid spacing from ceiling parameters
    all_params = ceiling.Parameters
    spacing_x = None
    spacing_y = None

    # Look for parameters that might contain grid spacing
    for param in all_params:
        param_name = param.Definition.Name.lower()
        if "grid" in param_name:
            if "spacing" in param_name or "size" in param_name:
                try:
                    value = param.AsDouble()
                    if value > 0:
                        if spacing_x is None:
                            spacing_x = value
                        else:
                            spacing_y = value
                except:
                    pass

    # Default to 2' x 2' ACT ceiling grid (standard size)
    if use_default:
        if spacing_x is None:
            spacing_x = 2.0
        if spacing_y is None:
            spacing_y = 2.0

    if spacing_x is None or spacing_y is None:
        return None, None

    return spacing_x, spacing_y


def get_closest_grid_intersection(fixture_location, ceiling, transform, ceiling_id):
    """Find the closest grid line intersection for fixture location
    transform: Transform from link instance if ceiling is linked, None if in host
    ceiling_id: ID of the ceiling element for offset lookup"""

    # Check if we have an offset for this ceiling
    if ceiling_id not in CEILING_GRID_OFFSETS:
        return None  # No offset calibrated for this ceiling yet

    # Get stored offset and spacing
    offset_x, offset_y, spacing_x, spacing_y = CEILING_GRID_OFFSETS[ceiling_id]

    if not spacing_x or not spacing_y:
        return None

    # Find bounding box of ceiling
    bbox = ceiling.get_BoundingBox(None)

    if not bbox:
        return None

    # If ceiling is from a link, work in link coordinates then transform back
    work_location = fixture_location
    if transform:
        inverse_transform = transform.Inverse
        work_location = inverse_transform.OfPoint(fixture_location)

    # Calculate position relative to the grid pattern
    # The grid line intersections are at positions where (X % spacing) = offset_x
    # Cell centers are at intersection + (spacing/2, spacing/2)

    # Find which cell the fixture is in, accounting for the offset
    adjusted_x = work_location.X - offset_x
    adjusted_y = work_location.Y - offset_y

    # Use floor to find which cell the fixture is currently in
    # This prevents fixtures near cell boundaries from jumping to adjacent cells
    import math
    cell_index_x = int(math.floor(adjusted_x / spacing_x))
    cell_index_y = int(math.floor(adjusted_y / spacing_y))

    # Calculate the corner intersection position for this cell
    intersection_x = cell_index_x * spacing_x + offset_x
    intersection_y = cell_index_y * spacing_y + offset_y

    # Now offset by half spacing to get to the CENTER of the cell
    center_x = intersection_x + (spacing_x / 2.0)
    center_y = intersection_y + (spacing_y / 2.0)
    center_z = work_location.Z

    center_point = DB.XYZ(center_x, center_y, center_z)

    # Transform back to host coordinates if needed
    if transform:
        center_point = transform.OfPoint(center_point)

    return center_point


def check_hosted_fixtures(fixtures):
    """Check if any fixtures are ceiling-hosted and return list of hosted fixtures with details
    Returns: list of (fixture, family_name, type_name) tuples"""
    hosted_fixtures = []

    for fixture in fixtures:
        try:
            # Check if fixture has a host (ceiling-hosted)
            if hasattr(fixture, 'Host') and fixture.Host is not None:
                family_name = fixture.Symbol.FamilyName
                type_name = fixture.Symbol.Name
                hosted_fixtures.append((fixture, family_name, type_name))
        except:
            pass

    return hosted_fixtures


def center_fixture_in_grid(fixture, link_instance):
    """Center a lighting fixture at ceiling grid line intersection
    Returns: (success, ceiling_id) tuple"""

    # Get fixture location
    location = fixture.Location

    if not isinstance(location, DB.LocationPoint):
        return False, None

    # Use the fixture's bounding box center for X/Y, but keep Z at the insertion point level
    # This handles recessed fixtures where the bounding box includes housing above the ceiling
    try:
        bbox = fixture.get_BoundingBox(None)
        if bbox:
            # Calculate bounding box center for X and Y, but use insertion point Z
            # This ensures we're working at the visible trim level, not the housing level
            fixture_point = DB.XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                location.Point.Z  # Use insertion point Z (at ceiling level)
            )
        else:
            # Fall back to insertion point if no bounding box
            fixture_point = location.Point
    except:
        # Fall back to insertion point if error
        fixture_point = location.Point

    # Find ceiling at fixture location
    ceiling, transform = get_ceiling_at_point(fixture_point, link_instance)

    if not ceiling:
        return False, None

    # Get ceiling ID
    ceiling_id = ceiling.Id.IntegerValue

    # Get closest grid intersection
    target_point = get_closest_grid_intersection(fixture_point, ceiling, transform, ceiling_id)

    if not target_point:
        return False, ceiling_id  # Return ceiling_id so caller knows which ceiling needs calibration

    # Move fixture to grid intersection
    # Note: We calculated target based on bounding box center, but we move the insertion point
    # So we need to calculate the translation from the current center to the target
    try:
        translation = target_point - fixture_point
        location.Move(translation)
    except Exception as move_error:
        print("ERROR: Failed to move fixture - {}".format(str(move_error)))
        return False, ceiling_id

    return True, ceiling_id


# Main execution
t = None
try:
    # Get current view
    active_view = doc.ActiveView

    if not isinstance(active_view, DB.ViewPlan):
        forms.alert("Please run this tool in a ceiling plan view.")
        import sys
        sys.exit()

    # Prompt user to select the linked file with ceiling grid
    link_collector = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()

    if not link_collector:
        forms.alert("No linked Revit files found in this project.")
        import sys
        sys.exit()

    # Create a selection dialog for link files
    link_dict = {}

    for link in link_collector:
        if link.GetLinkDocument():
            link_dict[link.Name] = link

    if not link_dict:
        forms.alert("No loaded linked Revit files found.")
        import sys
        sys.exit()

    # Try to retrieve the previously selected link from config
    # Use a shared config section so reset button can access it
    config = script.get_config(section='LightCenterTool')
    saved_link_name = getattr(config, 'selected_link_name', None)
    selected_link = None

    # Check if the saved link still exists and is loaded
    if saved_link_name and saved_link_name in link_dict:
        # Use the saved link
        selected_link = link_dict[saved_link_name]
    else:
        # No saved link or it's been reset - show selection window
        link_window = LinkSelectionWindow(link_dict)
        result = link_window.ShowDialog()

        if not result or not link_window.selected_link:
            forms.alert("No link selected. Operation cancelled.")
            import sys
            sys.exit()

        selected_link = link_window.selected_link
        # Save the selected link name for next time
        config.selected_link_name = selected_link.Name
        script.save_config()

    # Load all saved ceiling offsets and spacings from config
    for attr_name in dir(config):
        if attr_name.startswith('grid_offset_ceiling_{}_{}'.format(selected_link.Name, '')):
            try:
                # Parse ceiling ID from attribute name
                parts = attr_name.split('_')
                if len(parts) >= 5:
                    ceiling_id = int(parts[4])
                    offset_x = getattr(config, 'grid_offset_ceiling_{}_{}_{}'.format(selected_link.Name, ceiling_id, 'x'), None)
                    offset_y = getattr(config, 'grid_offset_ceiling_{}_{}_{}'.format(selected_link.Name, ceiling_id, 'y'), None)
                    spacing_x = getattr(config, 'grid_spacing_ceiling_{}_{}_{}'.format(selected_link.Name, ceiling_id, 'x'), None)
                    spacing_y = getattr(config, 'grid_spacing_ceiling_{}_{}_{}'.format(selected_link.Name, ceiling_id, 'y'), None)
                    if all(v is not None for v in [offset_x, offset_y, spacing_x, spacing_y]):
                        CEILING_GRID_OFFSETS[ceiling_id] = (offset_x, offset_y, spacing_x, spacing_y)
            except:
                pass

    # Get selected elements or prompt for selection
    selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

    if not selection:
        # Prompt user to select fixtures
        selected_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            "Select light fixtures to center in ceiling grid"
        )

        selection = [doc.GetElement(ref.ElementId) for ref in selected_refs]

    if not selection:
        forms.alert("No elements selected.")
        import sys
        sys.exit()

    # Filter for lighting fixtures
    lighting_fixtures = [elem for elem in selection
                        if elem.Category and
                        elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_LightingFixtures)]

    if not lighting_fixtures:
        forms.alert("No lighting fixtures selected.")
        import sys
        sys.exit()

    # Check for ceiling-hosted fixtures
    hosted_fixtures = check_hosted_fixtures(lighting_fixtures)

    if hosted_fixtures:
        # Build message showing hosted fixtures
        message = "WARNING: Ceiling-hosted fixtures detected!\n\n"
        message += "The following fixtures are ceiling-hosted and may cause issues:\n"
        message += "-" * 60 + "\n\n"

        # Group by family and type
        fixture_types = {}
        for fixture, family_name, type_name in hosted_fixtures:
            key = (family_name, type_name)
            if key not in fixture_types:
                fixture_types[key] = []
            fixture_types[key].append(fixture)

        # Display grouped fixtures
        for (family_name, type_name), fixtures_list in sorted(fixture_types.items()):
            message += "Family: {}\n".format(family_name)
            message += "Type: {}\n".format(type_name)
            message += "Count: {} fixture(s)\n".format(len(fixtures_list))
            message += "\n"

        message += "-" * 60 + "\n\n"
        message += "Ceiling-hosted fixtures may fail to move if the new position\n"
        message += "is outside their host ceiling's boundary.\n\n"
        message += "RECOMMENDATION: Change these fixtures to work plane-based\n"
        message += "before running this tool.\n\n"
        message += "Do you want to continue anyway?"

        # Ask user if they want to continue
        continue_anyway = forms.alert(
            message,
            title="Ceiling-Hosted Fixtures Warning",
            ok=False,
            yes=True,
            no=True
        )

        if not continue_anyway:
            import sys
            sys.exit()

    # Find all unique ceilings that need calibration
    uncalibrated_ceilings = set()
    for fixture in lighting_fixtures:
        location = fixture.Location
        if isinstance(location, DB.LocationPoint):
            fixture_point = location.Point
            ceiling, transform = get_ceiling_at_point(fixture_point, selected_link)
            if ceiling:
                ceiling_id = ceiling.Id.IntegerValue
                if ceiling_id not in CEILING_GRID_OFFSETS:
                    uncalibrated_ceilings.add(ceiling_id)

    # Calibrate any uncalibrated ceilings
    if uncalibrated_ceilings:
        forms.alert(
            "Grid calibration needed:\n\n"
            "You have fixtures on {} ceiling(s) that haven't been calibrated yet.\n\n"
            "For each ceiling, you will click on TWO GRID LINE INTERSECTIONS.\n\n"
            "IMPORTANT: Zoom in close and click EXACTLY where grid lines cross!\n"
            "Click on two intersections that are diagonally opposite (across one grid cell).".format(len(uncalibrated_ceilings))
        )

        # Get the transform for coordinate conversion
        link_transform = selected_link.GetTransform()
        inverse_transform = link_transform.Inverse

        for ceiling_id in uncalibrated_ceilings:
            try:
                # Let user pick two grid line intersections in HOST coordinates
                intersection1_host = uidoc.Selection.PickPoint("Ceiling {}: Click EXACTLY on a GRID LINE INTERSECTION (zoom in!)".format(ceiling_id))
                intersection2_host = uidoc.Selection.PickPoint("Ceiling {}: Click EXACTLY on ANOTHER INTERSECTION (diagonally opposite)".format(ceiling_id))

                # Transform to LINK coordinates since the ceiling grid is in the linked file
                intersection1_link = inverse_transform.OfPoint(intersection1_host)
                intersection2_link = inverse_transform.OfPoint(intersection2_host)

                # Calculate spacing from the two intersection points IN LINK COORDINATES
                spacing_x = abs(intersection2_link.X - intersection1_link.X)
                spacing_y = abs(intersection2_link.Y - intersection1_link.Y)

                # Validate spacing - should be reasonable values (not too small)
                if spacing_x < 0.5 or spacing_y < 0.5:
                    forms.alert("ERROR: Grid spacing too small ({:.2f}' x {:.2f}'). \nMake sure you clicked on TWO DIFFERENT intersections that are diagonally opposite.\nSkipping ceiling {}.".format(spacing_x, spacing_y, ceiling_id))
                    continue

                # Use intersection1 IN LINK COORDINATES as reference point for offset calculation
                # The offset is where grid intersections fall relative to the coordinate system
                grid_offset_x = intersection1_link.X % spacing_x
                grid_offset_y = intersection1_link.Y % spacing_y

                # Save for this ceiling (offset + spacing, ALL IN LINK COORDINATES)
                CEILING_GRID_OFFSETS[ceiling_id] = (grid_offset_x, grid_offset_y, spacing_x, spacing_y)

                # Save to config
                offset_x_key = 'grid_offset_ceiling_{}_{}_{}'.format(selected_link.Name, ceiling_id, 'x')
                offset_y_key = 'grid_offset_ceiling_{}_{}_{}'.format(selected_link.Name, ceiling_id, 'y')
                spacing_x_key = 'grid_spacing_ceiling_{}_{}_{}'.format(selected_link.Name, ceiling_id, 'x')
                spacing_y_key = 'grid_spacing_ceiling_{}_{}_{}'.format(selected_link.Name, ceiling_id, 'y')

                setattr(config, offset_x_key, grid_offset_x)
                setattr(config, offset_y_key, grid_offset_y)
                setattr(config, spacing_x_key, spacing_x)
                setattr(config, spacing_y_key, spacing_y)
                script.save_config()
            except Exception as calib_error:
                forms.alert("Calibration cancelled. Skipping ceiling {}.".format(ceiling_id))
                continue

    # Start transaction
    t = DB.Transaction(doc, "Center Lights in Grid")
    t.Start()

    success_count = 0
    fail_count = 0

    for fixture in lighting_fixtures:
        success, ceiling_id = center_fixture_in_grid(fixture, selected_link)
        if success:
            success_count += 1
        else:
            fail_count += 1

    t.Commit()

    # Report results
    message = "Centered {} fixture(s) successfully.".format(success_count)
    if fail_count > 0:
        message += "\n{} fixture(s) could not be centered (no ceiling grid found or not calibrated).".format(fail_count)

    forms.alert(message)

except Exception as e:
    # Rollback transaction if it was started
    if t and t.HasStarted():
        t.RollBack()
    forms.alert("An error occurred:\n{}".format(str(e)))