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

# Get light fixture types from elements
def get_fixture_types_dict(fixtures):
    type_dict = {}
    for fixture in fixtures:
        type_id = fixture.GetTypeId()
        if type_id != DB.ElementId.InvalidElementId:
            fixture_type = fixture.Document.GetElement(type_id)
            if fixture_type:
                type_name = fixture_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
                if type_name not in type_dict:
                    type_dict[type_name] = []
                type_dict[type_name].append(fixture)
    return type_dict

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
        selection_method = forms.CommandSwitchWindow.show(
            ['Select by Type', 'Pick in View'],
            message='How would you like to select light fixtures?'
        )

        if selection_method == 'Select by Type':
            # Get all light fixtures in the model
            all_fixtures = get_all_light_fixtures(doc)

            if not all_fixtures:
                forms.alert("No light fixtures found in the model.", title="Light Fixture Flip")
            else:
                # Get fixture types dictionary
                type_dict = get_fixture_types_dict(all_fixtures)

                if not type_dict:
                    forms.alert("No valid light fixture types found.", title="Light Fixture Flip")
                else:
                    # Let user select which types to flip
                    selected_types = forms.SelectFromList.show(
                        sorted(type_dict.keys()),
                        title='Select Light Fixture Types to Flip Work Plane',
                        multiselect=True,
                        button_name='Flip Selected Types'
                    )

                    if selected_types:
                        # Collect all fixtures of selected types
                        for type_name in selected_types:
                            fixtures_to_flip.extend(type_dict[type_name])

        elif selection_method == 'Pick in View':
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
