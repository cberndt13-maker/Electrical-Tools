#pylint: disable=E0401,C0103,C0111
"""Helper module to set the default publish set for a ViewSheetSet."""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSheetSet,
    PrintRange
)


def get_viewsheetset_by_name(doc, name):
    """Get a ViewSheetSet by its name."""
    collector = FilteredElementCollector(doc).OfClass(ViewSheetSet)
    for vss in collector:
        if vss.Name == name:
            return vss
    return None


def set_default_publish_set(doc, viewsheetset):
    """Set the given ViewSheetSet as the default publish set.

    Args:
        doc: The Revit document
        viewsheetset: The ViewSheetSet to set as default

    Returns:
        bool: True if successful, False otherwise
    """
    if viewsheetset is None:
        return False

    try:
        print_manager = doc.PrintManager
        print_manager.PrintRange = PrintRange.Select

        # Get the ViewSheetSetting which controls the publish settings
        view_sheet_setting = print_manager.ViewSheetSetting

        # Set the current view sheet set
        view_sheet_setting.CurrentViewSheetSet = viewsheetset

        return True
    except Exception as e:
        print("Error setting default publish set: {}".format(str(e)))
        return False


def set_default_publish_set_by_name(doc, set_name):
    """Set the default publish set by name.

    Args:
        doc: The Revit document
        set_name: The name of the ViewSheetSet to set as default

    Returns:
        bool: True if successful, False otherwise
    """
    viewsheetset = get_viewsheetset_by_name(doc, set_name)
    if viewsheetset is None:
        print("Could not find ViewSheetSet with name: {}".format(set_name))
        return False

    return set_default_publish_set(doc, viewsheetset)
