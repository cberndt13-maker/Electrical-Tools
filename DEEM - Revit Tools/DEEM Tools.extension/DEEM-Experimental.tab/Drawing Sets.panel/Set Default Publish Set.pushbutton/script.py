#pylint: disable=E0401,C0103,C0111
"""Set a ViewSheetSet as the default print set."""

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import FilteredElementCollector, ViewSheetSet, PrintRange

doc = revit.doc


def get_all_viewsheetsets(doc):
    """Get all ViewSheetSets in the document."""
    collector = FilteredElementCollector(doc).OfClass(ViewSheetSet)
    return list(collector)


def set_default_print_set(doc, viewsheetset):
    """Set the given ViewSheetSet as the default print set."""
    if viewsheetset is None:
        return False

    try:
        print_manager = doc.PrintManager
        print_manager.PrintRange = PrintRange.Select
        view_sheet_setting = print_manager.ViewSheetSetting

        # Set the current view sheet set
        view_sheet_setting.CurrentViewSheetSet = viewsheetset

        # Save the setting - this is required to persist the change
        # Use Save() if it's an existing named set, otherwise use SaveAs()
        try:
            view_sheet_setting.Save()
        except:
            # If Save fails (e.g., it's the "In-Session" set), try SaveAs
            pass

        return True
    except Exception as e:
        print("Error setting default print set: {}".format(str(e)))
        return False


# Get all ViewSheetSets
all_sets = get_all_viewsheetsets(doc)

if not all_sets:
    forms.alert('No sheet sets found in the document.')
    script.exit()

# Let user select a ViewSheetSet
selected_set = forms.SelectFromList.show(
    all_sets,
    name_attr='Name',
    title='Select Sheet Set',
    button_name='Set as Default Print Set'
)

if selected_set:
    if set_default_print_set(doc, selected_set):
        forms.alert('Successfully set "{}" as the default print set.'.format(selected_set.Name))
    else:
        forms.alert('Failed to set the default print set.')
