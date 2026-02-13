"""Reset the saved linked file selection for Center Lights tool"""
__title__ = "Reset Link\nSelection"
__author__ = "Christopher Berndt"
__doc__ = "Clears the saved linked file selection. Next time you run Center Lights, you'll be prompted to select a link."

from pyrevit import script, forms

# Clear the saved link selection and all grid offsets
# Use the same shared config section as the main tool
config = script.get_config(section='LightCenterTool')

print("=" * 60)
print("RESET: Starting calibration reset...")
print("=" * 60)

# Clear link selection
old_link_name = getattr(config, 'selected_link_name', None)
config.selected_link_name = None
print("Cleared link selection: {}".format(old_link_name if old_link_name else "(none)"))

# Clear all grid offset and spacing configurations
cleared_count = 0
cleared_items = []

for attr_name in dir(config):
    if attr_name.startswith('grid_offset_') or attr_name.startswith('grid_spacing_'):
        try:
            delattr(config, attr_name)
            cleared_items.append(attr_name)
            cleared_count += 1
        except Exception as e:
            print("ERROR deleting {}: {}".format(attr_name, str(e)))

print("\nCleared {} calibration parameter(s):".format(cleared_count))
for item in sorted(cleared_items):
    print("  - {}".format(item))

script.save_config()
print("\nConfiguration saved.")
print("=" * 60)

forms.alert(
    "Link selection and grid calibration have been reset.\n\n"
    "Next time you run 'Center Lights In Grid', you will be prompted to:\n"
    "1. Select a linked file\n"
    "2. Calibrate the grid by clicking a reference point"
)
