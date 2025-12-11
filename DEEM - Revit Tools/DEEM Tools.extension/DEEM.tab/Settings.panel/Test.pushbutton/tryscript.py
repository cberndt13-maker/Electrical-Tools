# -*- coding: utf-8 -*-
__title__ = "Check\nxlsxwriter"
__author__ = "Christopher + GPT-5"
__doc__ = "Checks if the xlsxwriter module is installed and available."

from pyrevit import script

output = script.get_output()

try:
    import xlsxwriter
    output.print_md("✅ **xlsxwriter is installed and available!**")
except ImportError:
    output.print_md("❌ **xlsxwriter is NOT installed.**\n\nYou can install it via:\n```\npyrevit env\npip install xlsxwriter\n```")
