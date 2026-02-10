# -*- coding: utf-8 -*-
from pyrevit.loader import sessionmgr
from pyrevit import forms

# Reloads all pyRevit extensions
try:
    sessionmgr.reload_pyrevit()
    forms.alert('pyRevit reloaded successfully!', title='Reload Complete')
except Exception as e:
    forms.alert('Failed to reload pyRevit.\n\nError: {}'.format(str(e)), title='Reload Failed')
