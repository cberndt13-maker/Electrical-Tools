# Hides pyRevit Tab
 
#Author : Chris Berndt
 
#Date : 10/16/25

import clr
clr.AddReference('AdWindows')
import Autodesk.Windows as adWin

# Get the Revit ribbon
ribbon = adWin.ComponentManager.Ribbon

# Try to find the pyRevit tab
pyrevit_tab = ribbon.FindTab('pyRevit')

if pyrevit_tab:
    # Toggle visibility
    if pyrevit_tab.IsVisible:
        pyrevit_tab.IsVisible = False
        print('pyRevit tab hidden.')
    else:
        pyrevit_tab.IsVisible = True
        print('pyRevit tab shown.')
else:
    print('pyRevit tab not found.')
