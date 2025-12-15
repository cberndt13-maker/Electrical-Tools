#-*- coding: utf-8 -*-



# ███████╗███████╗██████╗  ██████╗ ██╗   ██╗███████╗ ██████╗ ███╗   ██╗

# ██╔════╝██╔════╝██╔══██╗██╔════╝ ██║   ██║██╔════╝██╔═══██╗████╗  ██║

# █████╗  █████╗  ██████╔╝██║  ███╗██║   ██║███████╗██║   ██║██╔██╗ ██║

# ██╔══╝  ██╔══╝  ██╔══██╗██║   ██║██║   ██║╚════██║██║   ██║██║╚██╗██║

# ██║     ███████╗██║  ██║╚██████╔╝╚██████╔╝███████║╚██████╔╝██║ ╚████║

# ╚═╝     ╚══════╝╚═╝  ╚═╝ ╚═════╝  ╚═════╝ ╚══════╝ ╚═════╝ ╚═╝  ╚═══╝



#######################



# Title : View Templates



pyRevit_tool_name = "View Templates" # Name used by tracking tool

#   __commandname__



# Description : This tool will create new view templates based on filters.



# Author : Chris Berndt



# Date : 10-16-25



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



# itertools Library

from itertools import groupby



# DEEM Library

# from DEEM import general, error_handling, notifications, select, setup, tracking, units, variables



# Error info

import traceback



# Getting the path of the current command

direct = __commandpath__



py_split = direct.split("pyRevit")



# This is going to be the path of files used in scripts

ref_path = py_split[0] + "pyRevit\DEEM - Revit Tools"



#######################



#    ______ _    _ _   _  _____ _______ _____ ____  _   _  _____ 

#   |  ____| |  | | \ | |/ ____|__   __|_   _/ __ \| \ | |/ ____|

#   | |__  | |  | |  \| | |       | |    | || |  | |  \| | (___  

#   |  __| | |  | | . ` | |       | |    | || |  | | . ` |\___ \ 

#   | |    | |__| | |\  | |____   | |   _| || |__| | |\  |____) |

#   |_|     \____/|_| \_|\_____|  |_|  |_____\____/|_| \_|_____/ 

#                                                                



#######################



# Final option to notify user of error logging. Not important, but want to have something to try to notify of issues so they can be resolved and the data collected.

def final_dialog(tool):

    dialog = UI.TaskDialog("Error - Logging Tracking")

    dialog.MainIcon = UI.TaskDialogIcon.TaskDialogIconError

    dialog.MainInstruction = "Please notify your manager/tool creator of this message."

    dialog.MainContent = "Check details below."

    dialog.ExpandedContent = "There was an issue logging the tracking data when using the " + str(tool) + " tool."

    #dialog.FooterText = ""

    dialog.Show()



# Popup to show errors

def show_error_dialog(title="Error", message="", details=""):

    dialog = UI.TaskDialog(title)

    dialog.MainIcon = UI.TaskDialogIcon.TaskDialogIconError

    dialog.MainInstruction = message

    dialog.MainContent = "Check details below."

    dialog.ExpandedContent = details

    #dialog.FooterText = ""

    dialog.Show()



# Script beleow needs to be wrapped in this code

## This should catch errors/exceptions with the script

def safe_run(func):

    def wrapper(*args, **kwargs):

        try:

            result = func(*args, **kwargs)

            return result, None

        except Exception as ex:

            details = traceback.format_exc()

            show_error_dialog(title="Hanger Tool Error", message="An unexpected error occurred.", details=details)

            return None, details

    return wrapper



def duplicate_view_template(temp, view_name):

    dup = temp.CreateViewTemplate()

    dup.Name = view_name

    return dup



def browser_organization(dup_view, purpose, trade, type):

    view_purpose = dup_view.LookupParameter("Browser View Purpose")

    view_purpose.Set(purpose)

    trade_designation = dup_view.LookupParameter("Browser Trade Designation")

    trade_designation.Set(trade)

    view_type = dup_view.LookupParameter("Browser View Type")

    #view_type.Set(type)

    if type == "":

        non_control = dup_view.GetNonControlledTemplateParameterIds()

        updated_list = List[DB.ElementId]()

        for non in non_control:

            updated_list.Add(non)

        if view_type.Id not in list(updated_list):

            updated_list.Add(view_type.Id)

        dup_view.SetNonControlledTemplateParameterIds(updated_list)

    else:

        view_type.Set(type)



def set_discipline(dup_view, trade):

    if trade == "P" or trade == "GP" or trade == "PP" or trade == "MG":

        dup_view.Discipline = DB.ViewDiscipline.Plumbing

    elif trade == "MD" or trade == "MP":

        dup_view.Discipline = DB.ViewDiscipline.Mechanical



def set_worksets(worksets, dup_view, type):

    for w in worksets:

        if w.Name == "Mechanical - Ductwork": # Globally this should always be on

            dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

        elif w.Name == "Mechanical - Piping": # Globally this should always be on

            dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

        elif w.Name == "Plumbing": # Globally this should always be on

            dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

        elif w.Name == "QAQC":

            if type == "QAQC":

                dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.Visible)

            else: # Globally this should always be OFF

                dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

        elif w.Name == "z-ARCH/STRUCT":

            if type == "Export":

                dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.Hidden)

            else: # Globally this should always be on

                dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

        elif w.Name == "z-Design Links":

            if type == "Working": # Globally this should always be on

                dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

            else:

                dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.Hidden)

        elif w.Name == "z-NFC": # Globally this should always be OFF

            dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

        elif w.Name == "z-Scope Boxes": # Globally this should always be OFF

            dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

        elif w.Name == "z-Shared Levels and Grids":

            if type == "Export":

                dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.Hidden)

            else: # Globally this should always be on

                dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)

        elif w.Name == "z-Trimble Points": # Globally this should always be OFF

            dup_view.SetWorksetVisibility(w.Id, DB.WorksetVisibility.UseGlobalSetting)



def set_filters(dup_view, type, trade):

    filters = dup_view.GetFilters()

    for f in filters: # All Filters should be on in the Master Template

        n = doc.GetElement(f)

        filter_name = n.Name



        if filter_name.startswith("Overall"):

            if "Insulation" in filter_name:

                if type =="Working" or type == "Export" or type == "Sleeves":

                    dup_view.SetFilterVisibility(f, True)

                else:

                    dup_view.SetFilterVisibility(f, False)

            else:

                dup_view.SetFilterVisibility(f, True)

        

        elif filter_name.startswith("__"):

            if filter_name == "__QAQC":

                if type == "QAQC":

                    dup_view.SetFilterVisibility(f, True)

                else:

                    dup_view.SetFilterVisibility(f, False)

            elif filter_name == "__Hangers - All Hangers (FAB)":

                if type == "Working" or type == "Export" or type == "Hangers" or type == "QAQC":

                    dup_view.SetFilterVisibility(f, True)

                else:

                    dup_view.SetFilterVisibility(f, False)

            elif filter_name == "__Sleeves - All Sleeves (FAB)":

                if type == "Working" or type == "Export" or type == "Sleeves" or type == "QAQC":

                    dup_view.SetFilterVisibility(f, True)

                else:

                    dup_view.SetFilterVisibility(f, False)

            elif filter_name == "__Only Sleeves - Only Show Sleeves (FAB)": # Reverse

                if type == "Sleeves":

                    dup_view.SetFilterVisibility(f, False)

                else:

                    dup_view.SetFilterVisibility(f, True)

            elif filter_name == "__HK Pad":

                if type == "Pad Layout"  or type == "Export" or type == "Working" or type == "QAQC":

                    dup_view.SetFilterVisibility(f, True)

                else:

                    dup_view.SetFilterVisibility(f, False)

        

        elif filter_name.startswith("_All Sheets"):

            dup_view.SetFilterVisibility(f, False)

        

        else:

            if trade == "ALL": # All Trades on

                if filter_name.startswith("MPipe") or filter_name.startswith("Process") or filter_name.startswith("Steam") or filter_name.startswith("Plumbing") or filter_name.startswith("Drainage") or filter_name.startswith("Gases") or filter_name.startswith("Duct"):

                    dup_view.SetFilterVisibility(f, True)

                elif filter_name.startswith("EQUIP"): # All Equipment on

                    dup_view.SetFilterVisibility(f, True)

                elif filter_name.startswith("_Sheets"):

                    if type == "Export":

                        dup_view.SetFilterVisibility(f, False)

                    else:

                        if "ALL" in filter_name:

                            dup_view.SetFilterVisibility(f, True)

                        else:

                            dup_view.SetFilterVisibility(f, False)



            elif trade == "MD":

                if filter_name.startswith("Duct"):

                    dup_view.SetFilterVisibility(f, True)

                elif filter_name.startswith("MPipe") or filter_name.startswith("Process") or filter_name.startswith("Steam") or filter_name.startswith("Plumbing") or filter_name.startswith("Drainage") or filter_name.startswith("Gases"):

                    dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("EQUIP"):

                    if "MD" in filter_name: # All MD on

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("_Sheets"):

                    if "MD" in filter_name:

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)



            elif trade == "MP":

                if filter_name.startswith("MPipe") or filter_name.startswith("Process") or filter_name.startswith("Steam"):

                    dup_view.SetFilterVisibility(f, True)

                elif filter_name.startswith("Duct") or filter_name.startswith("Plumbing") or filter_name.startswith("Drainage") or filter_name.startswith("Gases"):

                    dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("EQUIP"):

                    if "MP" in filter_name: # All MP on

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("_Sheets"):

                    if "MP" in filter_name:

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)



            elif trade == "P":

                if filter_name.startswith("Drainage") or filter_name.startswith("Plumbing"):

                    dup_view.SetFilterVisibility(f, True)

                elif filter_name.startswith("MPipe") or filter_name.startswith("Process") or filter_name.startswith("Steam") or filter_name.startswith("Duct") or filter_name.startswith("Gases"):

                    dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("EQUIP"):

                    if "_P" in filter_name or "/P" in filter_name: # All P on

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("_Sheets"):

                    if filter_name == "_Sheets - P Sections and Callouts":

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)



            elif trade == "PP":

                if filter_name.startswith("Plumbing"):

                    dup_view.SetFilterVisibility(f, True)

                elif filter_name.startswith("MPipe") or filter_name.startswith("Process") or filter_name.startswith("Steam") or filter_name.startswith("Duct") or filter_name.startswith("Drainage") or filter_name.startswith("Gases"):

                    dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("EQUIP"):

                    if "_P" in filter_name or "/P" in filter_name: # All P on

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("_Sheets"):

                    if "PP" in filter_name:

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)



            elif trade == "GP":

                if filter_name.startswith("Drainage"):

                    dup_view.SetFilterVisibility(f, True)

                elif filter_name.startswith("MPipe") or filter_name.startswith("Process") or filter_name.startswith("Steam") or filter_name.startswith("Plumbing") or filter_name.startswith("Duct") or filter_name.startswith("Gases"):

                    dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("EQUIP"):

                    if "_P" in filter_name or "/P" in filter_name: # All P on

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("_Sheets"):

                    if "GP" in filter_name:

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)



            elif trade == "MG":

                if filter_name.startswith("Gases"):

                    dup_view.SetFilterVisibility(f, True)

                elif filter_name.startswith("MPipe") or filter_name.startswith("Process") or filter_name.startswith("Steam") or filter_name.startswith("Plumbing") or filter_name.startswith("Drainage") or filter_name.startswith("Duct"):

                    dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("EQUIP"):

                    if "MG" in filter_name: # All MG on

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)

                elif filter_name.startswith("_Sheets"):

                    if "MG" in filter_name:

                        dup_view.SetFilterVisibility(f, True)

                    else:

                        dup_view.SetFilterVisibility(f, False)





#######################



#   __      __     _____  _____ ____  _      ______  _____ 

#   \ \    / /\   |  __ \|_   _|  _ \| |    |  ____|/ ____|

#    \ \  / /  \  | |__) | | | | |_) | |    | |__  | (___  

#     \ \/ / /\ \ |  _  /  | | |  _ <| |    |  __|  \___ \ 

#      \  / ____ \| | \ \ _| |_| |_) | |____| |____ ____) |

#       \/_/    \_\_|  \_\_____|____/|______|______|_____/ 

#                                                          



#######################



start_time = time.time() # Start time to track how long script takes to run



doc = __revit__.ActiveUIDocument.Document

app = __revit__.Application



#######################



#     _____ ____  _____  ______ 

#    / ____/ __ \|  __ \|  ____|

#   | |   | |  | | |  | | |__   

#   | |   | |  | | |  | |  __|  

#   | |___| |__| | |__| | |____ 

#    \_____\____/|_____/|______|

#                               



#######################



# Safe Run

@error_handling.FERG_safe_run

def my_code_run():



    # Trades

    trades = ["ALL", "MD", "MP", "P", "PP", "GP", "MG"]



    # View Template Types

    types = ["Working", "Export", "Sleeves", "Shop Drawings", "Hangers", "Pad Layout", "QAQC"]



    views = DB.FilteredElementCollector(doc).OfClass(DB.View)



    viewTemplates = [v for v in views if v.IsTemplate == True]



    view_temp_dict = {v.Name:v for v in viewTemplates}



    select_master = forms.SelectFromList.show(sorted(view_temp_dict),title='Select Master View Template', multiselect=False, button_name='Select')



    master_view_template = view_temp_dict.get(select_master)



    select_trades = forms.SelectFromList.show(sorted(trades),title='Select Trades', multiselect=True, button_name='Select')



    select_types = forms.SelectFromList.show(sorted(types),title='Select Types', multiselect=True, button_name='Select')



    view_family_types = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).ToElements()



    workset_list = DB.FilteredWorksetCollector(doc).ToWorksets()



    worksets = [x for x in workset_list if str(x.Kind) == "UserWorkset"]



    # Getting the 3D view type to be used to create the temp 3D view later

    vft_3d = None

    for vft in view_family_types:

        if vft.ViewFamily == DB.ViewFamily.ThreeDimensional:

            vft_3d = vft

            break



    # Start Transaction

    t = DB.Transaction(doc, 'Create View Template(s)')

    t.Start()



    count_variable = 0 # Counting templates made



    for type in select_types:

        for trade in select_trades:



            if type == "Working":

                view_name = "!Working - " + trade + " (XXX)"

                dup_view = duplicate_view_template(master_view_template, view_name)

                set_worksets(worksets, dup_view, type)

                set_filters(dup_view, type, trade)

                browser_organization(dup_view, "_Working Views", trade, "")

                set_discipline(dup_view, trade)

                count_variable += 1

            elif type == "Export":

                if vft_3d != None:

                    view_name = "00_Exporting 3D View - " + trade

                    # Need to create a 3D view here to create a 3D view template

                    temp3D = DB.View3D.CreateIsometric(doc, vft_3d.Id)

                    temp3D.ApplyViewTemplateParameters(master_view_template)

                    dup_view = duplicate_view_template(temp3D, view_name)

                    doc.Delete(temp3D.Id) # Delete temporary view made

                    set_worksets(worksets, dup_view, type)

                    set_filters(dup_view, type, trade)

                    browser_organization(dup_view, "Export Views", trade, "3D")

                    set_discipline(dup_view, trade)

                    count_variable += 1

            elif type == "Sleeves":

                view_name = "01_Sleeve Sheet View - " + trade

                dup_view = duplicate_view_template(master_view_template, view_name)

                set_worksets(worksets, dup_view, type)

                set_filters(dup_view, type, trade)

                browser_organization(dup_view, "Sheet Views - Sleeves", trade, "Plan")

                set_discipline(dup_view, trade)

                count_variable += 1

            elif type == "Shop Drawings":

                view_name = "02_Sheet View - " + trade

                dup_view = duplicate_view_template(master_view_template, view_name)

                set_worksets(worksets, dup_view, type)

                set_filters(dup_view, type, trade)

                browser_organization(dup_view, "Sheet Views", trade, "Plan")

                set_discipline(dup_view, trade)

                count_variable += 1

            elif type == "Hangers":

                view_name = "03_Hanger Sheet View - " + trade

                dup_view = duplicate_view_template(master_view_template, view_name)

                set_worksets(worksets, dup_view, type)

                set_filters(dup_view, type, trade)

                browser_organization(dup_view, "Sheet Views - Hangers", trade, "Plan")

                set_discipline(dup_view, trade)

                count_variable += 1

            elif type == "Pad Layout":

                view_name = "04_Pad Layout Sheet View - " + trade

                dup_view = duplicate_view_template(master_view_template, view_name)

                set_worksets(worksets, dup_view, type)

                set_filters(dup_view, type, trade)

                browser_organization(dup_view, "Sheet Views - Pad Layouts", trade, "Plan")

                set_discipline(dup_view, trade)

                count_variable += 1

            elif type == "QAQC":

                view_name = "!QAQC - " + trade

                dup_view = duplicate_view_template(master_view_template, view_name)

                set_worksets(worksets, dup_view, type)

                set_filters(dup_view, type, trade)

                browser_organization(dup_view, "_QAQC Views", trade, "")

                set_discipline(dup_view, trade)

                count_variable += 1



    # End Transaction

    t.Commit()

    

    return count_variable # This is len() for list, but for numbers/floats just return variable



result, error = my_code_run() #result is the count variable for tracking





out_dialog = UI.TaskDialog("View Template Creation")

out_dialog.MainIcon = UI.TaskDialogIcon.TaskDialogIconInformation

out_dialog.MainInstruction = "Success"

out_dialog.MainContent = str(result) + " View Templates created."

#out_dialog.ExpandedContent = details

#out_dialog.FooterText = ""

out_dialog.Show()





#######################



#     ____  _    _ _______ _____  _    _ _______  ___      ____   _____ 

#    / __ \| |  | |__   __|  __ \| |  | |__   __|/ / |    / __ \ / ____|

#   | |  | | |  | |  | |  | |__) | |  | |  | |  / /| |   | |  | | |  __ 

#   | |  | | |  | |  | |  |  ___/| |  | |  | | / / | |   | |  | | | |_ |

#   | |__| | |__| |  | |  | |    | |__| |  | |/ /  | |___| |__| | |__| |

#    \____/ \____/   |_|  |_|     \____/   |_/_/   |______\____/ \_____|

#                                                                       



#######################



if error:

    try:

        tracking.pyRevit_log_tracking(start_time, result, pyRevit_tool_name, "Error", error)

    except Exception as log_ex:

        try:

            tracking.pyRevit_error_collecting_tracking(pyRevit_tool_name,str(log_ex),traceback.format_exc())

        except:

            try:

                final_dialog(pyRevit_tool_name)

            except:

                pass



else:

    try:

        tracking.pyRevit_log_tracking(start_time, result, pyRevit_tool_name, "Success", "Succesful")

    except Exception as log_ex:

        try:

            tracking.pyRevit_error_collecting_tracking(pyRevit_tool_name,str(log_ex),traceback.format_exc())

        except:

            try:

                final_dialog(pyRevit_tool_name)

            except:

                pass

