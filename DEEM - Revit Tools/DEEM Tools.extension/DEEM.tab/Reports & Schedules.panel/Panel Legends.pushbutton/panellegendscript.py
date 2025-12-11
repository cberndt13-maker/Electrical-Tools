# -*- coding: utf-8 -*-
"""
Panel Legends
Creates DEEM-formatted panel legend drafting views named for each panel.

This script replaces the previous Excel export behaviour and creates a
drafting view per panel (or for a selected panel). It uses best-effort
fallbacks for TextNote and FilledRegion types and reports errors clearly.
"""

from __future__ import absolute_import
import sys
import traceback
from pyrevit import forms, revit
from Autodesk.Revit.DB import *
from System.Collections.Generic import List

# ---------------------------------------------------------------------------
# Configuration (simple, IronPython-compatible dictionary)
# ---------------------------------------------------------------------------
CONFIG = {
    'panel_settings': {
        'default': {
            'slots': 42,
            'fed_from_text': '120/208 V - 3 Ph - 4 W',
            'phase_colors': {
                'A': '#ff6666',
                'B': '#66b3ff',
                'C': '#66ff66',
                'N': '#ffff66',
                'G': '#bfbfbf',
            }
        },
        'Main Panel': {'slots': 48, 'fed_from_text': '480/277 V - 3 Ph - 4 W'},
        'Sub Panel': {'slots': 36},
    },
    'user_prompt': {
        'select_mode': True,
        'alert_on_complete': True,
    },
    'behavior': {
        'overwrite_views': True,
    }
}

USER_PROMPT = CONFIG.get('user_prompt', {})
OVERWRITE_VIEWS = CONFIG.get('behavior', {}).get('overwrite_views', True)
PANEL_SETTINGS = CONFIG.get('panel_settings', {})
DEFAULT_PANEL_SETTINGS = PANEL_SETTINGS.get('default', {})

# ---------------------------------------------------------------------------
# Layout tuning (tweak these values to adjust spacing, box sizes, and offsets)
# ---------------------------------------------------------------------------
LAYOUT = {
    # vertical spacing between rows (increase to reduce overlap)
    'line_h': 0.34,
    # box half-width and half-height (box will be box_w x box_h)
    'box_w': 0.26,
    'box_h': 0.14,
    # horizontal positions for columns (can be tuned)
    'load_left_x': -1.6,
    'brk_left_x': 2.6,
    'center_left_x': 4.9,
    'center_right_x': 5.6,
    'brk_right_x': 6.3,
    'load_right_x': 8.3,
    # small x-offset applied when placing circuit numbers to better center them
    'number_x_offset': 0.08,
    # header vertical offsets
    'title_y': 0.0,
    'header_y_offset': 0.42,
    # distance from header to first row start
    'rows_start_offset': 0.36,
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def get_panel_config(panel_name):
    for key, settings in PANEL_SETTINGS.items():
        if key != 'default' and key.lower() in (panel_name or '').lower():
            return settings
    return DEFAULT_PANEL_SETTINGS


def get_all_panels(doc):
    fec = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    return [p for p in fec if p is not None]


def get_panel_circuits(panel):
    circuits = []
    try:
        systems = getattr(panel, 'ElectricalSystems', None)
        if systems:
            for s in systems:
                try:
                    circuits.append({
                        'Circuit': getattr(s, 'CircuitNumber', None),
                        'Load': getattr(s, 'LoadName', '') or '',
                        'Breaker': (s.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_BREAKER_PARAM).AsValueString() if s.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_BREAKER_PARAM) else ''),
                        'Phase': (s.get_Parameter(BuiltInParameter.RBS_ELEC_PHASE_PARAM).AsString() if s.get_Parameter(BuiltInParameter.RBS_ELEC_PHASE_PARAM) else ''),
                        'LoadValue': (s.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD_PARAM).AsValueString() if s.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD_PARAM) else ''),
                    })
                except:
                    continue
    except:
        pass
    return circuits


def get_view_family_type(doc, view_family):
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        try:
            if vft.ViewFamily == view_family:
                return vft
        except:
            continue
    return None


def find_existing_drafting_view(doc, name):
    for v in FilteredElementCollector(doc).OfClass(ViewDrafting).ToElements():
        try:
            if v.Name == name:
                return v
        except:
            continue
    return None


def ensure_text_note_type(doc):
    types = FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()
    if types:
        return types[0].Id
    return ElementId.InvalidElementId


def ensure_filled_region_type(doc, base_name='DEEM_Fill'):
    for fr in FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements():
        try:
            if fr.Name == base_name:
                return fr.Id
        except:
            continue
    for fr in FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements():
        try:
            new_id = fr.Duplicate(base_name)
            return new_id
        except:
            continue
    return ElementId.InvalidElementId


# ---------------------------------------------------------------------------
# Create the Drafting View and draw the legend
# ---------------------------------------------------------------------------
def create_panel_drafting_view(doc, panel_name, circuits):
    vft = get_view_family_type(doc, ViewFamily.Drafting)
    if not vft:
        raise Exception('No drafting view family type in project')

    t = Transaction(doc, 'Create Panel Legend View')
    if t.Start() != TransactionStatus.Started:
        raise Exception('Failed to start transaction')

    try:
        # handle overwrite
        if OVERWRITE_VIEWS:
            existing = find_existing_drafting_view(doc, panel_name)
            if existing:
                try:
                    doc.Delete(existing.Id)
                except:
                    pass

        view = ViewDrafting.Create(doc, vft.Id)
        view.Name = panel_name

        # prepare types
        text_type_id = ensure_text_note_type(doc)
        any_fr = None
        for fr in FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements():
            any_fr = fr
            break
        any_fr_id = any_fr.Id if any_fr else ElementId.InvalidElementId

        # layout (values pulled from LAYOUT dict so they're easy to tune)
        title_y = LAYOUT.get('title_y', 0.0)
        header_y = title_y - LAYOUT.get('header_y_offset', 0.42)
        line_h = LAYOUT.get('line_h', 0.34)
        load_left_x = LAYOUT.get('load_left_x', -1.6)
        brk_left_x = LAYOUT.get('brk_left_x', 2.6)
        center_left_x = LAYOUT.get('center_left_x', 4.9)
        center_right_x = LAYOUT.get('center_right_x', 5.6)
        brk_right_x = LAYOUT.get('brk_right_x', 6.3)
        load_right_x = LAYOUT.get('load_right_x', 8.3)
        rows_start_y = header_y - LAYOUT.get('rows_start_offset', 0.36)

        # Title and header
        if text_type_id != ElementId.InvalidElementId:
            TextNote.Create(doc, view.Id, XYZ(load_left_x, title_y, 0), 'PANEL:', text_type_id)
            TextNote.Create(doc, view.Id, XYZ(3.0, title_y, 0), 'DEEM', text_type_id)
            TextNote.Create(doc, view.Id, XYZ(7.0, title_y, 0), DEFAULT_PANEL_SETTINGS.get('fed_from_text', '120/208V 3 PH 4W'), text_type_id)
            TextNote.Create(doc, view.Id, XYZ(load_left_x, header_y, 0), 'FED FROM:', text_type_id)
            TextNote.Create(doc, view.Id, XYZ(load_left_x, header_y - 0.18, 0), 'LOAD SERVED', text_type_id)
            TextNote.Create(doc, view.Id, XYZ(brk_left_x, header_y - 0.18, 0), 'BRK', text_type_id)
            TextNote.Create(doc, view.Id, XYZ((center_left_x + center_right_x) / 2.0, header_y - 0.18, 0), 'CIR', text_type_id)
            TextNote.Create(doc, view.Id, XYZ(brk_right_x, header_y - 0.18, 0), 'BRK', text_type_id)
            TextNote.Create(doc, view.Id, XYZ(load_right_x, header_y - 0.18, 0), 'LOAD SERVED', text_type_id)

        center_fr_id = ElementId.InvalidElementId
        if any_fr_id != ElementId.InvalidElementId:
            try:
                center_fr_id = doc.GetElement(any_fr_id).Duplicate('DEEM Phase CENTER')
            except:
                center_fr_id = any_fr_id

        settings = get_panel_config(panel_name)
        slots = int(settings.get('slots', DEFAULT_PANEL_SETTINGS.get('slots', 42)))
        rows = slots // 2

        # check for detail curve capability
        can_detail = True
        try:
            test_line = Line.CreateBound(XYZ(0, 0, 0), XYZ(0.05, 0, 0))
            el = doc.Create.NewDetailCurve(view, test_line)
            try:
                doc.Delete(el.Id)
            except:
                pass
        except:
            can_detail = False

        for r in range(rows):
            y = rows_start_y - r * line_h
            left_num = r * 2 + 1
            right_num = left_num + 1

            left_c = None
            right_c = None
            for cd in circuits:
                try:
                    n = int(cd.get('Circuit') or 0)
                    if n == left_num:
                        left_c = cd
                    elif n == right_num:
                        right_c = cd
                except:
                    continue

            if text_type_id != ElementId.InvalidElementId:
                TextNote.Create(doc, view.Id, XYZ(load_left_x, y, 0), left_c.get('Load') if left_c else '', text_type_id)
                TextNote.Create(doc, view.Id, XYZ(brk_left_x, y, 0), left_c.get('Breaker') if left_c else '', text_type_id)
                TextNote.Create(doc, view.Id, XYZ(load_right_x, y, 0), right_c.get('Load') if right_c else '', text_type_id)
                TextNote.Create(doc, view.Id, XYZ(brk_right_x, y, 0), right_c.get('Breaker') if right_c else '', text_type_id)

            # center filled boxes - use LAYOUT box sizes
            box_w = LAYOUT.get('box_w', 0.26)
            box_h = LAYOUT.get('box_h', 0.14)
            for cx in (center_left_x, center_right_x):
                if center_fr_id != ElementId.InvalidElementId:
                    try:
                        p1 = XYZ(cx - box_w, y - box_h, 0)
                        p2 = XYZ(cx + box_w, y - box_h, 0)
                        p3 = XYZ(cx + box_w, y + box_h, 0)
                        p4 = XYZ(cx - box_w, y + box_h, 0)
                        l1 = Line.CreateBound(p1, p2)
                        l2 = Line.CreateBound(p2, p3)
                        l3 = Line.CreateBound(p3, p4)
                        l4 = Line.CreateBound(p4, p1)
                        FilledRegion.Create(doc, center_fr_id, view.Id, List[CurveLoop]([CurveLoop.Create([l1, l2, l3, l4])]))
                        if can_detail:
                            try:
                                doc.Create.NewDetailCurve(view, l1)
                                doc.Create.NewDetailCurve(view, l2)
                                doc.Create.NewDetailCurve(view, l3)
                                doc.Create.NewDetailCurve(view, l4)
                            except:
                                can_detail = False
                    except:
                        pass

            if text_type_id != ElementId.InvalidElementId:
                # apply a small x-offset to better visually center numbers in the boxes
                num_off = LAYOUT.get('number_x_offset', 0.08)
                TextNote.Create(doc, view.Id, XYZ(center_left_x + num_off, y, 0), str(left_num), text_type_id)
                TextNote.Create(doc, view.Id, XYZ(center_right_x + num_off, y, 0), str(right_num), text_type_id)

        # Phase legend
        legend_x = center_right_x + 1.0
        legend_y = rows_start_y - rows * line_h - 1.2
        legend_items = list(DEFAULT_PANEL_SETTINGS.get('phase_colors', {}).items())
        ly = legend_y
        for key, _hexcol in legend_items:
            fr_id = ElementId.InvalidElementId
            if any_fr_id != ElementId.InvalidElementId:
                try:
                    fr_id = doc.GetElement(any_fr_id).Duplicate('DEEM Phase %s' % key)
                except:
                    fr_id = any_fr_id

            try:
                p1 = XYZ(legend_x, ly - 0.12, 0)
                p2 = XYZ(legend_x + 2.2, ly - 0.12, 0)
                p3 = XYZ(legend_x + 2.2, ly + 0.12, 0)
                p4 = XYZ(legend_x, ly + 0.12, 0)
                l1 = Line.CreateBound(p1, p2)
                l2 = Line.CreateBound(p2, p3)
                l3 = Line.CreateBound(p3, p4)
                l4 = Line.CreateBound(p4, p1)
                if fr_id != ElementId.InvalidElementId:
                    FilledRegion.Create(doc, fr_id, view.Id, List[CurveLoop]([CurveLoop.Create([l1, l2, l3, l4])]))
                    if can_detail:
                        try:
                            doc.Create.NewDetailCurve(view, l1)
                            doc.Create.NewDetailCurve(view, l2)
                            doc.Create.NewDetailCurve(view, l3)
                            doc.Create.NewDetailCurve(view, l4)
                        except:
                            pass
            except:
                pass

            if text_type_id != ElementId.InvalidElementId:
                TextNote.Create(doc, view.Id, XYZ(legend_x + 2.3, ly, 0), key, text_type_id)
            ly -= 0.35

        t.Commit()
        return view
    except Exception:
        try:
            t.RollBack()
        except:
            pass
        raise


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
doc = revit.doc

if USER_PROMPT.get('select_mode', True):
    mode_choice = forms.SelectFromList.show(['Single Panel', 'All Panels'], title='Choose mode', button_name='OK')
    if not mode_choice:
        forms.alert('Operation canceled by user.')
        sys.exit()
    single_mode = (mode_choice == 'Single Panel')
else:
    single_mode = False

all_panels = get_all_panels(doc)
if single_mode:
    if not all_panels:
        forms.alert('No panels found in project.')
        sys.exit()

    choices = []
    mapping = {}
    for p in all_panels:
        try:
            label = '%s | id:%s' % (p.Name if p.Name else '<no name>', p.Id.IntegerValue)
        except:
            label = '<unknown>'
        choices.append(label)
        mapping[label] = p

    choice = forms.SelectFromList.show(choices, title='Select Panel', button_name='Select')
    if not choice:
        forms.alert('Operation canceled by user.')
        sys.exit()
    panels = [mapping[choice]]
else:
    panels = all_panels

if not panels:
    forms.alert('No panels to process.')
    sys.exit()

created = []
errors = []
for p in panels:
    try:
        circuits = get_panel_circuits(p)
        name = p.Name if getattr(p, 'Name', None) else 'Panel'
        v = create_panel_drafting_view(doc, name, circuits)
        created.append(v.Name if v else name)
    except Exception as ex:
        errors.append((getattr(p, 'Name', '<unknown>'), str(ex), traceback.format_exc()))

if USER_PROMPT.get('alert_on_complete', True):
    if created:
        msg = 'Drafting views created:\n' + '\n'.join(created)
        if errors:
            msg += '\n\nSome panels failed:\n'
            msg += '\n'.join(['%s: %s' % (e[0], e[1]) for e in errors])
        forms.alert(msg)
    else:
        if errors:
            msg = 'No drafting views were created. Errors:\n'
            msg += '\n'.join(['%s: %s' % (e[0], e[1]) for e in errors])
            forms.alert(msg)
        else:
            forms.alert('No drafting views were created.')

