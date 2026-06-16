# -*- coding: utf-8 -*-
"""
OrthogonalWires.py  —  pyRevit Tool
=====================================
Fixes diagonal wire graphics in electrical plan views so all segments
run parallel or perpendicular to the building grid (N/S and E/W only).

PROBLEM THIS SOLVES
-------------------
When wires are placed between fixtures/devices that aren't on the same
row or column, Revit draws them as diagonal lines (zigzag). This looks
unprofessional on construction documents and misrepresents the actual
conduit/wire routing intent.

WHAT IT DOES
------------
Reads each wire's vertex list, detects diagonal segments, and inserts
an "elbow" vertex to convert each diagonal into an L-shaped orthogonal
path. Also removes redundant collinear vertices and sets WiringType to
Chamfer for clean right-angle display.

ROUTING MODES
-------------
  Auto (Recommended)
    Picks H-first or V-first per wire based on its dominant axis.
    A wire that runs more east-west than north-south will route
    horizontally first. Produces the most natural-looking results
    for warehouse/grid layouts like the one in the image.

  Horizontal First (force all)
    All wires: run east/west first, then north/south.

  Vertical First (force all)
    All wires: run north/south first, then east/west.

SELECTION BEHAVIOR
------------------
  • Pre-select specific wires  →  only those are processed
  • Nothing selected           →  all wires in the active view

REQUIREMENTS
------------
  Revit 2018+  (Wire.GetVertices / SetVertices API available)
  pyRevit 4.7+

BUNDLE SETUP
------------
  YourExtension.extension/
  └── Electrical.tab/
      └── Wire Tools.panel/
          └── Ortho Wires.pushbutton/
              └── script.py    ←  this file
"""

# ── imports ───────────────────────────────────────────────────────────────────
from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import (
    Transaction, XYZ, FilteredElementCollector, ViewType
)
from Autodesk.Revit.DB.Electrical import Wire, WiringType

# ── logger / output ───────────────────────────────────────────────────────────
logger = script.get_logger()
output = script.get_output()

# ── constants ─────────────────────────────────────────────────────────────────
ORTHO_TOL = 0.01   # feet — segments within this of H or V are already "clean"

# ── plan view types that make sense for this tool ─────────────────────────────
VALID_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
}


# ══════════════════════════════════════════════════════════════════════════════
#  GEOMETRY HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def is_horizontal(a, b):
    return abs(b.Y - a.Y) < ORTHO_TOL


def is_vertical(a, b):
    return abs(b.X - a.X) < ORTHO_TOL


def is_orthogonal(a, b):
    return is_horizontal(a, b) or is_vertical(a, b)


def elbow_point(a, b, routing):
    """
    Return the single intermediate XYZ vertex that converts the diagonal
    segment A→B into two orthogonal legs.

    routing: "auto" | "h_first" | "v_first"
    """
    dx = abs(b.X - a.X)
    dy = abs(b.Y - a.Y)
    z  = a.Z  # keep in the same plan elevation

    if routing == "h_first" or (routing == "auto" and dx >= dy):
        # Horizontal leg first: move to b.X while keeping a.Y
        return XYZ(b.X, a.Y, z)
    else:
        # Vertical leg first: move to b.Y while keeping a.X
        return XYZ(a.X, b.Y, z)


def remove_collinear(vertices):
    """
    Strip intermediate vertices that sit on the same H or V line as their
    neighbors. Prevents duplicate/redundant points that can accumulate when
    wires are processed more than once or when existing bends get rebuilt.
    """
    if len(vertices) <= 2:
        return list(vertices)

    result = [vertices[0]]

    for i in range(1, len(vertices) - 1):
        prev = result[-1]
        curr = vertices[i]
        nxt  = vertices[i + 1]

        # Collinear vertically: prev, curr, and next all share the same X
        col_v = (abs(curr.X - prev.X) < ORTHO_TOL and
                 abs(curr.X - nxt.X)  < ORTHO_TOL)
        # Collinear horizontally: share the same Y
        col_h = (abs(curr.Y - prev.Y) < ORTHO_TOL and
                 abs(curr.Y - nxt.Y)  < ORTHO_TOL)

        if not (col_v or col_h):
            result.append(curr)

    result.append(vertices[-1])
    return result


def orthogonalize_vertices(vertices, routing):
    """
    Full pipeline:
      1. Walk each consecutive vertex pair
      2. Insert elbow where diagonal is detected
      3. Strip redundant collinear points

    Returns new list, or the original list unchanged if already all-orthogonal.
    """
    if len(vertices) < 2:
        return list(vertices)

    # Step 1: insert elbows
    result = [vertices[0]]
    for i in range(1, len(vertices)):
        a = result[-1]
        b = vertices[i]
        if is_orthogonal(a, b):
            result.append(b)
        else:
            result.append(elbow_point(a, b, routing))
            result.append(b)

    # Step 2: remove any collinear points created by pre-existing bends
    result = remove_collinear(result)
    return result


def vertices_changed(old_verts, new_verts):
    """True if the new vertex list differs meaningfully from the old one."""
    if len(old_verts) != len(new_verts):
        return True
    for a, b in zip(old_verts, new_verts):
        if (abs(a.X - b.X) > ORTHO_TOL or
                abs(a.Y - b.Y) > ORTHO_TOL):
            return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
#  WIRE COLLECTION
# ══════════════════════════════════════════════════════════════════════════════

def collect_wires(doc, view):
    """
    Return a list of Wire elements to process.
    Priority: current selection (wires only) → all wires in active view.
    """
    sel_ids = revit.get_selection().element_ids
    sel_wires = []

    for eid in sel_ids:
        elem = doc.GetElement(eid)
        if isinstance(elem, Wire):
            sel_wires.append(elem)

    if sel_wires:
        output.print_md("**Using {} selected wire(s).**".format(len(sel_wires)))
        return sel_wires

    all_wires = list(
        FilteredElementCollector(doc, view.Id)
        .OfClass(Wire)
        .ToElements()
    )
    output.print_md(
        "**No selection — processing all {} wire(s) in active view.**"
        .format(len(all_wires))
    )
    return all_wires


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN PROCESSING
# ══════════════════════════════════════════════════════════════════════════════

def process_wires(doc, wires, routing):
    """
    Iterate wires, rebuild their vertex lists, set WiringType = Chamfer.

    Returns (modified_count, skipped_count, error_count).
    """
    modified = 0
    skipped  = 0
    errors   = 0

    with Transaction(doc, "Orthogonalize Wire Graphics") as t:
        t.Start()

        for wire in wires:
            try:
                old_verts = list(wire.GetVertices())

                if len(old_verts) < 2:
                    skipped += 1
                    continue

                new_verts = orthogonalize_vertices(old_verts, routing)

                if not vertices_changed(old_verts, new_verts):
                    skipped += 1
                    continue

                wire.SetVertices(new_verts)
                wire.WiringType = WiringType.Chamfer
                modified += 1

            except Exception as e:
                logger.error(
                    "Wire {} — {}: {}".format(
                        wire.Id.IntegerValue,
                        type(e).__name__,
                        str(e)
                    )
                )
                errors += 1

        t.Commit()

    return modified, skipped, errors


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

def main():
    doc   = revit.doc
    uidoc = revit.uidoc
    view  = uidoc.ActiveView

    # ── guard: plan views only ────────────────────────────────────────────────
    if view.ViewType not in VALID_VIEW_TYPES:
        forms.alert(
            "This tool only works on plan views.\n\n"
            "Active view type: {}\n\n"
            "Switch to a Floor Plan or Engineering Plan and try again."
            .format(view.ViewType),
            title="Orthogonal Wires — Wrong View Type",
            warn_icon=True,
        )
        return

    # ── routing mode dialog ───────────────────────────────────────────────────
    options = [
        "Auto (smart per-wire)",
        "Horizontal First  [→ then ↑]  force all",
        "Vertical First    [↑ then →]  force all",
    ]

    choice = forms.CommandSwitchWindow.show(
        options,
        message=(
            "Routing direction for diagonal wires:\n\n"
            "  Auto          — each wire chooses H or V first based on its\n"
            "                  dominant axis. Best for warehouse/grid layouts.\n\n"
            "  Horizontal First — all wires run east/west first, then north/south.\n\n"
            "  Vertical First   — all wires run north/south first, then east/west.\n\n"
            "Tip: for lighting circuits on a building grid (like the H1/H1E\n"
            "     pattern in plan views), Auto usually gives the cleanest result."
        ),
    )

    if not choice:
        return  # user cancelled

    routing_map = {
        options[0]: "auto",
        options[1]: "h_first",
        options[2]: "v_first",
    }
    routing = routing_map[choice]

    # ── collect & process ─────────────────────────────────────────────────────
    wires = collect_wires(doc, view)

    if not wires:
        forms.alert(
            "No wires found in the active view.\n"
            "Make sure you're on a view that contains electrical wire graphics.",
            title="Orthogonal Wires",
        )
        return

    modified, skipped, errors = process_wires(doc, wires, routing)

    # ── results summary ───────────────────────────────────────────────────────
    output.print_md("---")
    output.print_md("## Orthogonal Wire Cleanup — Results")
    output.print_md("")
    output.print_md("| | |")
    output.print_md("|---|---|")
    output.print_md("| **Wires straightened** | {} |".format(modified))
    output.print_md("| Already orthogonal (skipped) | {} |".format(skipped))
    output.print_md("| Errors | {} |".format(errors))
    output.print_md("| Routing mode | {} |".format(choice))
    output.print_md("| View | {} |".format(view.Name))

    if errors:
        output.print_md(
            "\n> ⚠ **{} wire(s) could not be modified.**  "
            "These may be pinned, owned by another user, or have "
            "connector constraints. Check the pyRevit output log for details."
            .format(errors)
        )
    elif modified == 0:
        output.print_md(
            "\n> ℹ All wires were already orthogonal — nothing to fix."
        )
    else:
        output.print_md(
            "\n> ✅ Done. Use **Ctrl+Z** to undo if the result isn't what you expected."
        )


# ─────────────────────────────────────────────────────────────────────────────
main()