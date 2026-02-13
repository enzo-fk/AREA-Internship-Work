# -*- coding: utf-8 -*-
# pyRevit button script: ALL pipes in active view -> place supports by spacing
# + connect to CEILINGS (raycast up + set rod/drop length param)
# Revit 2024 compatible (pyRevit IronPython)

from pyrevit import revit, forms

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    FamilySymbol, Level, Line, XYZ, UnitUtils, UnitTypeId,
    ElementTransformUtils, Transaction,
    View3D, ViewFamilyType, ViewFamily,
    ElementCategoryFilter, ReferenceIntersector, FindReferenceTarget
)
from Autodesk.Revit.DB.Structure import StructuralType

doc = revit.doc


# -----------------------------
# Units helpers
# -----------------------------
def mm_to_internal(mm_val):
    """Millimeters -> Revit internal (feet)."""
    return UnitUtils.ConvertToInternalUnits(float(mm_val), UnitTypeId.Millimeters)

def pick_number_mm(prompt, default_mm):
    s = forms.ask_for_string(
        default=str(default_mm),
        prompt=prompt + " (mm)",
        title="Pipe Supports"
    )
    if s is None:
        return None
    try:
        return float(s.strip())
    except Exception:
        forms.alert("Invalid number: {}".format(s), title="Pipe Supports")
        return None


# -----------------------------
# Geometry helpers
# -----------------------------
def get_nearest_level(z_internal):
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    if not levels:
        return None
    levels.sort(key=lambda lv: abs(lv.Elevation - z_internal))
    return levels[0]

def project_to_xy(v):
    return XYZ(v.X, v.Y, 0.0)

def signed_angle_xy(from_vec, to_vec):
    """Signed angle in XY plane from from_vec to to_vec around +Z."""
    a = project_to_xy(from_vec)
    b = project_to_xy(to_vec)
    if a.GetLength() < 1e-9 or b.GetLength() < 1e-9:
        return 0.0
    a = a.Normalize()
    b = b.Normalize()
    angle = a.AngleTo(b)  # unsigned
    cross_z = a.CrossProduct(b).Z
    return -angle if cross_z < 0 else angle


# -----------------------------
# Find a usable 3D view (needed for ReferenceIntersector)
# -----------------------------
def get_any_3d_view(create_if_missing=True):
    # Prefer an existing non-template 3D view
    for v in FilteredElementCollector(doc).OfClass(View3D):
        try:
            if not v.IsTemplate:
                return v
        except Exception:
            continue

    if not create_if_missing:
        return None

    # Create an isometric 3D view if none exist
    vft = None
    for t in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        try:
            if t.ViewFamily == ViewFamily.ThreeDimensional:
                vft = t
                break
        except Exception:
            continue

    if vft is None:
        return None

    # Must be in a transaction to create a view
    # We'll create it in the main transaction if needed; here we just return a marker (None)
    return ("__CREATE__", vft.Id)


# -----------------------------
# Family symbol picker
# -----------------------------
def collect_support_symbols():
    # Supports/hangers are often modeled as Generic Models or Pipe Accessories.
    allowed_cats = set([
        int(BuiltInCategory.OST_GenericModel),
        int(BuiltInCategory.OST_PipeAccessory),
        int(BuiltInCategory.OST_SpecialityEquipment),
        int(BuiltInCategory.OST_MechanicalEquipment),
    ])

    symbols = []
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            cat_id = sym.Category.Id.IntegerValue if sym.Category else None
            if cat_id in allowed_cats:
                fam_name = sym.Family.Name if sym.Family else "?"
                type_name = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                display = "[{}] {} : {}".format(sym.Category.Name, fam_name, type_name)
                symbols.append((display, sym))
        except Exception:
            continue

    symbols.sort(key=lambda x: x[0].lower())
    return symbols

def pick_symbol(symbol_pairs):
    if not symbol_pairs:
        forms.alert(
            "No Family Types found in Generic Models / Pipe Accessories / Speciality / Mechanical.\n"
            "Load your support family into the project first.",
            title="Pipe Supports"
        )
        return None

    display_items = [p[0] for p in symbol_pairs]
    choice = forms.SelectFromList.show(
        display_items,
        title="Select Support Family Type",
        multiselect=False,
        width=900
    )
    if not choice:
        return None
    idx = display_items.index(choice)
    return symbol_pairs[idx][1]

def pick_parameter_name(symbol, purpose_text):
    """
    Lists TYPE parameters on the symbol as suggestions, but also lets you type manually.
    (Rod length is often an INSTANCE parameter, so manual entry is commonly needed.)
    """
    names = []
    try:
        for p in symbol.Parameters:
            try:
                if p.Definition and p.Definition.Name:
                    names.append(p.Definition.Name)
            except Exception:
                pass
    except Exception:
        pass

    names = sorted(set(names), key=lambda s: s.lower())
    options = ["<Skip>"] + names + ["<Type manually...>"]

    pick = forms.SelectFromList.show(
        options,
        title="Select parameter for {}".format(purpose_text),
        multiselect=False,
        width=700
    )
    if not pick:
        return None

    if pick == "<Skip>":
        return ""
    if pick == "<Type manually...>":
        typed = forms.ask_for_string(
            default="",
            prompt="Enter EXACT parameter name for {}".format(purpose_text),
            title="Pipe Supports"
        )
        if typed is None:
            return None
        return typed.strip()
    return pick


# -----------------------------
# Ceiling raycast helper
# -----------------------------
def find_ceiling_distance_up(view3d, origin_pt):
    """
    Raycast upward (+Z) to the nearest CEILING face.
    Returns distance (feet) or None if no ceiling hit.
    """
    try:
        cat_filter = ElementCategoryFilter(BuiltInCategory.OST_Ceilings)
        ri = ReferenceIntersector(cat_filter, FindReferenceTarget.Face, view3d)
        rwc = ri.FindNearest(origin_pt, XYZ.BasisZ)
        if rwc is None:
            return None
        # Proximity is the distance along the ray to the hit
        return rwc.Proximity
    except Exception:
        return None


# -----------------------------
# Main
# -----------------------------
def main():
    active_view = doc.ActiveView

    # 1) Collect ALL pipes in active view
    pipes = list(
        FilteredElementCollector(doc, active_view.Id)
        .OfCategory(BuiltInCategory.OST_PipeCurves)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    if not pipes:
        forms.alert("No pipes found in the active view.", title="Pipe Supports")
        return

    # 2) Pick support family type
    symbol_pairs = collect_support_symbols()
    symbol = pick_symbol(symbol_pairs)
    if symbol is None:
        return

    # 3) Inputs
    spacing_mm = pick_number_mm("Support spacing", 1500)
    if spacing_mm is None:
        return

    end_offset_mm = pick_number_mm("End offset (do not place within this distance from each end)", 200)
    if end_offset_mm is None:
        return

    clamp_clearance_mm = pick_number_mm("Clamp clearance (added to pipe diameter; 0 if none)", 0)
    if clamp_clearance_mm is None:
        return

    rod_clearance_mm = pick_number_mm("Ceiling clearance (subtract from rod length; 0 if none)", 0)
    if rod_clearance_mm is None:
        return

    # 4) Parameters
    clamp_param_name = pick_parameter_name(symbol, "Clamp/Pipe Diameter (optional)")
    if clamp_param_name is None:
        return

    rod_param_name = pick_parameter_name(symbol, "Rod/Drop Length (to reach ceiling)")
    if rod_param_name is None:
        return
    if not rod_param_name:
        forms.alert(
            "You must provide a rod/drop length parameter name to connect to ceilings.\n"
            "Pick '<Type manually...>' and enter the exact parameter name from the support family.",
            title="Pipe Supports"
        )
        return

    # Convert to internal units
    spacing = mm_to_internal(spacing_mm)
    end_offset = mm_to_internal(end_offset_mm)
    clamp_clearance = mm_to_internal(clamp_clearance_mm)
    rod_clearance = mm_to_internal(rod_clearance_mm)

    # 5) Need a 3D view for raycast
    view3d = get_any_3d_view(create_if_missing=True)

    t = Transaction(doc, "Auto Place Pipe Supports (Connect to Ceilings)")
    t.Start()
    try:
        # Create 3D view if missing
        if isinstance(view3d, tuple) and view3d[0] == "__CREATE__":
            vft_id = view3d[1]
            view3d = View3D.CreateIsometric(doc, vft_id)
            # keep it hidden; just used for raycast
            doc.Regenerate()

        if view3d is None or isinstance(view3d, tuple):
            raise Exception("No usable 3D view available for ceiling raycast (ReferenceIntersector).")

        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()

        created = 0
        skipped = 0
        no_curve = 0
        no_ceiling_hit = 0
        rod_set_fail = 0

        for pipe in pipes:
            loc = getattr(pipe, "Location", None)
            if not loc or not hasattr(loc, "Curve"):
                no_curve += 1
                continue

            curve = loc.Curve
            length = curve.Length
            if length < (2.0 * end_offset + 1e-6) or length < spacing:
                skipped += 1
                continue

            # Pipe diameter (internal feet)
            diam_param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
            if diam_param is None:
                skipped += 1
                continue
            pipe_diam = diam_param.AsDouble()

            d = end_offset
            while d < (length - end_offset - 1e-6):
                u = d / length
                pt = curve.Evaluate(u, True)

                # nearest level for placement
                level = get_nearest_level(pt.Z)
                if level is None:
                    levels = list(FilteredElementCollector(doc).OfClass(Level))
                    level = levels[0] if levels else None
                if level is None:
                    skipped += 1
                    d += spacing
                    continue

                # Place instance (non-hosted)
                inst = doc.Create.NewFamilyInstance(pt, symbol, level, StructuralType.NonStructural)

                # Rotate to align with pipe direction (XY)
                try:
                    deriv = curve.ComputeDerivatives(u, True)
                    tangent = deriv.BasisX
                    tang_xy = project_to_xy(tangent)
                    if tang_xy.GetLength() > 1e-9:
                        angle = signed_angle_xy(XYZ.BasisX, tang_xy)
                        axis = Line.CreateBound(pt, pt + XYZ.BasisZ)
                        ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle)
                except Exception:
                    pass

                # Set clamp diameter if requested
                if clamp_param_name:
                    target_val = pipe_diam + clamp_clearance
                    set_ok = False
                    try:
                        p_inst = inst.LookupParameter(clamp_param_name)
                        if p_inst and not p_inst.IsReadOnly:
                            p_inst.Set(target_val)
                            set_ok = True
                    except Exception:
                        pass
                    if not set_ok:
                        try:
                            p_type = symbol.LookupParameter(clamp_param_name)
                            if p_type and not p_type.IsReadOnly:
                                p_type.Set(target_val)
                                set_ok = True
                        except Exception:
                            pass

                # Raycast to ceiling and set rod length
                dist_up = find_ceiling_distance_up(view3d, pt)
                if dist_up is None:
                    no_ceiling_hit += 1
                else:
                    rod_len = dist_up - rod_clearance
                    if rod_len < 0:
                        rod_len = 0.0

                    set_ok = False
                    # Try instance rod param first
                    try:
                        rp_inst = inst.LookupParameter(rod_param_name)
                        if rp_inst and not rp_inst.IsReadOnly:
                            rp_inst.Set(rod_len)
                            set_ok = True
                    except Exception:
                        pass
                    # Then type rod param (less common)
                    if not set_ok:
                        try:
                            rp_type = symbol.LookupParameter(rod_param_name)
                            if rp_type and not rp_type.IsReadOnly:
                                rp_type.Set(rod_len)
                                set_ok = True
                        except Exception:
                            pass

                    if not set_ok:
                        rod_set_fail += 1

                created += 1
                d += spacing

        t.Commit()

    except Exception as ex:
        t.RollBack()
        forms.alert("Failed:\n{}".format(ex), title="Pipe Supports")
        return

    forms.alert(
        "Done.\n\n"
        "Active view: {}\n"
        "Pipes found in view: {}\n"
        "Supports created: {}\n"
        "Skipped pipes: {}\n"
        "Pipes without curve: {}\n"
        "No ceiling hit (raycast): {}\n"
        "Rod param set failed: {}\n\n"
        "Notes:\n"
        "- This places NON-HOSTED supports.\n"
        "- 'Connect to ceilings' here = rod/drop length computed to the nearest CEILING face.\n"
        "- If No ceiling hit is high, check you actually have Ceiling elements above the pipes."
        .format(active_view.Name, len(pipes), created, skipped, no_curve, no_ceiling_hit, rod_set_fail),
        title="Pipe Supports"
    )


if __name__ == "__main__":
    main()
