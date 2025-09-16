# -*- coding: utf-8 -*-
# Ventanas por habitación (bbox center ±offset adaptativo) + fallback + Comments
# Offsets probados (en pies): 1.0 m → 0.7 m → 0.4 m
# - Escribe en Comments: Department+Name de la habitación asignada, o "NA"
# - Excluye ventanas tipo/nombre "BUIT"
# - Informe por nivel y selección de no asignadas

import System
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit

doc  = revit.doc
uidoc = revit.uidoc

# ---- Config ----
OFFSETS_FT = [3.28084, 2.29659, 1.31234]  # 1.0 m, 0.7 m, 0.4 m (en pies)
DELTA_SILL_FT = 1.64042                   # 0.5 m
SELECT_UNASSIGNED = True
FALLBACK_ISPOINT = True

# ---- Utilidades ----
def rname(room):
    if not room: 
        return "(sin habitación)"
    p = room.get_Parameter(BuiltInParameter.ROOM_NAME)
    return (p.AsString() if p and p.AsString() else room.Name or "").strip()

def room_code(room):
    if not room: 
        return "NA"
    dep = room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT)
    name = room.get_Parameter(BuiltInParameter.ROOM_NAME)
    dep_val = dep.AsString() if dep and dep.AsString() else ""
    name_val = name.AsString() if name and name.AsString() else ""
    return "{}{}".format(dep_val.strip(), name_val.strip())

def is_balcony(name):
    n = (name or "").lower()
    return u"balcó" in n or u"terrassa" in n

def is_buit(window):
    try:
        wtype = doc.GetElement(window.GetTypeId())
        tname = wtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if wtype else ""
        famname = window.Symbol.FamilyName if hasattr(window, "Symbol") else ""
        return "buit" in ("{} {}".format(tname or "", famname or "")).lower()
    except:
        return False

def get_center_point(bb):
    return XYZ((bb.Min.X + bb.Max.X) * 0.5,
               (bb.Min.Y + bb.Max.Y) * 0.5,
               (bb.Min.Z + bb.Max.Z) * 0.5)

def get_sill_world(window):
    try:
        lvl = None
        if hasattr(window, "LevelId") and window.LevelId and window.LevelId != ElementId.InvalidElementId:
            lvl = doc.GetElement(window.LevelId)
        lvl_elev = lvl.Elevation if lvl else 0.0
        sill = window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
        if sill and sill.HasValue:
            return lvl_elev + sill.AsDouble()
    except:
        pass
    return None

def room_at_point(pt, rooms_cache):
    r = doc.GetRoomAtPoint(pt)
    if r or not FALLBACK_ISPOINT:
        return r
    for rr in rooms_cache:
        try:
            if rr.IsPointInRoom(pt):
                return rr
        except:
            continue
    return None

# ---- Datos base ----
rooms = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType())

windows = [
    w for w in FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Windows)
    .WhereElementIsNotElementType()
    if not is_buit(w)
]

room_to_windows = {}
unassigned = []  # (win, motivo, extra)

# ---- Transacción ----
t = Transaction(doc, "Asignar ventanas (offset adaptativo) y escribir Comments")
t.Start()

for win in windows:
    assigned_room = None
    try:
        bb = win.get_BoundingBox(None)
        if not bb:
            unassigned.append((win, "sin_bbox", "")); 
        else:
            center_bb = get_center_point(bb)
            n = win.FacingOrientation
            if not n:
                unassigned.append((win, "sin_facing", "")) 
            else:
                try:
                    n = n.Normalize()
                except:
                    unassigned.append((win, "sin_facing_normalizable", "")); n = None

                if n is not None:
                    # Alturas: Z centro + (sill + 0.5 m) si existe
                    heights = [center_bb.Z]
                    sill_world = get_sill_world(win)
                    if sill_world is not None:
                        heights.append(sill_world + DELTA_SILL_FT)

                    decided = False
                    # Probar offsets decrecientes y alturas
                    for off in OFFSETS_FT:
                        if decided: break
                        for z in heights:
                            c = XYZ(center_bb.X, center_bb.Y, z)
                            p_in  = c + n.Multiply(off)
                            p_out = c - n.Multiply(off)
                            r_in  = room_at_point(p_in, rooms)
                            r_out = room_at_point(p_out, rooms)

                            # Reglas de decisión
                            if r_in and r_out and r_in.Id == r_out.Id:
                                # Mismo room a ambos lados → probar siguiente altura/offset
                                continue

                            if r_in and r_out:
                                rin_name  = rname(r_in); rout_name = rname(r_out)
                                in_is_balcony  = is_balcony(rin_name)
                                out_is_balcony = is_balcony(rout_name)
                                if in_is_balcony ^ out_is_balcony:
                                    assigned_room = r_in if not in_is_balcony else r_out
                                    decided = True
                                    break
                                else:
                                    # ninguna/both balcón → probar siguiente
                                    continue
                            elif r_in or r_out:
                                # Una sola room → asignamos (aunque sea balcón)
                                assigned_room = r_in if r_in else r_out
                                decided = True
                                break
                            else:
                                # ninguna → probar siguiente
                                continue

                    if not decided and not assigned_room:
                        unassigned.append((win, "no_asignada", ""))

    except Exception as e:
        unassigned.append((win, "error", str(e)))
        assigned_room = None

    # --- Escribir Comments ---
    try:
        cparam = win.LookupParameter("Comments")
        if cparam and cparam.StorageType == StorageType.String:
            if assigned_room:
                cparam.Set(room_code(assigned_room))
                room_to_windows.setdefault(assigned_room.Id, []).append(win)
            else:
                cparam.Set("NA")
    except:
        pass

t.Commit()

# ---- Informe por nivel ----
by_level = {}
for room in rooms:
    rid = room.Id
    count = len(room_to_windows.get(rid, []))
    try:
        lvl = doc.GetElement(room.LevelId)
        lvl_name = lvl.Name if lvl else "(sin nivel)"
    except:
        lvl_name = "(sin nivel)"

    if lvl_name not in by_level:
        by_level[lvl_name] = {0: [], 1: [], 2: [], "2+": []}

    if count == 0:
        by_level[lvl_name][0].append(room)
    elif count == 1:
        by_level[lvl_name][1].append(room)
    elif count == 2:
        by_level[lvl_name][2].append(room)
    else:
        by_level[lvl_name]["2+"].append(room)

# ---- Salida ----
print("==============================================")
print("VENTANAS NO ASIGNADAS (offset adaptativo ≤1 m, fallback, excl. BUIT)")
print("Total: {}".format(len(unassigned)))
print("==============================================")
for win, reason, extra in unassigned:
    try:
        wtype = doc.GetElement(win.GetTypeId())
        tname = wtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if wtype else "(sin tipo)"
    except:
        tname = "(sin tipo)"
    print("Ventana Id: {} | Tipo: {} | Motivo: {} {}".format(
        win.Id.IntegerValue, tname, reason, ("| " + extra) if extra else "")
    )

print("\n==============================================")
print("HABITACIONES POR NIVEL SEGÚN Nº DE VENTANAS (post-asignación)")
print("==============================================")
for lvl, buckets in by_level.items():
    print("\nNivel: {}".format(lvl))
    print("  0 ventanas : {}".format(len(buckets[0])))
    print("  1 ventana  : {}".format(len(buckets[1])))
    print("  2 ventanas : {}".format(len(buckets[2])))
    print("  >2 ventanas: {}".format(len(buckets["2+"])))

# Selección opcional
if SELECT_UNASSIGNED and unassigned:
    ids = List[ElementId]()
    for w, _, _ in unassigned:
        ids.Add(w.Id)
    uidoc.Selection.SetElementIds(ids)
    uidoc.ShowElements(ids)

TaskDialog.Show("pyRevit", "Hecho. Offsets: 1.0/0.7/0.4 m (adaptativo). Fallback activo.\n{} ventanas no asignadas.".format(len(unassigned)))
