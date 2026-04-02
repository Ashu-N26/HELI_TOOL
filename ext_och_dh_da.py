# ============================================================
# ext_och_dh_da.py
# OCH/DH/DA Calculation Engine
# Replaces: Sub EXTOCHDHDA()    — VBA Module41 (Non-RAF)
#           Sub EXTOCHDHDARAF() — VBA Module43 (RAF)
#
# Original VBA by: Alan Hutchinson
#   Module41: 17th January 2000 (amended 4 May, 8 May, 12 Jul 2000)
#   Module43: 17th January 2000 (amended 8 May, 12 Jul 2000)
#
# PURPOSE:
#   For each runway procedure (up to 15):
#     1. Select the higher of OCH or Published DH/MDH
#     2. Compare with System Minima — take the higher
#     3. Calculate Decision Altitude (DH + elevation)
#     4. Round UP to nearest 10ft
#
# RAF DIFFERENCE (Module43 vs Module41):
#   When glidepath angle present (D26 != 0):
#   → RAF uses pre-computed RAF DH columns (BP,BQ,BR,BS)
#     which include a GP-angle-based increment added on top
#   → Non-RAF uses standard DH columns (F,I,L,O)
#
# KEY CELL MAPPING (per runway row z, starting at z=26):
#   Input columns:
#     E  = Cat A OCH (Published)
#     F  = Cat A DH/MDH (Published)   [RAF uses BP if GP present]
#     H  = Cat B OCH
#     I  = Cat B DH/MDH               [RAF uses BQ if GP present]
#     K  = Cat C OCH
#     L  = Cat C DH/MDH               [RAF uses BR if GP present]
#     N  = Cat D OCH
#     O  = Cat D DH/MDH               [RAF uses BS if GP present]
#     Y  = System Minima
#     C  = Datum type ("T"=Threshold, else Aerodrome)
#     D  = GlidePath angle (RAF only)
#
#   Intermediate:
#     AG = Cat A derived DH   = MAX(OCH_A, SYS_MIN)
#     AF = Cat A derived DA   = DH + elevation (C8 or C10)
#     AJ = Cat B derived DH
#     AI = Cat B derived DA
#     AM = Cat C derived DH
#     AL = Cat C derived DA
#     AP = Cat D derived DH
#     AO = Cat D derived DA
#
#   Output (rounded to nearest 10):
#     C46+ = Cat A final DH
#     B46+ = Cat A final DA
#     G46+ = Cat B final DH
#     F46+ = Cat B final DA
#     K46+ = Cat C final DH
#     J46+ = Cat C final DA
#     O46+ = Cat D final DH
#     N46+ = Cat D final DA
#
# SPECIAL VALUES:
#   "-"  = Not applicable (OCH or DH not relevant for this cat)
#   "-1" = Category not used → output N/A
# ============================================================

import math

NA        = "N/A"
HYPHEN    = "-"
SKIP      = "-1"
DATUM_THR = "T"


def _parse(val):
    """Parse input value — returns float, NA, HYPHEN, SKIP, or None."""
    if val is None or str(val).strip() == "":
        return None
    s = str(val).strip()
    if s == HYPHEN:
        return HYPHEN
    if s == SKIP:
        return SKIP
    if s == NA:
        return NA
    try:
        return float(s)
    except ValueError:
        return s


def _roundup10(value):
    """
    Round UP to nearest 10ft.
    Replaces VBA: =ROUNDUP(value, -1)
    """
    if value == NA or value is None:
        return NA
    try:
        v = float(value)
        return int(math.ceil(v / 10.0) * 10)
    except (ValueError, TypeError):
        return NA


def _calc_dh(och, published_dh, system_minima):
    """
    Derived DH = MAX(OCH or DH, System Minima)

    VBA equivalent:
        =If(OCH > SYS_MIN, OCH, SYS_MIN)   when only OCH given
        =If(DH  > SYS_MIN, DH,  SYS_MIN)   when only DH given
        =If(MAX(OCH,DH) > SYS_MIN, MAX(OCH,DH), SYS_MIN) when both
    """
    if och == NA or published_dh == NA:
        return NA

    # Determine the base value from OCH and/or DH
    och_val = _parse(och)
    dh_val  = _parse(published_dh)
    sys_val = _parse(system_minima)

    # Both present — take higher
    if isinstance(och_val, float) and isinstance(dh_val, float):
        base = max(och_val, dh_val)
    elif isinstance(dh_val, float):
        base = dh_val
    elif isinstance(och_val, float):
        base = och_val
    else:
        return NA

    if isinstance(sys_val, float):
        return max(base, sys_val)
    return base


def _calc_da(derived_dh, datum_type, elev_aerodrome, elev_threshold):
    """
    Decision Altitude = DH + elevation

    VBA equivalent:
        =If(DH="N/A","N/A",If(datum="T", DH+$C$10, DH+$C$8))

    C8  = Aerodrome elevation
    C10 = Threshold elevation
    datum_type "T" → use threshold elevation (C10)
    otherwise     → use aerodrome elevation (C8)
    """
    if derived_dh == NA or derived_dh is None:
        return NA
    try:
        dh_num = float(derived_dh)
        if str(datum_type).strip().upper() == DATUM_THR:
            return dh_num + float(elev_threshold)
        else:
            return dh_num + float(elev_aerodrome)
    except (ValueError, TypeError):
        return NA


def _process_category(och, published_dh, system_minima,
                       datum_type, elev_aerodrome, elev_threshold,
                       raf_dh=None, glidepath_angle=None, is_raf=False):
    """
    Process one Cat (A/B/C/D) for one runway.

    Returns:
        {
            "derived_dh":  float or "N/A"  (intermediate)
            "derived_da":  float or "N/A"  (intermediate)
            "final_dh":    int   or "N/A"  (rounded, output)
            "final_da":    int   or "N/A"  (rounded, output)
        }
    """
    och_p = _parse(och)
    dh_p  = _parse(published_dh)

    # ── SKIP: both are -1 → N/A ──────────────────────────────
    # VBA Line2/Line5/Line8/Line11: both OCH and DH = -1
    if och_p == SKIP and dh_p == SKIP:
        return {"derived_dh": NA, "derived_da": NA,
                "final_dh": NA, "final_da": NA}

    # ── RAF with glidepath: use pre-computed RAF DH ──────────
    # VBA Module43: If Range(l) <> 0 Then e = "BP" & z
    # The raf_dh already includes the GP angle increment
    if is_raf and glidepath_angle and float(glidepath_angle) != 0 and raf_dh is not None:
        raf_dh_p = _parse(raf_dh)
        if raf_dh_p == SKIP:
            return {"derived_dh": NA, "derived_da": NA,
                    "final_dh": NA, "final_da": NA}
        derived_dh = raf_dh_p
    else:
        # ── Standard: derived DH = MAX(OCH or DH, System Min) ─
        # VBA Line1/Line3 (only OCH), Line2/Line3 (DH present)
        derived_dh = _calc_dh(och_p, dh_p, system_minima)

    derived_da = _calc_da(derived_dh, datum_type,
                          elev_aerodrome, elev_threshold)

    return {
        "derived_dh": derived_dh,
        "derived_da": derived_da,
        "final_dh":   _roundup10(derived_dh),
        "final_da":   _roundup10(derived_da),
    }


# ── PUBLIC FUNCTION: NON-RAF ─────────────────────────────────
def extoch_dh_da(procedures: list, aerodrome_elev: float,
                 threshold_elev: float) -> list:
    """
    Calculate OCH/DH/DA for all runways — Non-RAF route.
    Replaces: Sub EXTOCHDHDA() — Module41

    Called by RAFSELECT when C11 != "Y" and D26 is empty.

    Args:
        procedures:      list of dicts (max 15), each:
            {
                "runway_id":    str
                "datum_type":   str    "T" or "A"    ← C column
                "system_minima":float               ← Y column
                "och_a":  float or "-" or "-1"      ← E column
                "dh_a":   float or "-" or "-1"      ← F column
                "och_b":  float or "-" or "-1"      ← H column
                "dh_b":   float or "-" or "-1"      ← I column
                "och_c":  float or "-" or "-1"      ← K column
                "dh_c":   float or "-" or "-1"      ← L column
                "och_d":  float or "-" or "-1"      ← N column
                "dh_d":   float or "-" or "-1"      ← O column
            }
        aerodrome_elev:  float  ← Sheet1 C8
        threshold_elev:  float  ← Sheet1 C10

    Returns:
        list of result dicts per runway:
            {
                "runway_id": str
                "cat_a": { "final_dh", "final_da", "derived_dh", "derived_da" }
                "cat_b": { ... }
                "cat_c": { ... }
                "cat_d": { ... }
            }
    """
    results = []

    for proc in procedures[:15]:
        # Stop loop if Cat A DH/MDH is empty
        # VBA: If Range(e) = Empty Then RVRCALCULATION
        if proc.get("dh_a") in (None, ""):
            break

        sys_min    = proc.get("system_minima", 0)
        datum      = proc.get("datum_type", "A")
        runway_id  = proc.get("runway_id", "")

        result = {"runway_id": runway_id}

        for cat, och_key, dh_key in [
            ("cat_a", "och_a", "dh_a"),
            ("cat_b", "och_b", "dh_b"),
            ("cat_c", "och_c", "dh_c"),
            ("cat_d", "och_d", "dh_d"),
        ]:
            result[cat] = _process_category(
                och            = proc.get(och_key),
                published_dh   = proc.get(dh_key),
                system_minima  = sys_min,
                datum_type     = datum,
                elev_aerodrome = aerodrome_elev,
                elev_threshold = threshold_elev,
                is_raf         = False
            )

        results.append(result)

    return results


# ── PUBLIC FUNCTION: RAF ─────────────────────────────────────
def extoch_dh_da_raf(procedures: list, aerodrome_elev: float,
                     threshold_elev: float) -> list:
    """
    Calculate OCH/DH/DA for all runways — RAF route.
    Replaces: Sub EXTOCHDHDARAF() — Module43

    Called by RAFSELECT when C11 = "Y" and D26 has GP angle.

    Key difference from Non-RAF:
        When glidepath_angle != 0, uses pre-computed RAF DH
        values from BP/BQ/BR/BS columns which include the
        GP-angle increment on top of the standard DH.

    Args:
        procedures:  list of dicts (max 15), each:
            {
                "runway_id":       str
                "datum_type":      str    "T" or "A"   ← C column
                "system_minima":   float               ← Y column
                "glidepath_angle": float               ← D column
                "och_a":  float or "-" or "-1"         ← E column
                "dh_a":   float or "-" or "-1"         ← F column
                "raf_dh_a": float or "-1"              ← BP column
                "och_b":  float or "-" or "-1"         ← H column
                "dh_b":   float or "-" or "-1"         ← I column
                "raf_dh_b": float or "-1"              ← BQ column
                "och_c":  float or "-" or "-1"         ← K column
                "dh_c":   float or "-" or "-1"         ← L column
                "raf_dh_c": float or "-1"              ← BR column
                "och_d":  float or "-" or "-1"         ← N column
                "dh_d":   float or "-" or "-1"         ← O column
                "raf_dh_d": float or "-1"              ← BS column
            }
        aerodrome_elev:  float  ← Sheet1 C8
        threshold_elev:  float  ← Sheet1 C10

    Returns:
        Same structure as extoch_dh_da()
    """
    results = []

    for proc in procedures[:15]:
        # VBA: If Range(e) = Empty Then RVRCALCULATION
        if proc.get("dh_a") in (None, ""):
            break

        # RAF requires DH — "-" is not allowed for RAF
        # VBA: If Range(e) = "-" Then MsgBox "DH must be input for RAF!"
        if proc.get("dh_a") == HYPHEN:
            raise ValueError(
                f"Runway {proc.get('runway_id','?')}: "
                f"For RAF aerodromes, Decision Height value must be input!"
            )

        sys_min   = proc.get("system_minima", 0)
        datum     = proc.get("datum_type", "A")
        runway_id = proc.get("runway_id", "")
        gp_angle  = proc.get("glidepath_angle", 0)

        result = {"runway_id": runway_id}

        for cat, och_key, dh_key, raf_key in [
            ("cat_a", "och_a", "dh_a", "raf_dh_a"),
            ("cat_b", "och_b", "dh_b", "raf_dh_b"),
            ("cat_c", "och_c", "dh_c", "raf_dh_c"),
            ("cat_d", "och_d", "dh_d", "raf_dh_d"),
        ]:
            result[cat] = _process_category(
                och             = proc.get(och_key),
                published_dh    = proc.get(dh_key),
                system_minima   = sys_min,
                datum_type      = datum,
                elev_aerodrome  = aerodrome_elev,
                elev_threshold  = threshold_elev,
                raf_dh          = proc.get(raf_key),
                glidepath_angle = gp_angle,
                is_raf          = True
            )

        results.append(result)

    return results
