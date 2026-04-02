# ============================================================
# raf_select.py
# RAF Aerodrome Selection Gate
# Replaces: Sub RAFSELECT() — VBA Module45
#
# Original VBA by: Alan Hutchinson, 17th January 2000
#
# PURPOSE:
#   Determines whether minima is calculated for a RAF
#   aerodrome or a civilian aerodrome, then routes to
#   the correct OCH/DH/DA calculation module accordingly.
#
# VBA LOGIC:
#   Reads Cell C11 → RAF flag (Y = RAF aerodrome)
#   Reads Cell D26 → Glidepath angle
#
#   If C11 = "Y" AND D26 has value → RAF route
#       → calls EXTOCHDHDARAF (Module43)
#   If C11 != "Y" AND D26 is empty → Non-RAF route
#       → calls EXTOCHDHDA (Module41)
#
#   Error conditions:
#   ERR1: C11="Y" but D26 is empty → GP angle missing
#   ERR2: D26 has value but C11 != "Y" → RAF flag missing
#
# HOW THIS CONNECTS TO THE FULL PIPELINE:
#   CheckValues (Module60) → RAFSELECT (Module45)
#                                  ↓              ↓
#                           EXTOCHDHDARAF    EXTOCHDHDA
#                           (RAF route)      (Non-RAF)
# ============================================================


# ----------------------------------------------------------
# RAF ROUTE CONSTANTS
# ----------------------------------------------------------
ROUTE_RAF    = "RAF"       # RAF aerodrome — use GP angle adjusted OCH
ROUTE_NONRAF = "NON_RAF"   # Standard civilian aerodrome
ROUTE_ERROR  = "ERROR"     # Invalid combination


def raf_select(raf_flag: str, glidepath_angle) -> dict:
    """
    Determine calculation route: RAF or Non-RAF.

    Replaces VBA logic in RAFSELECT():
        C11 = raf_flag     ("Y" or anything else)
        D26 = glidepath_angle (present or empty)

    Args:
        raf_flag:        Value from Sheet1 C11
                         "Y" = RAF aerodrome
                         anything else = Non-RAF

    glidepath_angle:     Value from Sheet1 D26
                         float = glidepath angle present (e.g. 3.0)
                         None / "" = not present

    Returns:
        dict:
        {
            "route":    str     "RAF" | "NON_RAF" | "ERROR"
            "valid":    bool
            "message":  str     status / error description
            "next_step": str    which calculation to run next
        }
    """

    is_raf        = str(raf_flag).strip().upper() == "Y" if raf_flag else False
    has_glidepath = glidepath_angle not in (None, "", 0)

    # ── RAF AERODROME ──────────────────────────────────────
    # VBA: If Range("C11") = "Y" Then
    #          If Range("D26") <> "" Then GoTo RAF
    #          Else GoTo ERR1
    if is_raf:
        if has_glidepath:
            return {
                "route":     ROUTE_RAF,
                "valid":     True,
                "message":   "RAF Minima will be calculated.",
                "next_step": "EXTOCHDHDARAF",   # Module43
            }
        else:
            # ERR1: GP angle missing for RAF
            return {
                "route":     ROUTE_ERROR,
                "valid":     False,
                "message":   "GlidePath angle MUST be present to calculate RAF minima.",
                "next_step": None,
            }

    # ── NON-RAF AERODROME ──────────────────────────────────
    # VBA: Else
    #          If Range("D26") = "" Then GoTo NONRAF
    #          Else GoTo ERR2
    else:
        if not has_glidepath:
            return {
                "route":     ROUTE_NONRAF,
                "valid":     True,
                "message":   "Standard (Non-RAF) Minima will be calculated.",
                "next_step": "EXTOCHDHDA",      # Module41
            }
        else:
            # ERR2: GP angle present but C11 not set to Y
            return {
                "route":     ROUTE_ERROR,
                "valid":     False,
                "message":   "Cell C11 MUST be equal to Y to calculate RAF minima.",
                "next_step": None,
            }
