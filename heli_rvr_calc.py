# ============================================================
# heli_rvr_calc.py
# Core RVR Calculation Engine for Helicopter Minima
# Replaces: Sub HELIRVRCALCULATION() — VBA Module51
#
# Original VBA by: Alan Hutchinson, 11th June 2001
# Python conversion: Automated migration from Excel VBA
# ============================================================

from rvr_tables import vlookup_rvr

# ----------------------------------------------------------
# PROCEDURE TYPE CODE MAP
# Hundreds digit = Approach Type
# Units digit    = Lighting Type (20=Precision, 10=Non-Precision)
#
# Code | Approach Type        | Lighting  | Tables Used
# -----|----------------------|-----------|---------------------------
#  420 | Precision            | None      | RVRPN only
#  410 | Non-Precision        | None      | NPRVRN only
#  320 | Precision            | Basic     | RVRPB + RVRPN
#  310 | Non-Precision        | Basic     | NPRVRB + NPRVRN
#  220 | Precision            | Inter     | RVRPI + RVRPB + RVRPN
#  210 | Non-Precision        | Inter     | NPRVRI + NPRVRB + NPRVRN
#  120 | Precision            | Full      | RVRPF + RVRPI + RVRPB + RVRPN
#  110 | Non-Precision        | Full      | NPRVRF + NPRVRI + NPRVRB + NPRVRN
# ----------------------------------------------------------

PROC_TYPE_MAP = {
    420: {
        "name":  "Precision — No Lights",
        "full":  None,      # CE column → N/A
        "inter": None,      # CF column → N/A
        "basic": None,      # CG column → N/A
        "none":  "RVRPN",   # CH column → lookup
    },
    410: {
        "name":  "Non-Precision — No Lights",
        "full":  None,
        "inter": None,
        "basic": None,
        "none":  "NPRVRN",
    },
    320: {
        "name":  "Precision — Basic Lights",
        "full":  None,
        "inter": None,
        "basic": "RVRPB",
        "none":  "RVRPN",
    },
    310: {
        "name":  "Non-Precision — Basic Lights",
        "full":  None,
        "inter": None,
        "basic": "NPRVRB",
        "none":  "NPRVRN",
    },
    220: {
        "name":  "Precision — Intermediate Lights",
        "full":  None,
        "inter": "RVRPI",
        "basic": "RVRPB",
        "none":  "RVRPN",
    },
    210: {
        "name":  "Non-Precision — Intermediate Lights",
        "full":  None,
        "inter": "NPRVRI",
        "basic": "NPRVRB",
        "none":  "NPRVRN",
    },
    120: {
        "name":  "Precision — Full Lights",
        "full":  "RVRPF",
        "inter": "RVRPI",
        "basic": "RVRPB",
        "none":  "RVRPN",
    },
    110: {
        "name":  "Non-Precision — Full Lights",
        "full":  "NPRVRF",
        "inter": "NPRVRI",
        "basic": "NPRVRB",
        "none":  "NPRVRN",
    },
}


# ----------------------------------------------------------
# MAIN CALCULATION FUNCTION
# Replaces: Sub HELIRVRCALCULATION()
# ----------------------------------------------------------
def heli_rvr_calculation(runways: list) -> list:
    """
    Calculate RVR values for all helicopter runways.

    Equivalent of VBA Loop:
        Do Until x = 15
            t = CB column (proc type code)
            w = AG column (decision height)
            e = CE column (RVR Full output)
            h = CF column (RVR Inter output)
            k = CG column (RVR Basic output)
            n = CH column (RVR None output)

    Args:
        runways: list of dicts (max 15), each containing:
            {
                "runway_id":  str   e.g. "27L"
                "proc_type":  int   e.g. 420, 320, 110
                "dh":         float e.g. 250.0  (Decision Height)
            }

    Returns:
        list of result dicts, one per runway:
            {
                "runway_id":  str
                "proc_type":  int
                "proc_name":  str   human readable
                "dh":         float
                "rvr_full":   int or "N/A"   (CE column)
                "rvr_inter":  int or "N/A"   (CF column)
                "rvr_basic":  int or "N/A"   (CG column)
                "rvr_none":   int or "N/A"   (CH column)
            }
    """
    results = []

    # Loop max 15 runways — same as VBA: Do Until x = 15
    for runway in runways[:15]:

        proc_type = runway.get("proc_type", 0)
        dh        = runway.get("dh", "N/A")
        runway_id = runway.get("runway_id", "")

        # Skip if no procedure — same as VBA: If Range(t) = 0 Then GoTo NEXTRW
        if proc_type == 0:
            continue

        # Get table mapping for this procedure type
        proc = PROC_TYPE_MAP.get(proc_type)
        if proc is None:
            continue

        # Calculate RVR for each lighting category
        # Replaces VBA: MyText = "=If(" & w & "=""N/A"",""N/A"",Vlookup(...))"
        rvr_full  = vlookup_rvr(dh, proc["full"])  if proc["full"]  else "N/A"
        rvr_inter = vlookup_rvr(dh, proc["inter"]) if proc["inter"] else "N/A"
        rvr_basic = vlookup_rvr(dh, proc["basic"]) if proc["basic"] else "N/A"
        rvr_none  = vlookup_rvr(dh, proc["none"])  if proc["none"]  else "N/A"

        results.append({
            "runway_id":  runway_id,
            "proc_type":  proc_type,
            "proc_name":  proc["name"],
            "dh":         dh,
            "rvr_full":   rvr_full,    # CE column equivalent
            "rvr_inter":  rvr_inter,   # CF column equivalent
            "rvr_basic":  rvr_basic,   # CG column equivalent
            "rvr_none":   rvr_none,    # CH column equivalent
        })

    return results
