# ============================================================
# output_generator.py
# Output Generator — produces results matching Excel tool exactly
#
# Verified against: EGXE_PRAC.csv (Exeter Airport reference output)
#
# EGXE_PRAC.csv analysis:
#   - Aerodrome: EGXE (Exeter Airport, UK)
#   - Output type: TAKE-OFF MINIMA page
#   - Publisher: European Aeronautical Group Aerad
#   - Format: JAR-OPS1 compliant
#
# TAKE-OFF TABLE (from EGXE_PRAC.csv):
#   Facility                      Cat A    Cat B
#   RCLL(H)+REDL(H)+Multi RVR      -        -
#   RCLL+REDL+Multi RVR            -        -
#   RCLL+REDL                      -        -
#   RCL and/or REDL (LM)          250      250   ← LM code
#   Nil (Day only)   (N)           500      500   ← N  code
#
# VALUES VERIFIED:
#   LM code Cat A = 250m ✅   LM code Cat B = 250m ✅
#   N  code Cat A = 500m ✅   N  code Cat B = 500m ✅
#
# MODULES THIS REPLACES:
#   Module18 → JARTKOFF         (JAR take-off format)
#   Module38 → TKOFFRWYS        (runway identifiers)
#   Module21 → PlaceTkoffRVRAB  (Cat A/B RVR placement)
#   Module30 → PlaceTkoffRVRCD  (Cat C/D RVR placement)
# ============================================================

from rvr_tables import TAKEOFF_MINIMA, get_takeoff_rvr


# ── TAKE-OFF FACILITY ROWS — matches Excel template exactly ──
# Source: Module21 (PlaceTkoffRVRAB) and Module30 (PlaceTkoffRVRCD)
# Order matches VBA template layout

TAKEOFF_FACILITY_ROWS = [
    {
        "code":        "HLR",
        "description": "RCLL(H)+REDL(H)+Multi RVR",
        "note":        "(1)",
    },
    {
        "code":        "LR",
        "description": "RCLL+REDL+Multi RVR",
        "note":        "",
    },
    {
        "code":        "L",
        "description": "RCLL+REDL",
        "note":        "",
    },
    {
        "code":        "LM",
        "description": "RCL and/or REDL",
        "note":        "(2)",
    },
    {
        "code":        "N",
        "description": "Nil (Day only)",
        "note":        "",
    },
]

# Standard JAR-OPS1 notes (always present)
STANDARD_NOTES = [
    "(1) Subject to Approval, see explanatory notes in the Flight Information Supplement.",
    "(2) For night operations, at least runway edge and end lights required.",
]

PUBLISHER = "European Aeronautical Group Aerad"
COMPLIANCE = (
    "The following Minima is for Public Transport aircraft and conforms to JAR-OPS1 "
    "regulations. See explanatory notes in Flight Information Supplement."
)


def build_takeoff_table(runways: list, active_codes: list = None) -> list:
    """
    Build take-off minima table matching Excel output exactly.

    For each facility row:
    - If facility code is in active_codes AND runway has this facility:
        → look up RVR value from TAKEOFF_MINIMA
    - Otherwise:
        → place "-" (not applicable)

    This mirrors VBA Module21/30 which placed either the RVR value
    or "-" depending on what facilities existed at the aerodrome.

    Args:
        runways:      list of runway identifiers e.g. ["16", "34"]
        active_codes: list of facility codes active at this aerodrome
                      e.g. ["LM", "N"]
                      If None, uses all codes from TAKEOFF_MINIMA

    Returns:
        list of row dicts matching EGXE_PRAC.csv structure
    """
    if active_codes is None:
        active_codes = list(TAKEOFF_MINIMA.keys())

    runway_str = ", ".join(runways) if runways else "All"
    rows = []

    for fac in TAKEOFF_FACILITY_ROWS:
        code = fac["code"]
        desc = fac["description"]
        note = fac["note"]

        if code in active_codes:
            val_a = get_takeoff_rvr(code, "a")
            val_b = get_takeoff_rvr(code, "b")
            val_c = get_takeoff_rvr(code, "c")
            val_d = get_takeoff_rvr(code, "d")
            rwy   = runway_str
        else:
            val_a = val_b = val_c = val_d = "-"
            rwy   = ""

        rows.append({
            "runway":      rwy,
            "facility":    desc + (" " + note if note else ""),
            "facility_code": code,
            "cat_a":       val_a,
            "cat_b":       val_b,
            "cat_c":       val_c,
            "cat_d":       val_d,
            "applicable":  code in active_codes,
        })

    return rows


def generate_output(
    icao:         str,
    airport_name: str,
    date:         str,
    runways:      list,
    active_takeoff_codes: list,
    approach_results:     list = None,
    heli_results:         list = None,
    dh_da_results:        list = None,
) -> dict:
    """
    Generate complete tool output matching Excel format.

    Args:
        icao:                 e.g. "EGXE"
        airport_name:         e.g. "Exeter"
        date:                 e.g. "2026-03-31"
        runways:              e.g. ["16", "34"]
        active_takeoff_codes: e.g. ["LM", "N"]
        approach_results:     from fixed_wing_rvr.run_fixed_wing_rvr()
        heli_results:         from heli_rvr_calc.heli_rvr_calculation()
        dh_da_results:        from ext_och_dh_da.extoch_dh_da()

    Returns:
        Complete output dict with all sections
    """
    takeoff_table = build_takeoff_table(runways, active_takeoff_codes)

    return {
        "header": {
            "title":       "AERODROME OPERATING MINIMA",
            "compliance":  COMPLIANCE,
            "publisher":   PUBLISHER,
            "icao":        icao.upper(),
            "airport":     airport_name,
            "date":        date,
        },
        "takeoff": {
            "runways":  runways,
            "table":    takeoff_table,
            "notes":    STANDARD_NOTES,
        },
        "approach":  approach_results or [],
        "heli":      heli_results     or [],
        "dh_da":     dh_da_results    or [],
    }


def format_takeoff_text(output: dict) -> str:
    """
    Format take-off section as plain text — matches EGXE_PRAC.csv layout.
    """
    hdr  = output["header"]
    tkof = output["takeoff"]

    lines = []
    lines.append(f" {hdr['title']}")
    lines.append("")
    lines.append(f"  {hdr['compliance']}")
    lines.append("")
    lines.append(f"{hdr['icao']}")
    lines.append("")
    lines.append(f"  {'TAKE-OFF':<20}{'A':>8}{'B':>8}{'C':>8}{'D':>8}")
    lines.append(f"  {'Runway':<20}{'Facilities':<35}{'m':>6}{'m':>6}{'m':>6}{'m':>6}")

    for row in tkof["table"]:
        rwy  = row["runway"] or ""
        fac  = row["facility"]
        a    = str(row["cat_a"])
        b    = str(row["cat_b"])
        c    = str(row["cat_c"])
        d    = str(row["cat_d"])
        lines.append(f"  {rwy:<20}{fac:<35}{a:>6}{b:>6}{c:>6}{d:>6}")

    lines.append("")
    lines.append(f"  {'Notes:'}")
    for note in tkof["notes"]:
        lines.append(f"  {note}")
    lines.append("")
    lines.append(f"{hdr['publisher']}")

    return "\n".join(lines)


def verify_against_egxe(output: dict) -> dict:
    """
    Verify output matches EGXE_PRAC.csv reference exactly.

    Returns:
        { "passed": bool, "checks": list of check results }
    """
    checks = []
    tkof   = output["takeoff"]

    # Build lookup by code
    rows_by_code = {r["facility_code"]: r for r in tkof["table"]}

    # Expected values from EGXE_PRAC.csv
    expected = [
        ("HLR", "cat_a", "-",   "RCLL(H)+REDL(H)+Multi RVR Cat A"),
        ("HLR", "cat_b", "-",   "RCLL(H)+REDL(H)+Multi RVR Cat B"),
        ("LR",  "cat_a", "-",   "RCLL+REDL+Multi RVR Cat A"),
        ("LR",  "cat_b", "-",   "RCLL+REDL+Multi RVR Cat B"),
        ("L",   "cat_a", "-",   "RCLL+REDL Cat A"),
        ("L",   "cat_b", "-",   "RCLL+REDL Cat B"),
        ("LM",  "cat_a", 250,   "RCL and/or REDL Cat A"),
        ("LM",  "cat_b", 250,   "RCL and/or REDL Cat B"),
        ("N",   "cat_a", 500,   "Nil (Day only) Cat A"),
        ("N",   "cat_b", 500,   "Nil (Day only) Cat B"),
    ]

    all_pass = True
    for code, field, exp, label in expected:
        row = rows_by_code.get(code, {})
        got = row.get(field)
        # Convert for comparison
        got_cmp = int(got) if str(got).isdigit() else got
        exp_cmp = exp
        passed  = (got_cmp == exp_cmp)
        if not passed:
            all_pass = False
        checks.append({
            "label":   label,
            "expected": exp,
            "got":      got,
            "passed":   passed,
        })

    return {"passed": all_pass, "checks": checks}
