# ============================================================
# rvr_tables.py
# Complete RVR Lookup Tables — source: V:Tables.xls / TABLES.csv
#
# STRUCTURE OF TABLES.csv:
# ─────────────────────────────────────────────────────────────
# SECTION 1: Reference Data (Rows 3-33)
#   - System Minima by procedure type
#   - Lights classification
#   - Procedure type codes
#   - Sort order (SORTPROC)
#   - RAF Incremental Add-ons by GP angle
#   - Take-Off Minima codes
#   - HELI/CAT1 Facilities tables
#   - Procedure code → RVR Table name mapping
#
# SECTION 2: HELICOPTER RVR TABLES (Rows 38-74)
#   Precision:     TABLE 120 (Full), 220 (Int), 320 (Basic), 420 (None)
#   Non-Precision: TABLE 110 (Full), 210 (Int), 310 (Basic), 410 (None)
#   Input:  DH (Decision Height) ft
#   Output: RVR (m) — single value, same for all Cats A/B/C/D
#
# SECTION 3: CAT 1 RVR TABLES (Rows 81-117)
#   Precision:     TABLE 520 (Full), 620 (Int), 720 (Basic), 820 (None)
#   Non-Precision: TABLE 510 (Full), 610 (Int), 710 (Basic), 810 (None)
#   Input:  MDH (Minimum Descent Height) ft
#   Output: RVR (m) — per Cat A, B, C, D separately
#
# SECTION 4: Additional Reference (Rows 3-37)
#   - System Minima values
#   - RAF GP angle increments
#   - Take-Off Minima codes
#   - Procedure Code → Table mapping
#   - HELI/CAT1 Facilities
#   - Helicopter Circling minima
# ============================================================


# ──────────────────────────────────────────────────────────────
# SECTION 1: SYSTEM MINIMA
# Source: Rows 4-17, cols B (procedure) and C (system minima ft)
# ──────────────────────────────────────────────────────────────
SYSTEM_MINIMA = {
    "ILS":      200,
    "L":        300,
    "LLZ":      250,
    "LLZ/BB":   250,
    "MLS":      200,
    "NDB":      300,
    "PAR":      200,
    "PAR/NGP":  250,
    "SRA":      350,
    "SRA 0.5nm":250,
    "SRA 1nm":  300,
    "SRA 2nm":  350,
    "VDF":      300,
    "VOR":      300,
    "VOR/DME":  250,
}


# ──────────────────────────────────────────────────────────────
# SECTION 2: LIGHTS CLASSIFICATION
# Source: Rows 4-14, col D (code) and col E (lights type)
# ──────────────────────────────────────────────────────────────
LIGHTS_CLASSIFICATION = {
    0:    "None",
    1:    "Basic",
    419:  "Basic",
    420:  "Int",
    719:  "Int",
    720:  "Full",
    2500: "Full",
    "L":  "Basic",
}


# ──────────────────────────────────────────────────────────────
# SECTION 3: PROCEDURE TYPE — Precision/Non-Precision
# Source: Rows 4-14, cols G (procedure) and H (P/N)
# P = Precision, N = Non-Precision
# ──────────────────────────────────────────────────────────────
PROC_TYPE = {
    "ILS":      "P",
    "PAR":      "P",
    "PAR/NGP":  "P",
    "MLS":      "P",
    "LLZ":      "N",
    "SRA":      "N",
    "VOR":      "N",
    "VOR/DME":  "N",
    "NDB":      "N",
    "VDF":      "N",
    "IN D. APP.":"N",
}


# ──────────────────────────────────────────────────────────────
# SECTION 4: SORT PROCEDURE ORDER (SORTPROC)
# Source: Rows 4-25, cols J (procedure) and K (sort order)
# ──────────────────────────────────────────────────────────────
SORT_PROC = {
    "ILS":      1,
    "LLZ":      2,
    "MLS":      3,
    "PAR":      4,
    "PAR/ NGP": 5,
    "LLZ/BB":   6,
    "VOR/DME":  7,
    "VOR":      8,
    "SRA 0.5nm":9,
    "SRA 1nm":  10,
    "SRA 1.5nm":11,
    "SRA 2nm":  12,
    "SRA 2.5nm":13,
    "SRA 3nm":  14,
    "SRA 3.5nm":15,
    "SRA 4nm":  16,
    "SRA 4.5nm":17,
    "SRA 5nm":  18,
    "SRA":      19,
    "NDB":      20,
    "L":        21,
    "VDF":      22,
}


# ──────────────────────────────────────────────────────────────
# SECTION 5: RAF GP ANGLE INCREMENTAL ADD-ONS
# Source: Rows 5-14, cols M (GP angle), N/O/P/Q (Cat A/B/C/D)
# Used by EXTOCHDHDARAF (Module43) to add RAF increment to DH
# ──────────────────────────────────────────────────────────────
RAF_INCREMENTS = {
    # gp_angle: {cat_a, cat_b, cat_c, cat_d}
    0.0:  {"a": 0,  "b": 0,  "c": 0,  "d": 0},
    2.4:  {"a": 0,  "b": 0,  "c": 0,  "d": 0},
    2.5:  {"a": 0,  "b": 10, "c": 20, "d": 30},
    2.6:  {"a": 10, "b": 20, "c": 30, "d": 40},
    2.7:  {"a": 10, "b": 20, "c": 30, "d": 40},
    2.8:  {"a": 20, "b": 30, "c": 40, "d": 50},
    2.9:  {"a": 20, "b": 30, "c": 40, "d": 50},
    3.0:  {"a": 30, "b": 40, "c": 50, "d": 60},
    3.1:  {"a": 0,  "b": 0,  "c": 0,  "d": 0},
    99.0: {"a": 0,  "b": 0,  "c": 0,  "d": 0},
}

def get_raf_increment(gp_angle: float, cat: str) -> int:
    """
    Get RAF DH increment for a given GP angle and category.
    Finds the closest GP angle entry.
    cat: 'a', 'b', 'c', or 'd'
    """
    if gp_angle is None or gp_angle == 0:
        return 0
    # Find closest GP angle key
    keys = sorted(RAF_INCREMENTS.keys())
    matched = None
    for k in keys:
        if gp_angle >= k:
            matched = k
    if matched is not None:
        return RAF_INCREMENTS[matched].get(cat.lower(), 0)
    return 0


# ──────────────────────────────────────────────────────────────
# SECTION 6: TAKE-OFF MINIMA
# Source: Rows 18-37, cols Q (code), R (description), S/T/U/V (Cat A/B/C/D)
# Code → RVR values per category
# ──────────────────────────────────────────────────────────────
TAKEOFF_MINIMA = {
    # Code:  description,                Cat A, Cat B, Cat C, Cat D
    "HLR": {"desc": "HRCL + HRL + Multi RVR", "a": 125, "b": 125, "c": 125, "d": 150},
    "LR":  {"desc": "RCL + RL + Multi RVR",   "a": 150, "b": 150, "c": 150, "d": 200},
    "L":   {"desc": "RCL + RL",               "a": 200, "b": 200, "c": 200, "d": 250},
    "LM":  {"desc": "RL + RCLM",              "a": 250, "b": 250, "c": 250, "d": 300},
    "N":   {"desc": "Nil (No RCL,RL,RCLM)",   "a": 500, "b": 500, "c": 500, "d": 500},
}

def get_takeoff_rvr(code: str, cat: str) -> int:
    """
    Get take-off RVR for a given code and category.
    code: 'HLR', 'LR', 'L', 'LM', 'N'
    cat:  'a', 'b', 'c', 'd'
    """
    entry = TAKEOFF_MINIMA.get(str(code).strip().upper())
    if entry is None:
        return None
    return entry.get(cat.lower())


# ──────────────────────────────────────────────────────────────
# SECTION 7: PROCEDURE CODE → RVR TABLE MAPPING
# Source: Rows 18-33, cols M (proc code) and N (table name)
# Used to cross-reference which RVR table applies to each proc code
# ──────────────────────────────────────────────────────────────
PROC_CODE_TABLE_MAP = {
    # HELI tables
    110: "NPRVRF",
    120: "RVRPF",
    210: "NPRVRI",
    220: "RVRPI",
    310: "NPRVRB",
    320: "RVRPB",
    410: "NPRVRN",
    420: "RVRPN",
    # CAT 1 tables
    510: "RVRNPF",
    520: "PRVRF",
    610: "RVRNPI",
    620: "PRVRI",
    710: "RVRNPB",
    720: "PRVRB",
    810: "RVRNPN",
    820: "PRVRN",
}


# ──────────────────────────────────────────────────────────────
# SECTION 8: HELI FACILITIES TABLE
# Source: Rows 25-30, cols B/C (HELI) and D/E (CAT 1)
# ──────────────────────────────────────────────────────────────
HELI_FACILITIES = {
    "Basic": 300,
    "Full":  100,
    "Int":   200,
    "None":  400,
}

CAT1_FACILITIES = {
    "Basic": 700,
    "Full":  500,
    "Int":   600,
    "None":  800,
}

NP_CHECK_TABLE = {
    "N": 10,   # Non-Precision
    "P": 20,   # Precision
}


# ──────────────────────────────────────────────────────────────
# SECTION 9: HELICOPTER CIRCLING MINIMA
# Source: Row 22, cols B (label), C (MDH), D (visibility)
# ──────────────────────────────────────────────────────────────
HELI_CIRCLING = {
    "mdh": 250,
    "visibility": 800,
}

# Cat1 OCH Circling (Rows 19-22)
CAT1_OCH_CIRCLING = {
    "A": {"mdh": 400, "mdh_alt": 400, "vis": 1500},
    "B": {"mdh": 500, "mdh_alt": 500, "vis": 1600},
    "C": {"mdh": 600, "mdh_alt": 600, "vis": 2400},
    "D": {"mdh": 700, "mdh_alt": 700, "vis": 3600},
}


# ══════════════════════════════════════════════════════════════
# HELICOPTER RVR TABLES (Rows 38-74)
# TABLE 120 = RVRPF  — Precision Full Lights
# TABLE 220 = RVRPI  — Precision Intermediate Lights
# TABLE 320 = RVRPB  — Precision Basic Lights
# TABLE 420 = RVRPN  — Precision No Lights
# TABLE 110 = NPRVRF — Non-Precision Full Lights
# TABLE 210 = NPRVRI — Non-Precision Intermediate Lights
# TABLE 310 = NPRVRB — Non-Precision Basic Lights
# TABLE 410 = NPRVRN — Non-Precision No Lights
#
# Format: {DH_threshold: RVR_metres}
# VLookup: find largest key <= input DH
# ══════════════════════════════════════════════════════════════

# ── PRECISION HELI TABLES ──────────────────────────────────────
# Source: Rows 41-47, cols C-D (Full), F-G (Int), I-J (Basic), L-M (None)

RVRPF = {   # TABLE 120 — Precision Full Lights (HELI)
    200:   500,
    201:   550,
    250:   550,
    251:   600,
    300:   600,
    301:   750,
    10000: 750,
}

RVRPI = {   # TABLE 220 — Precision Intermediate Lights (HELI)
    200:   600,
    201:   650,
    250:   650,
    251:   700,
    300:   700,
    301:   800,
    10000: 800,
}

RVRPB = {   # TABLE 320 — Precision Basic Lights (HELI)
    200:   700,
    201:   750,
    250:   750,
    251:   800,
    300:   800,
    301:   900,
    10000: 900,
}

RVRPN = {   # TABLE 420 — Precision No Lights (HELI)
    200:   1000,
    201:   1000,
    250:   1000,
    251:   1000,
    300:   1000,
    301:   1000,
    10000: 1000,
}

# ── NON-PRECISION HELI TABLES ─────────────────────────────────
# Source: Rows 54-61 (FULL/INT), Rows 66-73 (BASIC/NONE)
# NOTE: Non-Precision HELI tables have only Cat A RVR
# (single column — same value applies regardless of aircraft cat)

NPRVRF = {  # TABLE 110 — Non-Precision Full Lights (HELI)
    250:  600,
    299:  600,
    300:  800,
    449:  800,
    450:  1000,
    649:  1000,
    650:  1000,
    2000: 1000,
}

NPRVRI = {  # TABLE 210 — Non-Precision Intermediate Lights (HELI)
    250:  800,
    299:  800,
    300:  1000,
    449:  1000,
    450:  1000,
    649:  1000,
    650:  1000,
    2000: 1000,
}

NPRVRB = {  # TABLE 310 — Non-Precision Basic Lights (HELI)
    250:  1000,
    299:  1000,
    300:  1000,
    449:  1000,
    450:  1000,
    649:  1000,
    650:  1000,
    2000: 1000,
}

NPRVRN = {  # TABLE 410 — Non-Precision No Lights (HELI)
    250:  1000,
    299:  1000,
    300:  1000,
    449:  1000,
    450:  1000,
    649:  1000,
    650:  1000,
    2000: 1000,
}


# ══════════════════════════════════════════════════════════════
# CAT 1 RVR TABLES (Rows 81-117)
# TABLE 520 = PRVRF  — Precision Full Lights     (Cat 1)
# TABLE 620 = PRVRI  — Precision Intermediate    (Cat 1)
# TABLE 720 = PRVRB  — Precision Basic           (Cat 1)
# TABLE 820 = PRVRN  — Precision No Lights       (Cat 1)
# TABLE 510 = RVRNPF — Non-Precision Full        (Cat 1)
# TABLE 610 = RVRNPI — Non-Precision Intermediate(Cat 1)
# TABLE 710 = RVRNPB — Non-Precision Basic       (Cat 1)
# TABLE 810 = RVRNPN — Non-Precision No Lights   (Cat 1)
#
# Format: {MDH_threshold: {cat_a, cat_b, cat_c, cat_d}}
# NOTE: CAT 1 NP tables have per-category RVR (A/B/C/D)
# ══════════════════════════════════════════════════════════════

# ── PRECISION CAT 1 TABLES ────────────────────────────────────
# Source: Rows 84-90, same DH→RVR single value structure as HELI

PRVRF = {   # TABLE 520 — Precision Full Lights (CAT 1)
    200:   550,
    201:   600,
    250:   600,
    251:   650,
    300:   650,
    301:   800,
    10000: 800,
}

PRVRI = {   # TABLE 620 — Precision Intermediate (CAT 1)
    200:   700,
    201:   700,
    250:   700,
    251:   800,
    300:   800,
    301:   900,
    10000: 900,
}

PRVRB = {   # TABLE 720 — Precision Basic (CAT 1)
    200:   800,
    201:   800,
    250:   800,
    251:   900,
    300:   900,
    301:   1000,
    10000: 1000,
}

PRVRN = {   # TABLE 820 — Precision No Lights (CAT 1)
    200:   1000,
    201:   1000,
    250:   1000,
    251:   1200,
    300:   1200,
    301:   1200,
    10000: 1200,
}

# ── NON-PRECISION CAT 1 TABLES ────────────────────────────────
# Source: Rows 97-104 (FULL/INT), Rows 109-116 (BASIC/NONE)
# These have per-category A/B/C/D RVR values

RVRNPF = {  # TABLE 510 — Non-Precision Full (CAT 1)
    250:  {"a": 800,  "b": 800,  "c": 800,  "d": 1200},
    299:  {"a": 800,  "b": 800,  "c": 800,  "d": 1200},
    300:  {"a": 900,  "b": 1000, "c": 1000, "d": 1400},
    449:  {"a": 900,  "b": 1000, "c": 1000, "d": 1400},
    450:  {"a": 1000, "b": 1200, "c": 1200, "d": 1600},
    649:  {"a": 1000, "b": 1200, "c": 1200, "d": 1600},
    650:  {"a": 1200, "b": 1400, "c": 1400, "d": 1800},
    2000: {"a": 1200, "b": 1400, "c": 1400, "d": 1800},
}

RVRNPI = {  # TABLE 610 — Non-Precision Intermediate (CAT 1)
    250:  {"a": 1000, "b": 1100, "c": 1200, "d": 1400},
    299:  {"a": 1000, "b": 1100, "c": 1200, "d": 1400},
    300:  {"a": 1200, "b": 1300, "c": 1400, "d": 1600},
    449:  {"a": 1200, "b": 1300, "c": 1400, "d": 1600},
    450:  {"a": 1400, "b": 1500, "c": 1600, "d": 1800},
    649:  {"a": 1400, "b": 1500, "c": 1600, "d": 1800},
    650:  {"a": 1500, "b": 1500, "c": 1800, "d": 2000},
    2000: {"a": 1500, "b": 1500, "c": 1800, "d": 2000},
}

RVRNPB = {  # TABLE 710 — Non-Precision Basic (CAT 1)
    250:  {"a": 1200, "b": 1300, "c": 1400, "d": 1600},
    299:  {"a": 1200, "b": 1300, "c": 1400, "d": 1600},
    300:  {"a": 1300, "b": 1400, "c": 1600, "d": 1800},
    449:  {"a": 1300, "b": 1400, "c": 1600, "d": 1800},
    450:  {"a": 1500, "b": 1500, "c": 1800, "d": 2000},
    649:  {"a": 1500, "b": 1500, "c": 1800, "d": 2000},
    650:  {"a": 1500, "b": 1500, "c": 2000, "d": 2000},
    2000: {"a": 1500, "b": 1500, "c": 2000, "d": 2000},
}

RVRNPN = {  # TABLE 810 — Non-Precision No Lights (CAT 1)
    250:  {"a": 1500, "b": 1500, "c": 1600, "d": 1800},
    299:  {"a": 1500, "b": 1500, "c": 1600, "d": 1800},
    300:  {"a": 1500, "b": 1500, "c": 1800, "d": 2000},
    449:  {"a": 1500, "b": 1500, "c": 1800, "d": 2000},
    450:  {"a": 1500, "b": 1500, "c": 2000, "d": 2000},
    649:  {"a": 1500, "b": 1500, "c": 2000, "d": 2000},
    650:  {"a": 1500, "b": 1500, "c": 2000, "d": 2000},
    2000: {"a": 1500, "b": 1500, "c": 2000, "d": 2000},
}


# ══════════════════════════════════════════════════════════════
# ALL TABLES REGISTRY
# ══════════════════════════════════════════════════════════════

# HELI tables — single RVR value (no per-cat differentiation)
HELI_TABLES = {
    "RVRPF":  RVRPF,
    "RVRPI":  RVRPI,
    "RVRPB":  RVRPB,
    "RVRPN":  RVRPN,
    "NPRVRF": NPRVRF,
    "NPRVRI": NPRVRI,
    "NPRVRB": NPRVRB,
    "NPRVRN": NPRVRN,
}

# CAT 1 Precision tables — single value
CAT1_P_TABLES = {
    "PRVRF": PRVRF,
    "PRVRI": PRVRI,
    "PRVRB": PRVRB,
    "PRVRN": PRVRN,
}

# CAT 1 Non-Precision tables — per-cat A/B/C/D values
CAT1_NP_TABLES = {
    "RVRNPF": RVRNPF,
    "RVRNPI": RVRNPI,
    "RVRNPB": RVRNPB,
    "RVRNPN": RVRNPN,
}

ALL_TABLES = {**HELI_TABLES, **CAT1_P_TABLES, **CAT1_NP_TABLES}


# ══════════════════════════════════════════════════════════════
# VLOOKUP FUNCTIONS
# Replaces: VBA Vlookup(dh, 'V:Tables.xls'!TABLE_NAME, 2)
# Logic: find largest key <= input value (approximate match)
# ══════════════════════════════════════════════════════════════

def vlookup_rvr(dh_value, table_name: str, cat: str = "a"):
    """
    Look up RVR from any table based on DH/MDH input.

    For HELI tables (120-420, 110-410):
        Returns single int — cat param ignored

    For CAT1 NP tables (510-810):
        Returns RVR for specific category (a/b/c/d)

    For CAT1 Precision tables (520-820):
        Returns single int — cat param ignored

    Args:
        dh_value:   Decision/Minimum Height (float or "N/A")
        table_name: e.g. "RVRPF", "RVRNPF", "PRVRF"
        cat:        Category "a","b","c","d" (for NP CAT1 tables)

    Returns:
        RVR in metres (int) or "N/A"
    """
    if dh_value == "N/A" or dh_value is None:
        return "N/A"

    table = ALL_TABLES.get(table_name)
    if table is None:
        return "N/A"

    try:
        dh_num = float(dh_value)
        sorted_keys = sorted(table.keys())

        # Find largest key <= dh_num (VLookup approximate match)
        matched_key = None
        for key in sorted_keys:
            if dh_num >= key:
                matched_key = key

        if matched_key is None:
            return "N/A"

        value = table[matched_key]

        # CAT 1 NP tables return dict — extract per-cat value
        if isinstance(value, dict):
            return value.get(cat.lower(), "N/A")

        return value

    except (ValueError, TypeError):
        return "N/A"


def vlookup_heli_rvr(dh_value, table_name: str):
    """Convenience wrapper for HELI single-value lookup."""
    return vlookup_rvr(dh_value, table_name, cat="a")


def vlookup_cat1_rvr(mdh_value, table_name: str, cat: str):
    """Convenience wrapper for CAT1 per-category lookup."""
    return vlookup_rvr(mdh_value, table_name, cat=cat)


def get_table_name_for_proc_code(proc_code: int) -> str:
    """
    Get the RVR table name for a given procedure code.
    e.g. 120 → 'RVRPF', 510 → 'RVRNPF'
    """
    return PROC_CODE_TABLE_MAP.get(proc_code, "N/A")
