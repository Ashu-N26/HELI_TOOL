# ============================================================
# fixed_wing_rvr.py
# Fixed Wing RVR Calculation Engine
#
# Converts all 7 VBA modules exactly:
#   Module36 → rvr_calculation()        — Core CAT1 RVR lookup
#   Module42 → extoch_rvr()             — EXTOCH RVR comparison
#   Module27 → no_als_rvr_och()         — No ALS RVR for OCH
#   Module15 → no_als_rvr_extoch()      — No ALS RVR for EXTOCH
#   Module14 → rvr_compare()            — JAR ALS vs No ALS + State
#   Module17 → rvr_compare_extoch()     — EXTOCH JAR ALS vs No ALS + State
#   Module53 → rvr_vs_no_als()          — Final JAR vs State comparison
#
# COMPLETE CALL CHAIN (mirrors VBA exactly):
#
#   ext_och_dh_da (Module41/43)
#       ↓
#   rvr_calculation (Module36)   ← Core RVR from CAT1 tables
#       ↓ checks G26 (pub RVR)
#       ├─ G26 = "-"  → no_als_rvr_och (Module27)
#       │                   ↓ → rvr_compare (Module14) → END
#       └─ G26 has val → extoch_rvr (Module42)
#                           ↓ → no_als_rvr_extoch (Module15)
#                                   ↓ → rvr_compare_extoch (Module17)
#                                           ↓ → rvr_vs_no_als (Module53)
#
# KEY CELL MAPPINGS (per procedure row z, starting at z=26):
#   AC = Procedure type code          (820,810,720,710,620,610,520,510)
#   AG = Cat A derived DH             (from ext_och_dh_da)
#   AJ = Cat B derived DH
#   AM = Cat C derived DH
#   AP = Cat D derived DH
#   AH = Cat A OCH-based RVR          (intermediate)
#   AK = Cat B OCH-based RVR
#   AN = Cat C OCH-based RVR
#   AQ = Cat D OCH-based RVR
#   AR = Cat A No ALS RVR             (intermediate)
#   AS = Cat B No ALS RVR
#   AT = Cat C No ALS RVR
#   AU = Cat D No ALS RVR
#   AV = Cat A EXTOCH RVR (higher of pub vs OCH-based)
#   AW = Cat B EXTOCH RVR
#   AX = Cat C EXTOCH RVR
#   AY = Cat D EXTOCH RVR
#   G  = Cat A Published RVR          (user input from AIP/AIRAC)
#   J  = Cat B Published RVR
#   M  = Cat C Published RVR
#   P  = Cat D Published RVR
#   Q  = Cat A State No ALS RVR       (user input)
#   R  = Cat B State No ALS RVR
#   S  = Cat C State No ALS RVR
#   T  = Cat D State No ALS RVR
#
# OUTPUT cells (rows 46-60, one per procedure):
#   D = Cat A final ALS RVR
#   H = Cat B final ALS RVR
#   L = Cat C final ALS RVR
#   P = Cat D final ALS RVR
#   E = Cat A final No ALS RVR
#   I = Cat B final No ALS RVR
#   M = Cat C final No ALS RVR
#   Q = Cat D final No ALS RVR
# ============================================================

from rvr_tables import vlookup_rvr, vlookup_cat1_rvr

NA = "N/A"


def _parse(val):
    """Parse input value — returns float, NA, or None."""
    if val is None or str(val).strip() == "":
        return None
    s = str(val).strip()
    if s in (NA, "-", "N/A"):
        return NA
    try:
        return float(s)
    except ValueError:
        return s


def _higher(a, b):
    """Return the higher of two values. N/A treated as 0."""
    if a == NA and b == NA:
        return NA
    if a == NA:
        return b
    if b == NA:
        return a
    try:
        return max(float(a), float(b))
    except (ValueError, TypeError):
        return NA


# ── NO ALS TABLE MAPPING ─────────────────────────────────────
# Both Module27 and Module15 use the same No-ALS table logic:
# Precision tables (720,620,520) → always look up PRVRN (no lights)
# Non-Precision tables (710,610,510) → look up RVRNPN per cat
# 820/810 → N/A (no lights already placed — no ALS not needed)

NO_ALS_TABLE_MAP = {
    820: None,      # N/A — no lights already placed
    810: None,      # N/A — no lights already placed
    720: {"a": "PRVRN", "b": "PRVRN", "c": "PRVRN", "d": "PRVRN"},
    710: {"a": "RVRNPN","b": "RVRNPN","c": "RVRNPN","d": "RVRNPN"},
    620: {"a": "PRVRN", "b": "PRVRN", "c": "PRVRN", "d": "PRVRN"},
    610: {"a": "RVRNPN","b": "RVRNPN","c": "RVRNPN","d": "RVRNPN"},
    520: {"a": "PRVRN", "b": "PRVRN", "c": "PRVRN", "d": "PRVRN"},
    510: {"a": "RVRNPN","b": "RVRNPN","c": "RVRNPN","d": "RVRNPN"},
}

# ── ALS TABLE MAPPING ────────────────────────────────────────
# Module36: CAT1 RVR tables per procedure code
# Precision tables (820,720,620,520) → PRVRN/PRVRB/PRVRI/PRVRF
#   single value — same for all cats
# Non-Precision tables (810,710,610,510) → RVRNPN/RVRNPB/RVRNPI/RVRNPF
#   per-cat values (col 2=A, 3=B, 4=C, 5=D)

ALS_TABLE_MAP = {
    820: {"a": ("PRVRN","a"), "b": ("PRVRN","a"), "c": ("PRVRN","a"), "d": ("PRVRN","a")},
    810: {"a": ("RVRNPN","a"),"b": ("RVRNPN","b"),"c": ("RVRNPN","c"),"d": ("RVRNPN","d")},
    720: {"a": ("PRVRB","a"), "b": ("PRVRB","a"), "c": ("PRVRB","a"), "d": ("PRVRB","a")},
    710: {"a": ("RVRNPB","a"),"b": ("RVRNPB","b"),"c": ("RVRNPB","c"),"d": ("RVRNPB","d")},
    620: {"a": ("PRVRI","a"), "b": ("PRVRI","a"), "c": ("PRVRI","a"), "d": ("PRVRI","a")},
    610: {"a": ("RVRNPI","a"),"b": ("RVRNPI","b"),"c": ("RVRNPI","c"),"d": ("RVRNPI","d")},
    520: {"a": ("PRVRF","a"), "b": ("PRVRF","a"), "c": ("PRVRF","a"), "d": ("PRVRF","a")},
    510: {"a": ("RVRNPF","a"),"b": ("RVRNPF","b"),"c": ("RVRNPF","c"),"d": ("RVRNPF","d")},
}


def _lookup_als(dh, proc_code, cat):
    """Look up ALS RVR from CAT1 table for a given DH and category."""
    if dh == NA or dh is None:
        return NA
    mapping = ALS_TABLE_MAP.get(proc_code)
    if mapping is None:
        return NA
    table_name, col_cat = mapping[cat]
    return vlookup_rvr(dh, table_name, col_cat)


def _lookup_no_als(dh, proc_code, cat):
    """Look up No-ALS RVR from PRVRN/RVRNPN tables."""
    if dh == NA or dh is None:
        return NA
    mapping = NO_ALS_TABLE_MAP.get(proc_code)
    if mapping is None:
        return NA   # 820 or 810 — no lights already placed
    table_name = mapping[cat]
    return vlookup_rvr(dh, table_name, cat)


# ══════════════════════════════════════════════════════════════
# MODULE 36 — RVRCALCULATION
# Core CAT1 RVR lookup from tables for all procedure types
# Checks G26 (pub RVR) to route to EXTOCH or NoALS
# ══════════════════════════════════════════════════════════════
def rvr_calculation(procedures: list) -> dict:
    """
    Core CAT1 RVR calculation — replaces Sub RVRCALCULATION().

    For each procedure:
    1. Look up ALS RVR from CAT1 tables using derived DH (AG,AJ,AM,AP)
    2. Check if published RVR exists (G col) to determine route:
       - G = "-" or no pub RVR → route to no_als_rvr_och (NoALS path)
       - G has value           → route to extoch_rvr (EXTOCH path)

    Args:
        procedures: list of dicts (max 15), each:
        {
            "runway_id":    str
            "proc_code":    int     e.g. 820, 710, 520
            "dh_a":         float   Cat A derived DH (AG col)
            "dh_b":         float   Cat B derived DH (AJ col)
            "dh_c":         float   Cat C derived DH (AM col)
            "dh_d":         float   Cat D derived DH (AP col)
            "pub_rvr_a":    float or "-"   Cat A pub RVR (G col)
            "pub_rvr_b":    float or "-"   Cat B pub RVR (J col)
            "pub_rvr_c":    float or "-"   Cat C pub RVR (M col)
            "pub_rvr_d":    float or "-"   Cat D pub RVR (P col)
            "state_noals_a":float or ""    Cat A State No ALS (Q col)
            "state_noals_b":float or ""    Cat B State No ALS (R col)
            "state_noals_c":float or ""    Cat C State No ALS (S col)
            "state_noals_d":float or ""    Cat D State No ALS (T col)
        }

    Returns:
        {
            "route":      "EXTOCH" | "NO_ALS"
            "procedures": list of enriched procedure dicts with RVR values
        }
    """
    if not procedures:
        raise ValueError("Error - Procedure Type not entered correctly!")

    first = procedures[0]
    if not first.get("proc_code"):
        raise ValueError("Error - Procedure Type not entered correctly!")

    enriched = []

    for proc in procedures[:15]:
        code    = proc.get("proc_code", 0)
        dh_a    = _parse(proc.get("dh_a"))
        dh_b    = _parse(proc.get("dh_b"))
        dh_c    = _parse(proc.get("dh_c"))
        dh_d    = _parse(proc.get("dh_d"))

        # ALS RVR lookup — AH,AK,AN,AQ columns
        rvr_als_a = _lookup_als(dh_a, code, "a")
        rvr_als_b = _lookup_als(dh_b, code, "b")
        rvr_als_c = _lookup_als(dh_c, code, "c")
        rvr_als_d = _lookup_als(dh_d, code, "d")

        enriched.append({
            **proc,
            "rvr_als_a": rvr_als_a,   # AH col — OCH derived ALS RVR
            "rvr_als_b": rvr_als_b,   # AK col
            "rvr_als_c": rvr_als_c,   # AN col
            "rvr_als_d": rvr_als_d,   # AQ col
        })

    # Determine route based on G26 (pub_rvr_a of first procedure)
    # VBA: If Range("G26") = "-" → NoALSRVROCH else EXTOCHRVR
    first_pub = _parse(enriched[0].get("pub_rvr_a", "-"))
    route = "NO_ALS" if first_pub in (NA, None, "-") else "EXTOCH"

    return {"route": route, "procedures": enriched}


# ══════════════════════════════════════════════════════════════
# MODULE 42 — EXTOCHRVR
# Compare published RVR (from AIP) with OCH-derived RVR
# Take the HIGHER value for each cat
# ══════════════════════════════════════════════════════════════
def extoch_rvr(procedures: list) -> list:
    """
    EXTOCH RVR comparison — replaces Sub EXTOCHRVR().

    For each procedure:
        Cat A: MAX(pub_rvr_a [G], rvr_als_a [AH]) → AV col
        Cat B: MAX(pub_rvr_b [J], rvr_als_b [AK]) → AW col
        Cat C: MAX(pub_rvr_c [M], rvr_als_c [AN]) → AX col
        Cat D: MAX(pub_rvr_d [P], rvr_als_d [AQ]) → AY col

    VBA: =If(G26>AH26, G26, AH26)

    Args:
        procedures: list from rvr_calculation() output (enriched)

    Returns:
        list of procedures with extoch_rvr_a/b/c/d added
    """
    result = []
    for proc in procedures[:15]:
        pub_a = _parse(proc.get("pub_rvr_a"))
        pub_b = _parse(proc.get("pub_rvr_b"))
        pub_c = _parse(proc.get("pub_rvr_c"))
        pub_d = _parse(proc.get("pub_rvr_d"))

        als_a = proc.get("rvr_als_a", NA)
        als_b = proc.get("rvr_als_b", NA)
        als_c = proc.get("rvr_als_c", NA)
        als_d = proc.get("rvr_als_d", NA)

        # Stop if empty — no more procedures
        if pub_a is None and als_a in (NA, None):
            break

        result.append({
            **proc,
            "extoch_rvr_a": _higher(pub_a, als_a),  # AV col
            "extoch_rvr_b": _higher(pub_b, als_b),  # AW col
            "extoch_rvr_c": _higher(pub_c, als_c),  # AX col
            "extoch_rvr_d": _higher(pub_d, als_d),  # AY col
            # Final ALS output (D,H,L,P cols)
            "final_als_a":  _higher(pub_a, als_a),
            "final_als_b":  _higher(pub_b, als_b),
            "final_als_c":  _higher(pub_c, als_c),
            "final_als_d":  _higher(pub_d, als_d),
        })
    return result


# ══════════════════════════════════════════════════════════════
# MODULE 27 — NoALSRVROCH
# No ALS RVR lookup for OCH-based procedures
# 820/810 → N/A (no lights already covers it)
# Others  → look up PRVRN or RVRNPN
# ══════════════════════════════════════════════════════════════
def no_als_rvr_och(procedures: list) -> list:
    """
    No ALS RVR for OCH minima — replaces Sub NoALSRVROCH().

    Reads derived DH (AG,AJ,AM,AP) and looks up No-ALS RVR.
    820/810 → N/A (no lights value already placed by rvr_calculation)
    Others  → Vlookup(DH, Cat1PRVRN/Cat1RVRNPN)

    Returns:
        list with no_als_rvr_a/b/c/d added (AR,AS,AT,AU cols)
    """
    result = []
    for proc in procedures[:15]:
        code  = proc.get("proc_code", 0)
        dh_a  = _parse(proc.get("dh_a"))
        dh_b  = _parse(proc.get("dh_b"))
        dh_c  = _parse(proc.get("dh_c"))
        dh_d  = _parse(proc.get("dh_d"))

        if code in (820, 810):
            no_als_a = no_als_b = no_als_c = no_als_d = NA
        else:
            no_als_a = _lookup_no_als(dh_a, code, "a")
            no_als_b = _lookup_no_als(dh_b, code, "b")
            no_als_c = _lookup_no_als(dh_c, code, "c")
            no_als_d = _lookup_no_als(dh_d, code, "d")

        result.append({
            **proc,
            "no_als_rvr_a": no_als_a,  # AR col
            "no_als_rvr_b": no_als_b,  # AS col
            "no_als_rvr_c": no_als_c,  # AT col
            "no_als_rvr_d": no_als_d,  # AU col
        })
    return result


# ══════════════════════════════════════════════════════════════
# MODULE 15 — NoALSRVREXTOCH
# No ALS RVR lookup for EXTOCH procedures
# Identical table logic to Module27 — different call chain
# ══════════════════════════════════════════════════════════════
def no_als_rvr_extoch(procedures: list) -> list:
    """
    No ALS RVR for EXTOCH minima — replaces Sub NoALSRVREXTOCH().
    Same table logic as no_als_rvr_och but called from extoch path.
    """
    return no_als_rvr_och(procedures)


# ══════════════════════════════════════════════════════════════
# MODULE 14 — RVRCOMPARE
# Compare JAR ALS RVR vs JAR No ALS RVR → take higher
# Then compare with State published No ALS RVR → take higher
# ══════════════════════════════════════════════════════════════
def rvr_compare(procedures: list) -> list:
    """
    RVR comparison — replaces Sub RVRCOMPARE().

    For each procedure, for each cat:
    Step 1: JAR_NoALS = MAX(no_als_rvr, rvr_als)
            VBA: =If(AR26<=AH26, AH26, AR26)
    Step 2: If State No ALS RVR exists:
                final_no_als = MAX(JAR_NoALS, State_NoALS)
            else:
                final_no_als = JAR_NoALS

    Saves JAR value to BU/BV/BW/BX cols (for audit)

    Args:
        procedures: enriched list with rvr_als and no_als_rvr

    Returns:
        list with final_no_als_a/b/c/d added (E,I,M,Q output cols)
    """
    result = []

    for proc in procedures[:15]:
        als_a    = proc.get("rvr_als_a",    NA)
        als_b    = proc.get("rvr_als_b",    NA)
        als_c    = proc.get("rvr_als_c",    NA)
        als_d    = proc.get("rvr_als_d",    NA)

        no_als_a = proc.get("no_als_rvr_a", NA)
        no_als_b = proc.get("no_als_rvr_b", NA)
        no_als_c = proc.get("no_als_rvr_c", NA)
        no_als_d = proc.get("no_als_rvr_d", NA)

        state_a  = _parse(proc.get("state_noals_a"))
        state_b  = _parse(proc.get("state_noals_b"))
        state_c  = _parse(proc.get("state_noals_c"))
        state_d  = _parse(proc.get("state_noals_d"))

        # Stop on hyphen (end of procedures)
        if als_a == NA and no_als_a == NA:
            break

        def _compare(no_als, als, state):
            """
            VBA: =If(NoALS <= ALS, ALS, NoALS)  → JAR_val
            Then: if State exists → MAX(JAR_val, State)
            """
            if no_als == NA:
                jar_val = NA
                # Check State fallback
                if jar_val == NA and state not in (NA, None):
                    jar_val = state
            else:
                try:
                    na_num  = float(no_als)
                    als_num = float(als) if als != NA else 0
                    jar_val = als_num if na_num <= als_num else na_num
                except (ValueError, TypeError):
                    jar_val = NA

            # Compare JAR vs State No ALS — take higher
            if state not in (NA, None) and jar_val != NA:
                try:
                    jar_val = max(float(jar_val), float(state))
                except (ValueError, TypeError):
                    pass

            return jar_val

        final_no_a = _compare(no_als_a, als_a, state_a)
        final_no_b = _compare(no_als_b, als_b, state_b)
        final_no_c = _compare(no_als_c, als_c, state_c)
        final_no_d = _compare(no_als_d, als_d, state_d)

        result.append({
            **proc,
            # JAR values saved (BU-BX cols — audit trail)
            "jar_no_als_a": final_no_a,
            "jar_no_als_b": final_no_b,
            "jar_no_als_c": final_no_c,
            "jar_no_als_d": final_no_d,
            # Final No ALS output (E,I,M,Q cols)
            "final_no_als_a": final_no_a,
            "final_no_als_b": final_no_b,
            "final_no_als_c": final_no_c,
            "final_no_als_d": final_no_d,
        })

    return result


# ══════════════════════════════════════════════════════════════
# MODULE 17 — RVRCOMPAREEXTOCH
# Same as Module14 but uses EXTOCH RVR (AV,AW,AX,AY) as ALS base
# ══════════════════════════════════════════════════════════════
def rvr_compare_extoch(procedures: list) -> list:
    """
    EXTOCH RVR comparison — replaces Sub RVRCOMPAREEXTOCH().

    Same logic as rvr_compare() but uses extoch_rvr (AV-AY)
    instead of rvr_als (AH-AQ) as the ALS comparison value.

    VBA: =If(AR26<=AV26, AV26, AR26)  (extoch_rvr instead of rvr_als)
    """
    # Remap extoch_rvr → rvr_als for reuse of rvr_compare logic
    remapped = []
    for proc in procedures:
        remapped.append({
            **proc,
            "rvr_als_a": proc.get("extoch_rvr_a", proc.get("rvr_als_a", NA)),
            "rvr_als_b": proc.get("extoch_rvr_b", proc.get("rvr_als_b", NA)),
            "rvr_als_c": proc.get("extoch_rvr_c", proc.get("rvr_als_c", NA)),
            "rvr_als_d": proc.get("extoch_rvr_d", proc.get("rvr_als_d", NA)),
        })
    return rvr_compare(remapped)


# ══════════════════════════════════════════════════════════════
# MODULE 53 — RVRvsNOALS
# Final comparison: JAR RVR (ALS) vs State No ALS RVR
# Ensures JAR calculated value never drops below State published
# ══════════════════════════════════════════════════════════════
def rvr_vs_no_als(procedures: list) -> list:
    """
    Final RVR vs No ALS comparison — replaces Sub RVRvsNOALS().

    For each procedure, for each cat:
        If JAR_ALS_RVR > State_NoALS → use JAR_ALS_RVR
        Else → use State_NoALS

    VBA: =If(State<=JAR, JAR, State)

    This is the FINAL step — produces the definitive output.

    Returns:
        list with final_rvr_a/b/c/d — the definitive RVR values
    """
    result = []

    for proc in procedures[:15]:
        # JAR RVR (D,H,L,P cols — from extoch or als)
        jar_a = _parse(proc.get("final_als_a",    proc.get("rvr_als_a",    NA)))
        jar_b = _parse(proc.get("final_als_b",    proc.get("rvr_als_b",    NA)))
        jar_c = _parse(proc.get("final_als_c",    proc.get("rvr_als_c",    NA)))
        jar_d = _parse(proc.get("final_als_d",    proc.get("rvr_als_d",    NA)))

        # State No ALS RVR (E,I,M,Q cols)
        state_a = _parse(proc.get("final_no_als_a", NA))
        state_b = _parse(proc.get("final_no_als_b", NA))
        state_c = _parse(proc.get("final_no_als_c", NA))
        state_d = _parse(proc.get("final_no_als_d", NA))

        # Skip if empty or N/A
        if jar_a in (None, "") and state_a in (None, "", NA):
            continue
        if jar_a == NA:
            continue

        def _final(jar, state):
            """
            VBA: =If(State<=JAR, JAR, State)
            i.e. take the HIGHER of JAR and State No ALS
            """
            if state in (NA, None, ""):
                return jar
            if jar == NA:
                return state
            try:
                j = float(jar)
                s = float(state)
                return j if s <= j else s
            except (ValueError, TypeError):
                return jar

        result.append({
            **proc,
            "final_rvr_a": _final(jar_a, state_a),  # definitive Cat A
            "final_rvr_b": _final(jar_b, state_b),  # definitive Cat B
            "final_rvr_c": _final(jar_c, state_c),  # definitive Cat C
            "final_rvr_d": _final(jar_d, state_d),  # definitive Cat D
        })

    return result


# ══════════════════════════════════════════════════════════════
# MASTER FIXED WING RVR PIPELINE
# Runs all 7 modules in correct VBA call chain order
# ══════════════════════════════════════════════════════════════
def run_fixed_wing_rvr(procedures: list) -> dict:
    """
    Complete Fixed Wing RVR pipeline — runs all 7 modules in order.

    Mirrors exact VBA call chain:
        rvr_calculation → route check →
            NO_ALS path:  no_als_rvr_och → rvr_compare → END
            EXTOCH path:  extoch_rvr → no_als_rvr_extoch
                              → rvr_compare_extoch → rvr_vs_no_als

    Args:
        procedures: list of procedure dicts with all required fields

    Returns:
        {
            "route":      "EXTOCH" | "NO_ALS"
            "procedures": list with complete final RVR values per cat
        }
    """
    # Step 1: RVRCALCULATION (Module36) — core ALS RVR lookup + route
    calc_result = rvr_calculation(procedures)
    route       = calc_result["route"]
    procs       = calc_result["procedures"]

    if route == "NO_ALS":
        # NO_ALS path: Module27 → Module14
        procs = no_als_rvr_och(procs)
        procs = rvr_compare(procs)

    else:
        # EXTOCH path: Module42 → Module15 → Module17 → Module53
        procs = extoch_rvr(procs)
        procs = no_als_rvr_extoch(procs)
        procs = rvr_compare_extoch(procs)
        procs = rvr_vs_no_als(procs)

    return {"route": route, "procedures": procs}
