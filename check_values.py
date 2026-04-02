# ============================================================
# check_values.py
# Input Validator for Helicopter / Cat1 Minima
# Replaces: Sub CheckValues() — VBA Module60
#
# Original VBA by: Alan Hutchinson, 25th September 2001
# Amended: 10th January 2002, 16th May 2002
#
# COMPLETE CONVERSION — all VBA logic preserved:
#
# Loop 1 (x=0..15): OCH Straight-In        cols E,H,K,N rows 26-40
#   If OCH_A = "-"  → skip to DHMDH section (not an error)
#   If OCH_A = "-1" → skip Cat A, continue to Cat B check
#   If OCH_A = ""   → exit loop (no more procedures)
#   Rules: Cat A <= B <= C <= D
#
# Loop 2 (x=0..15): DH/MDH Straight-In     cols F,I,L,O rows 26-40
#   If DH_A = "-"   → skip to CIRC section (not an error)
#   If DH_A = "-1"  → skip Cat A, continue to Cat B check
#   If DH_A = ""    → exit loop
#   Rules: Cat A <= B <= C <= D
#
# Loop 3 (x=0..8):  Circling OCH Sheet1    cols B,C,D,E rows 72-79
#   If x=8 and Sheet2 B72 is empty → GoTo FINISH (skip Sheet2)
#   If value = "-1" → skip that category (AGAIN label)
#
# Loop 4 (x=0..8):  Circling OCH Sheet2    same structure
#
# QUEST dialog (after any error):
#   "Correct and restart? Y/N"
#   → "continue" = proceed despite error (still calls RAFSELECT)
#   → "restart"  = stop (user will fix and re-run)
#   In Python: "force_continue" flag controls this behaviour
#
# Calls: RAFSELECT at end of FINISH
# ============================================================

NA_VALUE   = "-"       # Not applicable — skip to next section
SKIP_VALUE = "-1"      # Category not used — skip this cat, continue loop
EMPTY      = (None, "")


def _parse(val):
    """Parse a cell value. Returns float, NA_VALUE, SKIP_VALUE, or None."""
    if val is None or str(val).strip() == "":
        return None
    s = str(val).strip()
    if s == NA_VALUE:
        return NA_VALUE
    if s == SKIP_VALUE:
        return SKIP_VALUE
    try:
        return float(s)
    except ValueError:
        return s


# ── LOOP 1: OCH STRAIGHT-IN ─────────────────────────────────
# VBA: Do Until x = 15
#        OCHA="E"&y, OCHB="H"&y, OCHC="K"&y, OCHD="N"&y
#        If OCH_A = "-"  → GoTo DHMDH
#        If OCH_A = "-1" → GoTo CATB (skip A, check B onwards)
#        If OCH_A = ""   → Exit Do
def _check_och(procedures: list) -> dict:
    """
    Loop 1 — Validate OCH values for Straight-In.
    Returns: { valid, errors, skip_to_dhmdh }
    skip_to_dhmdh=True means "-" found, jump straight to DH check.
    """
    errors = []

    if not procedures:
        return {"valid": False, "errors": ["Error - OCH must not be BLANK!"],
                "skip_to_dhmdh": False}

    # First row blank check (x=1 in VBA)
    first_a = _parse(procedures[0].get("och_a"))
    if first_a is None:
        return {"valid": False, "errors": ["Error - OCH must not be BLANK!"],
                "skip_to_dhmdh": False}

    # If first OCH_A is "-" → skip entire OCH check, go to DHMDH
    # VBA: If ActiveCell.Value = "-" Then GoTo DHMDH
    if first_a == NA_VALUE:
        return {"valid": True, "errors": [], "skip_to_dhmdh": True}

    for i, proc in enumerate(procedures[:15]):
        rwy = proc.get("runway_id", f"Runway {i+1}")

        a = _parse(proc.get("och_a"))
        b = _parse(proc.get("och_b"))
        c = _parse(proc.get("och_c"))
        d = _parse(proc.get("och_d"))

        # Empty → exit loop (no more procedures)
        if a is None:
            break

        # "-" anywhere in Cat A → skip to DHMDH section
        if a == NA_VALUE:
            return {"valid": True, "errors": [], "skip_to_dhmdh": True}

        # "-1" → skip Cat A, still check B onwards (CATB label)
        a_val = None if a == SKIP_VALUE else a

        # Cat B
        if b == NA_VALUE:
            return {"valid": True, "errors": [], "skip_to_dhmdh": True}
        b_val = None if b == SKIP_VALUE else b
        if isinstance(a_val, float) and isinstance(b_val, float):
            if b_val < a_val:
                errors.append(
                    f"{rwy}: Cat B OCH ({b_val}) is lower than Cat A ({a_val}). "
                    f"Correction required!"
                )

        # Cat C
        if c == NA_VALUE:
            return {"valid": True, "errors": [], "skip_to_dhmdh": True}
        c_val = None if c == SKIP_VALUE else c
        prev = b_val if isinstance(b_val, float) else a_val
        if isinstance(prev, float) and isinstance(c_val, float):
            if c_val < prev:
                errors.append(
                    f"{rwy}: Cat C OCH ({c_val}) is lower than Cat B. "
                    f"Correction required!"
                )

        # Cat D
        if d == NA_VALUE:
            return {"valid": True, "errors": [], "skip_to_dhmdh": True}
        d_val = None if d == SKIP_VALUE else d
        prev = c_val if isinstance(c_val, float) else prev
        if isinstance(prev, float) and isinstance(d_val, float):
            if d_val < prev:
                errors.append(
                    f"{rwy}: Cat D OCH ({d_val}) is lower than Cat C. "
                    f"Correction required!"
                )

    return {"valid": len(errors) == 0, "errors": errors, "skip_to_dhmdh": False}


# ── LOOP 2: DH/MDH STRAIGHT-IN ──────────────────────────────
# VBA: Do Until x = 15
#        DHMDHA="F"&y, DHMDHB="I"&y, DHMDHC="L"&y, DHMDHD="O"&y
#        If DH_A = "-"  → GoTo CIRC
#        If DH_A = "-1" → GoTo DHCATB
#        If DH_A = ""   → Exit Do
def _check_dh_mdh(procedures: list) -> dict:
    """
    Loop 2 — Validate DH/MDH values for Straight-In.
    Returns: { valid, errors, skip_to_circ }
    skip_to_circ=True means "-" found, jump straight to Circling check.
    """
    errors = []

    if not procedures:
        return {"valid": False,
                "errors": ["Error - DH/MDH must not be BLANK!"],
                "skip_to_circ": False}

    first_a = _parse(procedures[0].get("dh_mdh_a"))
    if first_a is None:
        return {"valid": False,
                "errors": ["Error - DH/MDH must not be BLANK!"],
                "skip_to_circ": False}

    # "-" on first → skip to CIRC
    if first_a == NA_VALUE:
        return {"valid": True, "errors": [], "skip_to_circ": True}

    for i, proc in enumerate(procedures[:15]):
        rwy = proc.get("runway_id", f"Runway {i+1}")

        a = _parse(proc.get("dh_mdh_a"))
        b = _parse(proc.get("dh_mdh_b"))
        c = _parse(proc.get("dh_mdh_c"))
        d = _parse(proc.get("dh_mdh_d"))

        if a is None:
            break

        # "-" → GoTo CIRC
        if a == NA_VALUE:
            return {"valid": True, "errors": [], "skip_to_circ": True}

        a_val = None if a == SKIP_VALUE else a

        if b == NA_VALUE:
            return {"valid": True, "errors": [], "skip_to_circ": True}
        b_val = None if b == SKIP_VALUE else b
        if isinstance(a_val, float) and isinstance(b_val, float):
            if b_val < a_val:
                errors.append(
                    f"{rwy}: Cat B DH/MDH ({b_val}) is lower than Cat A ({a_val}). "
                    f"Correction required!"
                )

        if c == NA_VALUE:
            return {"valid": True, "errors": [], "skip_to_circ": True}
        c_val = None if c == SKIP_VALUE else c
        prev = b_val if isinstance(b_val, float) else a_val
        if isinstance(prev, float) and isinstance(c_val, float):
            if c_val < prev:
                errors.append(
                    f"{rwy}: Cat C DH/MDH ({c_val}) is lower than Cat B. "
                    f"Correction required!"
                )

        # VBA: If DH_D = "-1" → GoTo CIRC (skip D entirely)
        if d == NA_VALUE or d == SKIP_VALUE:
            continue
        d_val = d
        prev = c_val if isinstance(c_val, float) else prev
        if isinstance(prev, float) and isinstance(d_val, float):
            if d_val < prev:
                errors.append(
                    f"{rwy}: Cat D DH/MDH ({d_val}) is lower than Cat C. "
                    f"Correction required!"
                )

    return {"valid": len(errors) == 0, "errors": errors, "skip_to_circ": False}


# ── LOOPS 3 & 4: CIRCLING OCH SHEET1 + SHEET2 ───────────────
# VBA Loop 3: Do Until x = 8, cols B,C,D,E rows 72-79, Sheet1
#   If x=8 AND Sheet2 B72 is empty → GoTo FINISH (skip Sheet2)
# VBA Loop 4: same structure, Sheet2
def _check_circling_sheet(rows: list, sheet_label: str) -> dict:
    """Check one sheet's circling OCH rows."""
    errors = []
    for i, row in enumerate(rows[:8]):
        a = _parse(row.get("cir_och_a"))
        b = _parse(row.get("cir_och_b"))
        c = _parse(row.get("cir_och_c"))
        d = _parse(row.get("cir_och_d"))

        # Empty → exit loop
        if a is None:
            break

        label = f"{sheet_label} Row {i+1}"
        a_val = None if a == SKIP_VALUE else a

        # Cat B
        b_val = None if b == SKIP_VALUE else b
        if isinstance(a_val, float) and isinstance(b_val, float):
            if b_val < a_val:
                errors.append(
                    f"{label}: Circling Cat B OCH ({b_val}) is lower than "
                    f"Cat A ({a_val}). Correction required!"
                )

        # Cat C
        c_val = None if c == SKIP_VALUE else c
        prev = b_val if isinstance(b_val, float) else a_val
        if isinstance(prev, float) and isinstance(c_val, float):
            if c_val < prev:
                errors.append(
                    f"{label}: Circling Cat C OCH ({c_val}) is lower than "
                    f"Cat B. Correction required!"
                )

        # Cat D — "-1" → AGAIN (skip D, continue loop = no error check on D)
        if d == SKIP_VALUE:
            continue
        d_val = d
        prev = c_val if isinstance(c_val, float) else prev
        if isinstance(prev, float) and isinstance(d_val, float):
            if d_val < prev:
                errors.append(
                    f"{label}: Circling Cat D OCH ({d_val}) is lower than "
                    f"Cat C. Correction required!"
                )

    return {"valid": len(errors) == 0, "errors": errors}


def _check_circling(circling_sheet1: list,
                    circling_sheet2: list,
                    sheet2_b72_empty: bool = False) -> dict:
    """
    Loops 3 & 4 — Validate Circling OCH Sheet1 and Sheet2.

    VBA critical logic:
        After Sheet1 loop (x=8):
            If Sheet2 B72 is empty → GoTo FINISH (skip Sheet2 entirely)

    Args:
        circling_sheet1:   list of up to 8 circling row dicts
        circling_sheet2:   list of up to 8 circling row dicts
        sheet2_b72_empty:  True if Sheet2 cell B72 is empty
                           (triggers VBA GoTo FINISH — skip Sheet2)
    """
    errors = []

    # Loop 3 — Sheet1
    if circling_sheet1:
        r1 = _check_circling_sheet(circling_sheet1, "Sheet1 Circling")
        errors.extend(r1["errors"])

    # VBA: If x = 8 Then check Sheet2 B72
    #      If empty → GoTo FINISH (skip Sheet2 loop)
    if sheet2_b72_empty:
        return {"valid": len(errors) == 0, "errors": errors}

    # Loop 4 — Sheet2 (only if Sheet2 has data)
    if circling_sheet2:
        r2 = _check_circling_sheet(circling_sheet2, "Sheet2 Circling")
        errors.extend(r2["errors"])

    return {"valid": len(errors) == 0, "errors": errors}


# ── MASTER FUNCTION ──────────────────────────────────────────
def check_values(procedures: list,
                 circling_sheet1: list,
                 circling_sheet2: list = None,
                 sheet2_b72_empty: bool = None,
                 force_continue: bool = False) -> dict:
    """
    Master validation — mirrors complete Sub CheckValues() flow.

    Execution order exactly matches VBA:
        Loop 1: OCH Straight-In
            "-" found → skip to DH/MDH (not an error)
        Loop 2: DH/MDH Straight-In
            "-" found → skip to Circling (not an error)
        Loop 3: Circling Sheet1
            If Sheet2 B72 empty → skip Sheet2
        Loop 4: Circling Sheet2
        QUEST: On any error → option to continue or restart
        FINISH: Call RAFSELECT

    Args:
        procedures:         list of procedure dicts (max 15)
        circling_sheet1:    list of circling OCH dicts (max 8) Sheet1
        circling_sheet2:    list of circling OCH dicts (max 8) Sheet2
        sheet2_b72_empty:   True  = Sheet2 has no circling data
                            False = Sheet2 has data, run Loop 4
                            None  = auto-detect from circling_sheet2
        force_continue:     True = continue despite errors (VBA QUEST "No")
                            False = stop on errors (default)

    Returns:
        {
            "valid":           bool   True = proceed to RAFSELECT
            "errors":          list   all error messages found
            "och_result":      dict
            "dh_result":       dict
            "circ_result":     dict
            "force_continued": bool   True if errors found but user continued
        }
    """
    circling_sheet2 = circling_sheet2 or []

    # Auto-detect Sheet2 empty state if not explicitly provided
    if sheet2_b72_empty is None:
        sheet2_b72_empty = len(circling_sheet2) == 0

    all_errors = []

    # ── Loop 1: OCH ────────────────────────────────────────
    och_result = _check_och(procedures)
    all_errors.extend(och_result["errors"])

    # If "-" in OCH → skip OCH check, jump directly to DH/MDH
    # VBA: GoTo DHMDH
    if och_result.get("skip_to_dhmdh"):
        dh_result = _check_dh_mdh(procedures)
        all_errors.extend(dh_result["errors"])
    else:
        # ── Loop 2: DH/MDH ─────────────────────────────────
        dh_result = _check_dh_mdh(procedures)
        all_errors.extend(dh_result["errors"])

    # If "-" in DH/MDH → skip to Circling
    # VBA: GoTo CIRC
    # (We always run Circling regardless — "-" just skips DH errors)

    # ── Loops 3 & 4: Circling ──────────────────────────────
    circ_result = _check_circling(
        circling_sheet1, circling_sheet2, sheet2_b72_empty
    )
    all_errors.extend(circ_result["errors"])

    # ── QUEST dialog logic ──────────────────────────────────
    # VBA: If errors → MsgBox "Correct and restart? Y/N"
    #      No  → GoTo FINISH (continue to RAFSELECT)
    #      Yes → End (stop completely)
    # Python: force_continue=True mirrors "No" (continue despite errors)
    has_errors = len(all_errors) > 0
    force_continued = False

    if has_errors and force_continue:
        # VBA: Case vbNo → GoTo FINISH
        force_continued = True
        valid = True
    else:
        valid = not has_errors

    return {
        "valid":           valid,
        "errors":          all_errors,
        "och_result":      och_result,
        "dh_result":       dh_result,
        "circ_result":     circ_result,
        "force_continued": force_continued,
    }
