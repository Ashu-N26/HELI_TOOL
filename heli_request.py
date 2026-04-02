# ============================================================
# heli_request.py
# Entry point gate for Helicopter Minima
# Replaces: Sub HELIREQUEST() — VBA Module57
#
# Original VBA by: Alan Hutchinson, 14th June 2002
# ============================================================

# ----------------------------------------------------------
# HELIREQUEST
# Checks if helicopter minima is required for this aerodrome
# Reads: Sheet1 cell L11 (YES/NO flag)
# ----------------------------------------------------------
def heli_request(heli_flag: str) -> dict:
    """
    Determine if helicopter minima calculation should proceed.

    Equivalent of VBA:
        If Range("L11") = "YES"/"Y" → ASK user
        If Range("L11") = "NO"/"N"  → NOHELI

    Args:
        heli_flag: value from Sheet1 L11
                   e.g. "YES", "Yes", "Y", "y",
                        "NO",  "No",  "N", "n",
                        "" (empty)

    Returns:
        dict:
            {
                "proceed":  bool    True = run calculation
                "message":  str     Status message for UI
                "status":   str     "ASK" | "NOHELI" | "PROCEED"
            }
    """
    flag = str(heli_flag).strip().upper() if heli_flag else ""

    # Empty or YES → Ask user confirmation (in Python = default proceed)
    if flag in ("", "YES", "Y"):
        return {
            "proceed": True,
            "status":  "ASK",
            "message": "Helicopter Minima being calculated."
        }

    # NO → Skip helicopter minima
    if flag in ("NO", "N"):
        return {
            "proceed": False,
            "status":  "NOHELI",
            "message": "Helicopter Minima NOT required for this aerodrome."
        }

    # Unknown value → treat as needs review
    return {
        "proceed": False,
        "status":  "UNKNOWN",
        "message": f"Unknown HELI flag value: '{heli_flag}'. Expected YES or NO."
    }
