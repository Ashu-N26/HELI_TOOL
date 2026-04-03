# ============================================================
# main.py
# Master Pipeline for Helicopter Minima Calculation
# Connects ALL converted VBA modules in correct execution order
#
# COMPLETE FLOW — mirrors Excel VBA execution exactly:
#
#   HELIREQUEST     (Module57) → Gate: HELI needed?
#       ↓ YES
#   CheckValues     (Module60) → Validate OCH/DH/Circling
#       ↓ VALID
#   RAFSELECT       (Module45) → Gate: RAF or Non-RAF?
#       ↓ RAF                    ↓ Non-RAF
#   EXTOCHDHDARAF            EXTOCHDHDA
#   (Module43)               (Module41)
#       ↓                        ↓
#   HELIRVRCALCULATION (Module51) → Core RVR Calculation
# ============================================================

from heli_request  import heli_request
from check_values  import check_values
from raf_select    import raf_select
from heli_rvr_calc import heli_rvr_calculation
import json
from datetime import datetime


def run_heli_minima(input_data: dict) -> dict:
    """
    Master function — runs full helicopter minima pipeline.

    Args:
        input_data: {
            "icao":            str    e.g. "VABB"         <- Sheet1 C5
            "airport_name":    str    e.g. "Mumbai"       <- Sheet1 B4
            "date":            str    e.g. "2026-03-31"   <- Sheet1 E9
            "heli_flag":       str    "YES"/"NO"          <- Sheet1 L11
            "raf_flag":        str    "Y"/"N"             <- Sheet1 C11
            "glidepath_angle": float  e.g. 3.0            <- Sheet1 D26
            "procedures": [                               <- Sheet1 A26:P40
                {
                    "runway_id":  str
                    "proc_type":  int    e.g. 420         <- CB column
                    "dh":         float  e.g. 250.0       <- AG column
                    "och_a":      float                   <- E column
                    "och_b":      float                   <- H column
                    "och_c":      float                   <- K column
                    "och_d":      float                   <- N column
                    "dh_mdh_a":   float                   <- F column
                    "dh_mdh_b":   float                   <- I column
                    "dh_mdh_c":   float                   <- L column
                    "dh_mdh_d":   float                   <- O column
                }, ... up to 15
            ],
            "circling_sheet1": [                          <- Sheet1 B72:E79
                {"cir_och_a": float, "cir_och_b": float,
                 "cir_och_c": float, "cir_och_d": float}
            ],
            "circling_sheet2": [ ... ]                    <- Sheet2 same
        }
    """

    timestamp = datetime.now().isoformat()

    # ── STEP 1: HELIREQUEST ──────────────────────────────────
    # Module57 — reads Sheet1 L11
    heli_flag = input_data.get("heli_flag", "")
    request   = heli_request(heli_flag)

    if not request["proceed"]:
        return {
            "status":     "NOHELI",
            "message":    request["message"],
            "timestamp":  timestamp,
            "icao":       input_data.get("icao", ""),
            "airport":    input_data.get("airport_name", ""),
            "date":       input_data.get("date", ""),
            "raf_route":  None,
            "validation": None,
            "results":    []
        }

    # ── STEP 2: CHECK VALUES ─────────────────────────────────
    # Module60 — validates OCH, DH/MDH, Circling OCH
    procedures      = input_data.get("procedures", [])
    circling_sheet1 = input_data.get("circling_sheet1", [])
    circling_sheet2 = input_data.get("circling_sheet2", [])

    validation = check_values(procedures, circling_sheet1, circling_sheet2)

    if not validation["valid"]:
        return {
            "status":     "VALIDATION_ERROR",
            "message":    "Input validation failed. Please correct errors and retry.",
            "timestamp":  timestamp,
            "icao":       input_data.get("icao", ""),
            "airport":    input_data.get("airport_name", ""),
            "date":       input_data.get("date", ""),
            "raf_route":  None,
            "validation": validation,
            "results":    []
        }

    # ── STEP 3: RAFSELECT ────────────────────────────────────
    # Module45 — reads Sheet1 C11 (RAF flag) and D26 (GP angle)
    # Routes to EXTOCHDHDARAF (RAF) or EXTOCHDHDA (Non-RAF)
    raf_flag        = input_data.get("raf_flag", "")
    glidepath_angle = input_data.get("glidepath_angle", None)

    raf_result = raf_select(raf_flag, glidepath_angle)

    if not raf_result["valid"]:
        return {
            "status":     "RAF_ERROR",
            "message":    raf_result["message"],
            "timestamp":  timestamp,
            "icao":       input_data.get("icao", ""),
            "airport":    input_data.get("airport_name", ""),
            "date":       input_data.get("date", ""),
            "raf_route":  raf_result,
            "validation": validation,
            "results":    []
        }

    # ── STEP 4: HELI RVR CALCULATION ────────────────────────
    # Module51 — loops all runways, VLookup RVR from Tables
    rvr_results = heli_rvr_calculation(procedures)

    return {
        "status":     "SUCCESS",
        "message":    "Helicopter Minima calculation completed successfully.",
        "timestamp":  timestamp,
        "icao":       input_data.get("icao", ""),
        "airport":    input_data.get("airport_name", ""),
        "date":       input_data.get("date", ""),
        "raf_route":  raf_result,
        "validation": validation,
        "results":    rvr_results
    }


# ── TEST SUITE ────────────────────────────────────────────────
if __name__ == "__main__":

    base = {
        "icao": "VABB", "airport_name": "Mumbai",
        "date": "2026-03-31", "heli_flag": "YES",
        "raf_flag": "N", "glidepath_angle": None,
        "procedures": [
            {"runway_id": "27", "proc_type": 120, "dh": 250.0,
             "och_a": 230.0, "och_b": 250.0, "och_c": 280.0, "och_d": 300.0,
             "dh_mdh_a": 200.0, "dh_mdh_b": 220.0, "dh_mdh_c": 250.0, "dh_mdh_d": 270.0},
            {"runway_id": "09", "proc_type": 310, "dh": 300.0,
             "och_a": 280.0, "och_b": 300.0, "och_c": 330.0, "och_d": 360.0,
             "dh_mdh_a": 280.0, "dh_mdh_b": 300.0, "dh_mdh_c": 330.0, "dh_mdh_d": 360.0},
        ],
        "circling_sheet1": [
            {"cir_och_a": 400, "cir_och_b": 500,
             "cir_och_c": 600, "cir_och_d": 700}],
        "circling_sheet2": []
    }

    tests = [
        ("TEST 1: Non-RAF — should SUCCEED",
         {**base}),
        ("TEST 2: RAF with glidepath — should route to RAF",
         {**base, "raf_flag": "Y", "glidepath_angle": 3.0}),
        ("TEST 3: RAF flag Y, no glidepath — should RAF_ERROR",
         {**base, "raf_flag": "Y", "glidepath_angle": None}),
        ("TEST 4: HELI flag NO — should NOHELI",
         {**base, "heli_flag": "NO"}),
        ("TEST 5: OCH Cat B < Cat A — should VALIDATION_ERROR",
         {**base, "procedures": [{
             "runway_id": "27", "proc_type": 120, "dh": 250.0,
             "och_a": 300.0, "och_b": 250.0,  # B < A
             "och_c": 280.0, "och_d": 300.0,
             "dh_mdh_a": 200.0, "dh_mdh_b": 220.0,
             "dh_mdh_c": 250.0, "dh_mdh_d": 270.0}]}),
    ]

    for title, test_data in tests:
        print(f"\n{'='*60}")
        print(title)
        print('='*60)
        r = run_heli_minima(test_data)
        print(f"  Status:    {r['status']}")
        if r.get("raf_route"):
            print(f"  RAF Route: {r['raf_route']['route']} → {r['raf_route'].get('next_step','')}")
        if r.get("validation") and not r["validation"]["valid"]:
            print(f"  Errors:    {r['validation']['errors']}")
        if r.get("results"):
            for rwy in r["results"]:
                print(f"\n  Runway: {rwy['runway_id']} | {rwy['proc_name']} | DH={rwy['dh']}ft")
                print(f"  Full={rwy['rvr_full']}m | Inter={rwy['rvr_inter']}m | Basic={rwy['rvr_basic']}m | None={rwy['rvr_none']}m")
        print(f"  Message:   {r['message']}")

if __name__ == "__main__":
    import json

    with open("input.json") as f:
        data = json.load(f)

    result = run_heli_minima(data)

    with open("output.json", "w") as f:
        json.dump(result, f, indent=2)

    print("Done")
