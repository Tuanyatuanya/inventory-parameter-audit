"""
Generate a synthetic inventory-parameter report for the demo.

Every value is randomly generated; site names and part numbers are fictional.
Some rows are deliberately seeded to violate each business rule.

    python sample_data/generate_sample_data.py
"""
from pathlib import Path

import numpy as np
import pandas as pd

RNG = np.random.default_rng(11)
SITES = [("S101", "Northbridge Plant"), ("S102", "Riverside Plant"), ("S103", "Lakeview Plant"),
         ("S104", "Hillcrest Plant"), ("S105", "Westport Plant")]
FAMILIES = ["Clip", "Bracket", "Foam Pad", "Fastener", "Trim Panel", "Seal", "Harness", "Bezel"]
INCO = ["EXW", "FCA", "FOB", "DAP", "DDP", "CPT", ""]
N = 1500


def land_safety_time(km):
    return 2 if km < 250 else 3 if km < 1000 else 5 if km < 2000 else 7


def main():
    rows = []
    for i in range(N):
        site, name = SITES[RNG.integers(len(SITES))]
        km = float(RNG.choice([RNG.uniform(20, 3900), RNG.uniform(4000, 9000), np.nan], p=[0.85, 0.05, 0.10]))
        daily = float(RNG.choice([0, round(RNG.lognormal(3, 1), 1)], p=[0.05, 0.95]))
        st = land_safety_time(km) if km == km and km < 4000 else 5
        if RNG.random() < 0.2:                       # seed Rule 2 mismatches
            st += int(RNG.integers(1, 4))
        ss = round(daily * RNG.uniform(0, 10), 0)
        if RNG.random() < 0.1:                       # seed Rule 1: SS / daily usage > 20
            ss = round(daily * RNG.uniform(21, 60), 0)
        mrp = 0 if RNG.random() < 0.08 else int(RNG.integers(100, 50000))   # seed Rule 3
        inco = str(RNG.choice(INCO))
        transport = int(RNG.integers(1, 15)) if RNG.random() < 0.5 else 0  # seed Rule 4
        rows.append({
            "Site": site, "Site Name": name,
            "Item": f"PN-{i:05d}", "Description": f"{RNG.choice(FAMILIES)} {RNG.integers(10, 99)}",
            "Type Trim": RNG.choice(["PG", "PG", "PG", "RM"]),
            "Safety Time": st, "Safety Stock": ss, "Daily Usage": daily, "Total MRP": mrp,
            "Inco Terms": inco, "Transport Days": transport,
            "Supplier Distance": round(km) if km == km else None,
        })
    df = pd.DataFrame(rows)
    path = Path(__file__).parent / "sample_inventory_report.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        pd.DataFrame([["Inventory parameter report (synthetic demo data)"]]).to_excel(w, index=False, header=False)
        df.to_excel(w, index=False, startrow=1)          # real header on row 2, like the ERP export
    print(f"wrote {path.name}: {len(df):,} rows")


if __name__ == "__main__":
    main()
