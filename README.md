# Inventory Parameter Audit Tool

A desktop tool that audits an ERP inventory-parameter export (safety stock, safety time, transport days and related fields) against supply-chain business rules. It flags every part whose settings should be corrected, proposes the new value, and writes an Excel report that planners can work through site by site.

Built during my 2026 Global Operations Excellence internship at an automotive-interiors supplier, for non-technical planners. It is packaged as a one-file Windows `.exe`. **This public version is sanitized:** the rules and code are the real ones, but all data is synthetic and no company or site names are included.

**Stack:** Python · pandas · openpyxl · Tkinter · PyInstaller

---

## Business rules
| Rule | Flag when | Proposed correction |
|---|---|---|
| **1 · Safety stock → safety time** | `Safety Stock / Daily Usage > 20` days | Set safety stock to 0 and add `⌈SS / daily usage⌉` days to safety time. Zero usage goes to manual review. |
| **2 · Safety time vs. supplier distance** | Safety time ≠ the expected land safety time for that distance | Set the expected value (see table). Missing distance or ≥ 4,000 km goes to manual review. |
| **3 · No MRP demand** | `Total MRP = 0` | Set safety stock to 0. |
| **4 · Incoterms vs. transport days** | Incoterm is not EXW / FCA / FOB / FAS (the supplier pays freight) and transport days ≠ 0 | Set transport days to 0. A blank Incoterm goes to manual review. |

| Supplier distance | Expected land safety time (days) |
|---|---:|
| < 250 km | 2 |
| 250–999 km | 3 |
| 1,000–1,999 km | 5 |
| 2,000–3,999 km | 7 |
| ≥ 4,000 km | manual review |

When a row triggers several rules, the proposals are merged by priority: MRP = 0, then Incoterms, then distance, then the safety-stock conversion. Each rule is an isolated `check_rule_*` function in [`processor.py`](processor.py), so planners' policy changes stay a one-function edit.

## Engineering details
- **Robust input parsing.** The header row is auto-detected under any title rows, column names are normalized, the distance column is matched from several aliases, and non-numeric cells are coerced safely so a bad cell never crashes a run.
- **Two-sheet Excel output.** The *Summary* sheet counts issues per site and issue type, with a status column for follow-up. The *Detail* sheet lists every row with current vs. proposed values, the triggered rules and a review flag. Problem cells are highlighted in red.
- **Built for non-technical users.** It has a Tkinter GUI with a progress bar and log, plain-English error messages, and a timestamped output name so it never overwrites a file that is open in Excel.

## Demo on synthetic data
```bash
pip install -r requirements.txt
python sample_data/generate_sample_data.py          # 1,500-row synthetic report
python cli.py sample_data/sample_inventory_report.xlsx
```
```
PG rows: 1,120
Rows needing attention: 722
Manual Review                    329
Transport Days Optimization      227
Safety Time Distance Mismatch    197
Safety Stock to Safety Time      100
Total MRP Zero Safety Stock       86
```
(The synthetic data is seeded with rule violations on purpose, so these counts are much higher than in a typical real report.)

GUI version: `python main.py`. Windows executable: run `build_exe.bat` on a Windows machine.

## Project structure
```
├── main.py             # Tkinter GUI and orchestration
├── cli.py              # headless entry point (same pipeline)
├── processor.py        # reading, header detection, the four business rules, reports
├── excel_writer.py     # formatted workbook with red highlighting
├── utils.py            # safe number/text parsing helpers
├── build_exe.bat       # PyInstaller one-file build
└── sample_data/generate_sample_data.py
```
