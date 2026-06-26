# Rejection Analysis Compiler

Compiles Process and Final rejection sheets from a monthly workbook into a single consolidated sheet.

## Setup

Install dependencies:
```
pip install streamlit pandas openpyxl
```

Run the app:
```
streamlit run rejection_compiler.py
```

## How to Use

1. Open the app in your browser
2. Upload the monthly workbook (`.xlsx`)
3. Click **Compile**
4. Download the updated workbook — it will contain a new **COMPILED** sheet

## Workbook Requirements

- The workbook must have exactly one sheet with "process" in its name and one with "final" in its name (case insensitive)
- Both sheets must have `DATE` and `COMPONENT` columns
- Sheets can have blank rows at the top (logo/header area) — the script detects the header row automatically

## Matching Logic

- Each Process row is matched to the nearest Final row with the same component where the Final date is on or after the Process date
- Process rows with no Final match → included with Date Final blank
- Final rows with no Process match → included with Date Process blank
- `REJ. %` and `REJ. PPM` are recalculated from summed totals, not summed directly

## Output Columns

`DATE PROCESS` · `DATE FINAL` · `COMPONENT` · `CUSTOMER` · `TOTAL PRODUCTION` · `TOTAL REJECTION` · `REJ. %` · `REJ. PPM` · all defect columns
