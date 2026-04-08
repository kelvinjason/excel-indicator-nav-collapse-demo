# Excel Indicator Navigation Demo

This repo contains a sample Excel workbook built with simulated indicator data.

## Included files

- `indicator_navigation_demo.xlsx` (native version, no VBA)
- `indicator_navigation_demo_macro.xlsm` (macro version, supports single-open toggle)

## What it does

### Sheet 1: `Summary`
- Shows a category overview table and a detailed indicator summary table.
- Each indicator name is clickable.
- Clicking an indicator jumps to the corresponding detail anchor row on the `Indicators` sheet.

### Sheet 2: `Indicators`
- Contains multiple category-specific tables on one sheet.
- Each category block is **collapsed by default**.
- Users can expand/collapse each category with Excel's native outline **+ / -** controls on the left side.
- A `Toggle` column is included as a visual cue for the grouped blocks.

### Macro behavior (`.xlsm`)
- In `indicator_navigation_demo_macro.xlsm`, click the `+ / -` cell in column A on a category row.
- It will **expand that one category and auto-collapse all other categories**.
- This implements the requested single-open accordion behavior.

## Notes

- `.xlsx` is macro-free and supports manual group expand/collapse.
- `.xlsm` includes VBA for auto-collapse of non-selected categories.
- If Excel opens with Protected View, enable editing and macros to use the toggle automation.

## How it was generated

The workbook is created by:

```bash
python generate_workbook.py
```

## Tech

- Python
- openpyxl
