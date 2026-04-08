# Excel Indicator Navigation Demo

This repo contains a sample Excel workbook built with simulated indicator data.

## Included file

- `indicator_navigation_demo.xlsx`

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

## Notes

- This version is a native `.xlsx` file with no VBA/macros.
- Because of that, the workbook supports default collapsed groups and manual expand/collapse, but **does not auto-close other sections when one section is expanded**.
- If needed, this can be extended into a macro-enabled `.xlsm` version later.

## How it was generated

The workbook is created by:

```bash
python generate_workbook.py
```

## Tech

- Python
- openpyxl
