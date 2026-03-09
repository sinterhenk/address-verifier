# Excel Chart Creator Add-In

Recreates canvas-style charts in Excel from any selected data range.
Supports 12 chart types with customisable colours, placement, and labels.

---

## Supported Chart Types

| Display Name       | Excel Type           |
|--------------------|----------------------|
| Column (Vertical)  | Clustered Column     |
| Bar (Horizontal)   | Clustered Bar        |
| Line               | Line                 |
| Line with Markers  | Line with Markers    |
| Pie                | Pie                  |
| Doughnut           | Doughnut             |
| Area               | Area                 |
| Scatter            | XY Scatter           |
| Scatter with Lines | XY Scatter + Lines   |
| Stacked Column     | Stacked Column       |
| Stacked Bar        | Stacked Bar          |
| Radar              | Radar                |

---

## Files

| File               | Purpose                                      |
|--------------------|----------------------------------------------|
| `ChartCreator.bas` | Core chart logic — import into any workbook  |
| `ChartForm.frm`    | UI form (builds itself at runtime, no .frx needed) |

---

## Installation (one-time setup)

### Option A — Add to an existing workbook

1. Open your workbook in Excel.
2. Press **Alt + F11** to open the VBA editor.
3. In the **Project Explorer** (left panel), right-click your workbook name → **Import File**.
4. Import `ChartCreator.bas`.
5. Repeat for `ChartForm.frm`.
6. Close the VBA editor.

### Option B — Create a personal add-in (.xlam) so it's always available

1. Open a blank workbook.
2. Import both files as above (Option A).
3. In the VBA editor, paste this into `ThisWorkbook`:

```vba
Private Sub Workbook_Open()
    ' Add a ribbon button or just use Alt+F8 → ShowChartCreator
End Sub
```

4. Save as **Excel Add-In** (`File → Save As → Excel Add-In *.xlam`).
   Suggested name: `ChartCreator.xlam`
5. Install it: `File → Options → Add-Ins → Go → Browse` → select your `.xlam`.

### Option C — Quick button on a sheet

After importing both files, insert a shape or button on any sheet,
right-click it → **Assign Macro** → select `ChartCreator.ShowChartCreator`.

---

## How to Use

1. **Select your data range** in the spreadsheet (include headers).
2. Run the macro: **Alt + F8** → `ChartCreator.ShowChartCreator` → **Run**
   (or click your assigned button).
3. In the dialog:
   - Confirm / adjust the **Data Range** (or click `...` to re-pick).
   - Choose a **Chart Type**.
   - (Optional) Add a **Chart Title**.
   - Choose **Placement** — right of data, below data, or a new sheet.
   - Choose a **Color Scheme**.
   - Set **Width / Height** in points (1 pt ≈ 1/72 inch; default 480 × 300).
   - Tick **headers**, **legend**, and **data labels** as needed.
4. Click **Create Chart**.

---

## Data Layout

```
         A          B          C
1   Category    Series 1   Series 2   ← header row (tick "First row = headers")
2   Jan         120        80
3   Feb         140        95
4   Mar         110        105
```

For **Pie / Doughnut** charts, use a single data column:

```
         A          B
1   Category    Value
2   Apples       40
3   Oranges      30
4   Bananas      30
```

---

## Color Schemes

| Name       | Description                              |
|------------|------------------------------------------|
| Default    | Excel's built-in theme colours           |
| Office     | Classic Office blue/orange palette       |
| Vivid      | High-contrast, bright colours            |
| Pastel     | Soft, presentation-friendly tones        |
| Greyscale  | Black-and-white printing                 |
| Dark       | Teal/amber palette for dark backgrounds  |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Cannot run macro" error | Make sure macros are enabled: `File → Options → Trust Center → Macro Settings → Enable all macros` |
| Form appears blank | The form builds controls at runtime — ensure `ChartForm.frm` was imported correctly |
| Chart lands in wrong place | Use the `...` range picker to confirm the exact range address |
| Colors not applying | "Default" intentionally skips custom colors; pick any other scheme |
