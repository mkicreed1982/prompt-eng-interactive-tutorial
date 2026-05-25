# Objective
Execute a Python script using `openpyxl` to process the attached financial workbook. Generate a new worksheet named **"BS ANALYSIS REPORT"** inside the loaded workbook containing a 9-section financial analysis driven by live Excel formulas and 3D charts. The script must be reusable across any uploaded balance sheet workbook regardless of the source sheet's name or the specific reporting years.

## Technical Constraints
- Ban all usage of `openpyxl.drawing` shapes, connectors, lines, and formula auditing/trace arrows.
- Ban the `FORMULATEXT()` function entirely to prevent `#NAME?` errors.
- Ban "Source coordinate basis" rows, Row 34 repetitions, and any metadata footer rows. Do not repeat headers or add padding rows.
- Do not instantiate `openpyxl.Workbook()`. The output workbook must be the **original workbook loaded via `load_workbook()`** with one new sheet appended. All pre-existing sheets must remain intact and unmodified.
- Do not hardcode the source sheet name. Resolve it at runtime per **Source Sheet Detection** and reference it by variable in every formula string.
- Do not hardcode reporting years or year-column letters. Detect them at runtime per **Period Detection**.
- Populate Column B with live, calculable Excel formulas referencing the detected source sheet via Python f-string interpolation (e.g., `f"='{SOURCE_SHEET}'!{COL_LATER}10/'{SOURCE_SHEET}'!{COL_LATER}20"`). The cell must evaluate to a number.
- Cap Column G content at **3 words maximum**. Reject sentences, percentages, and multi-clause phrases.

## Source Sheet Detection
The attached workbook may name its balance sheet anything (`FStmts.`, `Balance Sheet`, `BS`, `Financial Statements`, `Financials`, etc.). Resolve the source sheet name at runtime and store it as `SOURCE_SHEET`. Apply this priority chain:

1. If the workbook contains only one sheet, use it.
2. Else, match sheet names **case-insensitively** against this pattern list (first hit wins): `fstmts`, `financial statements`, `balance sheet`, `financials`, `statements`, `bs`.
3. Else, scan each sheet's Column A for canonical line items (`Total Assets`, `Total Liabilities`, `Total Current Assets`, `Total Current Liabilities`). Select the sheet with the highest match count.
4. If no candidate is found, raise `ValueError("Source balance sheet not detected")`. Do not guess.

Every formula written into the report sheet must reference `SOURCE_SHEET` dynamically. Do not bake the literal string `FStmts.` (or any other specific sheet name) into any cell.

## Period Detection
The source sheet contains two reporting years in adjacent columns. Resolve them at runtime:

1. Scan the header row(s) of the source sheet for cells whose value is a four-digit year (1990–2100) or a `datetime` type.
2. Assign the **earlier** year to `YEAR_EARLIER` and the **later** year to `YEAR_LATER`. Store their column letters as `COL_EARLIER` and `COL_LATER`.
3. Use these variables in every formula and label. Do not hardcode `2020`, `2021`, `C`, `D`, or any other fixed value.

## Layout Rules
- Place text, formulas, and numerical data **exclusively in Columns A through G** of the report sheet.
- Leave **Columns H and I empty** as a visual buffer.
- Place all 3D charts **exclusively in Columns J through P**.
- Anchor each of the 9 sections to a strict 30-row vertical block: Section 1 → Row 10, Section 2 → Row 40, Section 3 → Row 70, etc.

## Table Schema (strict)
Every section must use exactly these 7 columns. Do not invent, rename, or reorder.

```json
{
  "A": "Metric / Line Item",
  "B": "Value — live numerical formula referencing SOURCE_SHEET, must evaluate to a number",
  "C": "Selected Period",
  "D": "Strength & Weakness — Excel IF statement on Column B returning \"Company's Strength\" or \"Company's Weakness\"",
  "E": "Chart Category Label",
  "F": "Chart Dynamic Data Source — paired values for YEAR_EARLIER and YEAR_LATER",
  "G": "Analysis — 1-3 word descriptor of the YEAR_EARLIER→YEAR_LATER movement of the Column F value"
}
```

## Section Schema
Implement exactly these 9 sections with standard financial formulas:

1. **Liquidity Profile** — Current & Quick Ratios
2. **Solvency & Leverage** — Debt-to-Equity, Interest Coverage
3. **Operating Efficiency** — Asset Turnover, Inventory Turnover
4. **Profitability Margins** — Gross, Operating, Net
5. **DuPont Analysis** — Component Breakdown
6. **Cash Conversion Cycle** — DSO, DIO, DPO
7. **Working Capital Trends**
8. **Capital Structure Composition**
9. **Final Risk Flags & Thresholds**

## Analysis Column Rules (Column G)
- Summarize the directional movement of the Column F value from `YEAR_EARLIER` to `YEAR_LATER`.
- Maximum **3 words**. Use a *direction + magnitude* pattern.
- **Approved style:** `adequate increase`, `stable performance`, `slow growth`, `sharp decline`, `moderate recovery`, `flat trend`, `minor improvement`.
- **Forbidden:** full sentences, percentages, narrative explanations, repeating the metric name.
- Apply Column G to every row that holds paired earlier-year and later-year values.

## Chart Data Mapping
- **Section 1:** 3D Bar Chart comparing strictly *Total Current Assets* vs. *Total Current Liabilities*. Exclude ratio rows from the chart data range.
- **Sections 2–9:** Select only 2 to 4 comparable rows from Column F per chart. Ensure axis scale is visually meaningful.
- Anchor every chart inside **Columns J through P**. Charts must not overlap Columns A through I.

## Execution Workflow
1. **Load the original workbook with `openpyxl.load_workbook(<path>)`.** Read row/column coordinates for all required line items on the balance report sheet for both the year periods presented in the columns. Do not modify any existing sheet.
2. Resolve `SOURCE_SHEET`, `YEAR_EARLIER`, `YEAR_LATER`, `COL_EARLIER`, and `COL_LATER` per **Source Sheet Detection** and **Period Detection**.
3. Create the report sheet via `wb.create_sheet("BS ANALYSIS REPORT")`. Write financial formulas into the Column A–G grid for all 9 sections following the **Table Schema**, interpolating `SOURCE_SHEET` into every formula string.
4. For every row holding paired earlier-year and later-year values, compute the directional movement and write a 1–3 word descriptor into Column G per **Analysis Column Rules**.
5. Create a Data Validation dropdown in Cell **B2** listing the detected year periods. Construct `INDEX/MATCH` formulas in a hidden column as the dynamic data source driven by Cell B2.
6. Instantiate exactly **nine** 3D chart objects anchored in Columns J through P. Bind data sources per **Chart Data Mapping**.
7. Save the workbook under a **new filename**. The saved file must contain every original sheet plus the appended `BS ANALYSIS REPORT` sheet, all intact and unmodified.

## Contrastive Examples

<anti_pattern>

```python
wb = Workbook()                                                    # destroys source data
ws["B5"] = '=FORMULATEXT(FStmts.!C12)'                             # produces #NAME?, hardcodes name
ws["D5"] = "Formula Text"                                          # invented column
ws["B6"] = "='FStmts.'!C12/'FStmts.'!C25"                          # hardcoded sheet name
ws["F5"] = "='FStmts.'!D12"                                        # hardcoded year column letter
ws["G5"] = "The current ratio increased significantly indicating better liquidity"  # sentence
ws["G6"] = "Up 12.4% YoY"                                          # percentage, banned
ws["A34"] = "Source coordinate basis"                              # padding row, banned
chart_anchor = "H10"                                                # overlaps buffer zone
chart.add_data(Reference(ws, min_col=2, min_row=10, max_row=18))   # mixes ratios with currency
```

</anti_pattern>

<required_pattern>

```python
wb = load_workbook(path)                                                            # preserves all original sheets
SOURCE_SHEET = detect_source_sheet(wb)                                              # e.g., "FStmts." or "Balance Sheet"
YEAR_EARLIER, YEAR_LATER, COL_EARLIER, COL_LATER = detect_periods(wb[SOURCE_SHEET])
ws = wb.create_sheet("BS ANALYSIS REPORT")
ws["B5"] = f"='{SOURCE_SHEET}'!{COL_LATER}12/'{SOURCE_SHEET}'!{COL_LATER}25"
ws["D5"] = '=IF(B5>1.5, "Company\'s Strength", "Company\'s Weakness")'
ws["F5"] = f"='{SOURCE_SHEET}'!{COL_EARLIER}12"
ws["F6"] = f"='{SOURCE_SHEET}'!{COL_LATER}12"
ws["G5"] = "adequate increase"
ws["G6"] = "stable performance"
ws["G7"] = "slow growth"
chart_anchor = "J10"                                                                # inside chart zone
chart.add_data(Reference(ws, min_col=6, min_row=15, max_row=16))                    # comparable rows only
wb.save(new_path)                                                                   # original sheets preserved
```

</required_pattern>

## Final Output Requirement
Execute the script directly via the Python (Code Interpreter) tool against the attached workbook. Return the modified `.xlsx` file as a downloadable attachment in the chat. Do not paste the source code into the chat unless the user explicitly asks for it. Do not narrate the steps you took.
