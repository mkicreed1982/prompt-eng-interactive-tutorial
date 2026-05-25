# Objective
Execute a Python script using `openpyxl` to process the attached financial workbook. Generate a new worksheet named **"BS ANALYSIS REPORT"** containing a 9-section financial analysis driven by live Excel formulas and 3D charts.

## Technical Constraints
- Ban all usage of `openpyxl.drawing` shapes, connectors, lines, and formula auditing/trace arrows.
- Ban the `FORMULATEXT()` function entirely to prevent `#NAME?` errors.
- Ban "Source coordinate basis" rows, Row 34 repetitions, and any metadata footer rows. Do not repeat headers or add padding rows.
- Populate Column B with live, calculable Excel formulas referencing the `FStmts.` sheet (e.g., `=FStmts.!C10/FStmts.!C20`). The cell must evaluate to a number.
- Cap Column G content at **3 words maximum**. Reject sentences, percentages, and multi-clause phrases.

## Layout Rules
- Place text, formulas, and numerical data **exclusively in Columns A through G**.
- Leave **Columns H and I empty** as a visual buffer.
- Place all 3D charts **exclusively in Columns J through P**.
- Anchor each of the 9 sections to a strict 30-row vertical block: Section 1 → Row 10, Section 2 → Row 40, Section 3 → Row 70, etc.

## Table Schema (strict)
Every section must use exactly these 7 columns. Do not invent, rename, or reorder.

```json
{
  "A": "Metric / Line Item",
  "B": "Value — live numerical formula referencing FStmts., must evaluate to a number",
  "C": "Selected Period",
  "D": "Strength & Weakness — Excel IF statement on Column B returning \"Company's Strength\" or \"Company's Weakness\"",
  "E": "Chart Category Label",
  "F": "Chart Dynamic Data Source — paired values for 2020 and 2021",
  "G": "Analysis — 1-3 word descriptor of the 2020→2021 movement of the Column F value"
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
- Summarize the directional movement of the Column F value from 2020 to 2021.
- Maximum **3 words**. Use a *direction + magnitude* pattern.
- **Approved style:** `adequate increase`, `stable performance`, `slow growth`, `sharp decline`, `moderate recovery`, `flat trend`, `minor improvement`.
- **Forbidden:** full sentences, percentages, narrative explanations, repeating the metric name.
- Apply Column G to every row that holds paired 2020 and 2021 values.

## Chart Data Mapping
- **Section 1:** 3D Bar Chart comparing strictly *Total Current Assets* vs. *Total Current Liabilities*. Exclude ratio rows from the chart data range.
- **Sections 2–9:** Select only 2 to 4 comparable rows from Column F per chart. Ensure axis scale is visually meaningful.
- Anchor every chart inside **Columns J through P**. Charts must not overlap Columns A through I.

## Execution Workflow
1. Map exact row/column coordinates for all required line items on the `FStmts.` sheet for **both** the 2020 and 2021 period columns.
2. Write financial formulas into the Column A–G grid for all 9 sections following the **Table Schema**.
3. For every row holding paired 2020/2021 values, compute the directional movement and write a 1–3 word descriptor into Column G per **Analysis Column Rules**.
4. Create a Data Validation dropdown in Cell **B2** listing the reporting years. Construct `INDEX/MATCH` formulas in a hidden column as the dynamic data source driven by Cell B2.
5. Instantiate exactly **nine** 3D chart objects anchored in Columns J through P. Bind data sources per **Chart Data Mapping**.
6. Save the workbook under a **new filename**. Do not overwrite the original.

## Contrastive Examples

<anti_pattern>

```python
ws["B5"] = '=FORMULATEXT(FStmts.!C12)'           # produces #NAME?
ws["D5"] = "Formula Text"                          # invented column
ws["G5"] = "The current ratio increased significantly indicating better liquidity"  # sentence
ws["G6"] = "Up 12.4% YoY"                          # percentage, banned
ws["A34"] = "Source coordinate basis"              # padding row, banned
chart_anchor = "H10"                                # overlaps buffer zone
chart.add_data(Reference(ws, min_col=2, min_row=10, max_row=18))  # mixes ratios with currency
```

</anti_pattern>

<required_pattern>

```python
ws["B5"] = "='FStmts.'!C12/'FStmts.'!C25"
ws["D5"] = '=IF(B5>1.5, "Company\'s Strength", "Company\'s Weakness")'
ws["G5"] = "adequate increase"
ws["G6"] = "stable performance"
ws["G7"] = "slow growth"
chart_anchor = "J10"                                # inside chart zone
chart.add_data(Reference(ws, min_col=6, min_row=15, max_row=16))  # comparable rows only
```

</required_pattern>

## Final Output Requirement
Return a single executable Python script. Do not include explanatory prose outside the code block. Do not summarize the steps you took.
