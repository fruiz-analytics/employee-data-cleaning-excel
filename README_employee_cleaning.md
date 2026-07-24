# Employee Data Cleaning & Integrity Analysis | Excel

**Data quality audit, cleaning, imputation, and analysis of an HR dataset**
Author: Franklin Manuel Ruiz Guadamuz · Tools: Advanced Excel (formulas, PivotTables, formula-based conditional formatting)

---

## 1. Overview
A raw employee dataset (1,020 records, 12 columns) was audited, cleaned, and analyzed end to end. Beyond standard cleaning, the project adds **missing-value imputation**, a **cross-field integrity check**, and a **summary analysis layer** — techniques a real HR analytics task would require.

The headline lesson isn't a dramatic finding; it's the opposite. An initial check appeared to flag 190 "impossible" records, but pressure-testing that result showed it was an artifact of how missing values were handled — not a real inconsistency. Catching that before reporting is the point.

## 2. Workbook structure
| Sheet | Purpose |
|---|---|
| `00_README` | Objective, scope, method |
| `01_Raw` | Original data, untouched |
| `02_Audit` | 20 checks with severity and decisions |
| `03_Clean_Formulas` | Cleaning with formulas (raw + clean side by side) |
| `04_Clean_Final` | Final deliverable (values only) |
| `06_Analysis` | Three PivotTables |

## 3. Cleaning highlights
- **Combined field split** — `Department_Region` ("DevOps-California") split into Department and Region (TEXTSPLIT / TEXTAFTER).
- **Regional date parsing** — US M/D/YYYY parsed explicitly and standardized to ISO YYYY-MM-DD.
- **Phone formatting** — removed a spurious leading minus sign from every number.
- **Salary** — regional decimal fix; 24 corrupt "Invalid" values flagged and blanked so the column stays numeric.
- **Category standardization** — Status, Performance, Remote_Work verified against clean value sets (TRIM + PROPER + UNIQUE).

## 4. Missing-value imputation (new technique)
211 records (~21%) had no Age. Rather than deleting them or labeling them "Missing", ages were **imputed with the average age of the employee's own department** (AVERAGEIF, rounded to a whole number). An `Age_Was_Imputed` column marks every estimated value, so downstream users always know which ages are real and which are inferred. This keeps the records usable for department- and salary-level analysis without silently inventing data.

## 5. Cross-field integrity check (the honest finding)
Tenure was calculated (YEARFRAC of join date to today) and compared against age, on the premise that no one starts working before ~16.

- A first pass flagged **190 records as impossible.**
- **Investigation** showed this was an artifact of zero-filling the missing ages: `tenure > 0 − 16` is always true, so every age-less row looked impossible.
- **Re-running the check against the original recorded ages produced zero real impossibilities** — employees with a recorded age are internally consistent (minimum recorded age 25, maximum tenure 6 years).

The genuine integrity issue is different and more subtle: **21% of records lack age**, and ages appear only in **5-year increments** (25/30/35/40) — evidence that Age was generated synthetically and independently of Join_Date. The takeaway: a striking result was verified before being reported, not taken at face value.

## 6. Analysis layer (PivotTables)
- **Headcount & average salary by Department × Status** — workforce distribution and pay overview.
- **Performance distribution by Region** — spread of performance ratings across regions.
- **Age availability by Department** — quantifies the missing-age problem per department.

A **formula-based conditional format** highlights salaries more than 35% above their own department's average — a within-group comparison the standard rules can't express.

## 7. Skills demonstrated
`TRIM` · `PROPER` · `IF` / nested `IF` · `AND` / `OR` · `IFERROR` · `LEN` · `TEXTSPLIT` / `TEXTAFTER` · `LEFT` / `MID` / `RIGHT` · `DATE` · `YEARFRAC` · `ISNUMBER` · `AVERAGEIF` (group imputation) · `COUNTIF` · `UNIQUE` · PivotTables · formula-based conditional formatting

## 8. Results
| Metric | Value |
|---|---|
| Records | 1,020 |
| Checks catalogued | 20 (1 critical, 2 high, 6 medium, 11 low) |
| Missing ages imputed (and flagged) | 211 |
| Corrupt salaries flagged | 24 |
| Real age-tenure impossibilities | 0 (initial 190 were an artifact) |
| Categories verified clean | Status, Performance, Remote_Work |

---
**Files:** `employee_cleaning_project.xlsx` (all sheets) · `Messy_Employee_dataset.csv` (source) · this README
