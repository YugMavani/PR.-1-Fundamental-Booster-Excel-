# Project_1.xlsx – Excel Functions Practice Workbook

A practice workbook that demonstrates common Excel functions across three small datasets: student marks, product sales, and employee records. Each sheet has a data table (20 records) and a few lookup/helper blocks below it.

## Workbook Overview

| Sheet | Records | Purpose |
|---|---|---|
| `Student_Mark_Sheet` | 20 students | Averages, pass/fail logic, text functions, date calculations, lookups |
| `Seles_Details_Sheets` | 20 sales | Tiered discounts, rounding functions, SUMIFS, lookups, OFFSET/INDIRECT |
| `Employee_Data_Sheet` | 20 employees | Tenure calculation, INDEX+MATCH, XLOOKUP, INDIRECT |

> Sheet names are kept exactly as in the file (including the spelling "Seles"), so any formulas referencing them keep working.

---

## 1. Student_Mark_Sheet

**Data (A1:N21):** Student ID, Name, Math, Science, English, Enrollment Date, plus calculated columns.

| Column | Header | Formula / Logic |
|---|---|---|
| G | Average | `=AVERAGE(C2:E2)` |
| H | MATH & SCI Above 80 | `=IF((C2>80)*(D2>80),"Yes","No")` |
| I | First Name | `=LEFT(B2,FIND(" ",B2&" ")-1)` |
| J | Upparcase Name | `=UPPER(B2)` |
| K | Lowercase Name | `=LOWER(B2)` |
| L | Days Since Enrollment | `=IF(F2>TODAY(),"Future Enrollment",TODAY()-F2)` |
| M | Pass Threshold | Fixed value: `35` |
| N | Pass status | `=IF(G2>=M2,"Pass",Fail)` (array formula) |

**Summary / lookup block (rows 24–33):**

| Cell | Label | Formula |
|---|---|---|
| B24 | Math above 50 | `=COUNTIF(C2:C21,">50")` |
| B25 | All average (above 60) | `=COUNTIF(G2:G21,">60")` |
| B28 → B29 | Fetch student name by ID | `=IFERROR(VLOOKUP(B28,A2:E21,2,FALSE),"-")` |
| B32 | Students with average above 80 | `=FILTER(A2:A21,G2:G21>80," NO Student Found ")` (spills to B33) |

**Functions used:** AVERAGE, IF, LEFT, FIND, UPPER, LOWER, TODAY, COUNTIF, VLOOKUP, IFERROR, FILTER

---

## 2. Seles_Details_Sheets

**Data (A1:L21):** Sales ID, Product, Region, Salesperson, Amount, Date, plus calculated columns.

| Column | Header | Formula / Logic |
|---|---|---|
| G | Compute discount | Nested `IF` tiers (see table below) |
| H | Eligibility for discount | `=IF((E2>=10000)*(G2>0),"is eligible","is not eligible")` |
| I | Rounded Amount | `=ROUND($E2,-2)` |
| J | Ceiling Amount | `=CEILING(E2,100)` |
| K | Floor Amount | `=FLOOR(E2,100)` |
| L | mmm-yyyy | `=TEXT(F2,"mmm-yyyy")` |

**Discount tiers (column G):**

| Amount | Discount |
|---|---|
| ≥ 45,000 | 25% |
| ≥ 35,000 | 20% |
| ≥ 25,000 | 15% |
| ≥ 20,000 | 10% |
| ≥ 15,000 | 5% |
| ≥ 10,000 | 2.5% |
| < 10,000 | 0 |

**Lookup block (rows 24–47):**

| Cell | Purpose | Formula |
|---|---|---|
| B26 | Sum of sales by region + product | `=SUMIFS(E2:E21,C2:C21,B24,B2:B21,B25)` |
| B30 | Sales amount by Sales ID | `=IFERROR(VLOOKUP(B29,A2:E21,5,FALSE),"-")` |
| B35 | Amount by salesperson + month | `=IFERROR(INDEX($E$2:$E$21,MATCH(1,((D2:D21=B33)*(MONTH(F2:F21)=B34)),0)),"-")` (array formula) |
| B39 | Position of a product | `=XMATCH(B38,$B$2:$B$21,0)` |
| B43 | Sum of first N amounts | `=IFERROR(SUM(OFFSET($E$2,0,0,B42,1)),"-")` |
| B47 | Value of a cell typed as text | `=IFERROR(INDIRECT(B46),"-")` |

**Functions used:** IF (nested), ROUND, CEILING, FLOOR, TEXT, SUMIFS, VLOOKUP, INDEX, MATCH, MONTH, XMATCH, OFFSET, INDIRECT, IFERROR

---

## 3. Employee_Data_Sheet

**Data (A1:G21):** Employee ID, Name, Department, Salary, Joining Date, plus calculated columns.

| Column | Header | Formula / Logic |
|---|---|---|
| F | Joined Days Spend | `=IF(E2>TODAY(),"Future Joining",TODAY()-E2)` |
| G | mmm-yyyy | `=TEXT(E2,"mmm-yyyy")` |

**Lookup block (rows 24–34):**

| Cell | Purpose | Formula |
|---|---|---|
| B26 | Any detail by ID + column name | `=IFERROR(INDEX($B$2:$E$21,MATCH(B24,$A$2:$A$21,0),MATCH(B25,$B$1:$E$1,0)),"-")` |
| B30 | Salary by ID | `=XLOOKUP(B29,$A$2:$A$21,$D$2:$D$21,"-")` |
| B34 | Value of a cell typed as text | `=IFERROR(INDIRECT(B33),"-")` |

**Functions used:** IF, TODAY, TEXT, INDEX, MATCH, XLOOKUP, INDIRECT, IFERROR

---

## How to Use

1. Open `Project_1.xlsx` in **Excel 365 or Excel 2021+** (needed for `FILTER`, `XLOOKUP` and `XMATCH`).
2. Change the **input cells** in each lookup block to test the formulas:
   - `Student_Mark_Sheet`: B28 (student ID)
   - `Seles_Details_Sheets`: B24, B25 (region, product), B29 (Sales ID), B33, B34 (salesperson, month number), B38 (product), B42 (row count), B46 (cell address such as `E2`)
   - `Employee_Data_Sheet`: B24, B25 (employee ID, column name such as `Salary`), B29 (employee ID), B33 (cell address)
3. Results update automatically.

## Notes & Known Issues

- **`TODAY()` is volatile.** "Days Since Enrollment" and "Joined Days Spend" change every day. Dates later than today show "Future Enrollment" / "Future Joining".
- **Pass status formula:** `Student_Mark_Sheet!N2` uses `Fail` without quotes (`"Fail"` is intended). It works today because every student passes, but a student below the threshold would return `#NAME?`. Fix: `=IF(G2>=M2,"Pass","Fail")`.
- **Math above 50** (`B24`) is a count of students, not a sum or average.
- **Trailing spaces** exist in some text values (e.g. `Karan hujla `, `Vinu `), which can break exact-match lookups.
- **Names are not consistently capitalised** in the source data (e.g. `khushboo varma`, `heet Sharma`).
- **XLOOKUP, XMATCH and FILTER** will not work in Excel 2016/2019 or older.
- The `Seles_Details_Sheets` match on salesperson `vinu` (lowercase) works because Excel lookups are not case-sensitive, but the data value is `Vinu ` (with a trailing space), so `B35` may return `-`.
