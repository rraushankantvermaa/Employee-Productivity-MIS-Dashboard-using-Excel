# 📊 Employee Productivity MIS Dashboard

An end-to-end MIS (Management Information System) project built entirely in **Excel** — covering raw data intake, formula-driven data validation, cleaning, summary tables, charts, and an executive dashboard, using real MIS/analyst workflows rather than hardcoded numbers.

---

## 🗂️ Project Overview

This project analyzes **2,880 daily productivity records** across **20 employees**, **5 departments**, and **2 office locations** (Kolkata & Durgapur) over a 6-month period (April–September 2026).

The goal: turn raw attendance/task data into a decision-ready dashboard, the way an MIS analyst would deliver it to management — with every number traceable back to a formula, not a paste-value.

**File:** [`MIS_Employee_Productivity_Dashboard.xlsx`](./MIS_Employee_Productivity_Dashboard.xlsx)

---

## 📁 Workbook Structure (5 sheets)

| # | Sheet | Purpose |
|---|-------|---------|
| 1 | **Raw_Data** | Untouched source data exactly as received (2,880 rows × 20 columns) |
| 2 | **Checking_and_Manipulation** | Formula-driven data quality audit + documented manipulation log + before/after sample |
| 3 | **Cleaned_Data** | Fully formula-linked cleaned dataset — real Date/Time types, recomputed fields, helper columns |
| 4 | **Tables_and_Charts** | 7 summary tables (department, attendance, monthly trend, location, designation, employee ranking, priority × case-status matrix) + 6 native Excel charts |
| 5 | **Dashboard** | 8 KPI cards, 4 key charts, and a written insights summary — the executive view |

---

## 🔍 Data Quality Checks Performed

All checks run live via formulas against `Raw_Data` (not manually eyeballed):

- Blank Employee IDs / Names / Departments / Dates → `COUNTBLANK`
- Duplicate Employee ID + Date combinations → `SUMPRODUCT` + `COUNTIFS`
- Negative Working Hours / Pending Tasks → `COUNTIF`
- Tasks Completed > Tasks Assigned (logically impossible) → `SUMPRODUCT`
- Achievement % outside 0–100 range → `SUMPRODUCT`
- Working Hours outside a sane 0–12 range → `SUMPRODUCT`
- Date column stored as **text** instead of a real date → `ISNUMBER` check
- Login/Logout Time blanks cross-checked against Absent/Leave counts

**Result:** the source data was structurally clean (no blanks, no negatives, no duplicates). The real cleaning work was **type correction** — converting text-stored dates (`DD-MM-YYYY`) and times into genuine Excel Date/Time values so they could actually be used in trend charts and time filters.

---

## 🛠️ Manipulations Applied

| Issue | Fix |
|---|---|
| `Date` stored as text | Rebuilt with `DATE(VALUE(RIGHT(...)), VALUE(MID(...)), VALUE(LEFT(...)))` — locale-safe, avoids `DATEVALUE` ambiguity |
| `Login Time` / `Logout Time` stored as text | Converted with `IF(txt="","",TIMEVALUE(txt))`, blanks preserved intentionally for Absent/Leave |
| `Pending Tasks` | Recomputed independently (`Tasks Assigned − Tasks Completed`) to cross-verify against source — 0 mismatches |
| `Achievement %` | Recomputed independently (`IFERROR(Completed/Assigned*100, 0)`) — 0 mismatches |
| `Achievement %` vs `Productivity %` | Found to be **100% identical duplicate columns** in the source — both retained, but flagged transparently rather than silently dropped |
| Helper columns | Added `Month` and `Day of Week` (derived from the cleaned Date) to support pivoting and the monthly trend chart |

---

## 📈 Key Insights

- Overall attendance rate is strong (**~97%**); **HR** department has the lowest average achievement (**78.1%**) — worth a closer look.
- **Customer Support** and **Operations** lead on average achievement %, both above the company-wide average.
- Only **Low-priority** cases are ever marked *Closed* in this period — Critical/High/Medium priority cases stay In Progress, Pending, or Backlog, pointing to a potential process bottleneck.
- **Durgapur** tracks closely with **Kolkata** on achievement %, showing consistent performance across both sites.

---

## 🧰 Skills Demonstrated

- Data validation & auditing with `COUNTBLANK`, `COUNTIF(S)`, `SUMPRODUCT`
- Text-to-Date/Time type conversion without locale-dependent functions
- Cross-verification of derived metrics instead of trusting source values blindly
- `SUMIFS` / `AVERAGEIFS` / `COUNTIFS` for multi-dimensional summary tables
- `INDEX` / `MATCH` / `LARGE` / `SMALL` with a tie-break helper column for accurate Top-N / Bottom-N ranking
- Native Excel charting (bar, line, pie) linked live to formula-driven tables
- KPI-card style executive dashboard design

---

## 📌 How to Use

1. Download `MIS_Employee_Productivity_Dashboard.xlsx`
2. Open in Excel (or LibreOffice Calc)
3. Start on the **Dashboard** sheet for the executive summary, or **Checking_and_Manipulation** to see the full data-quality audit trail
4. All values are formulas — change any row in `Raw_Data` and the entire workbook recalculates end-to-end

---

## 👤 About This Project

Built as a portfolio project demonstrating MIS/data-analyst workflow: raw data → validation → cleaning → aggregation → visualization, entirely within Excel using formulas (no manual pasted values, no dynamic-array functions requiring modern Excel).
