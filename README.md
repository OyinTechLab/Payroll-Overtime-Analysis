# January Payroll: Where Overtime Costs Are Really Coming From

An end-to-end data analysis project — from a messy raw payroll export in Excel to a business-ready Power BI dashboard and a set of actionable recommendations.

**Tools used:** Excel · Power Query · Power BI (Data Modeling, DAX) · Data Storytelling

---

## 1. Introduction

Overtime is often treated as a routine payroll line item — a small, unavoidable cost of doing business. But overtime spend rarely distributes evenly across a workforce, and when it doesn't, it can quietly reveal a scheduling problem that costs more than it appears to on the surface.

This project starts with a single, unprocessed month of employee payroll data and ends with a dashboard built to answer one business question:

> **Is overtime spend at this company a sign of healthy flexibility, or a symptom of understaffing that's costing more than proper headcount planning would?**

This document walks through the full process — the raw data as received, every cleaning decision made in Excel, the modeling and analysis done in Power BI, and the recommendations that came out the other end.

---

## 2. About the Dataset

The source file was a single-month (January) payroll extract for **17 employees**, delivered as a typical payroll-software export: a wide cross-tab with:

- Employee name (last, first)
- Hourly wage
- Five **repeating weekly column blocks** — one set each for hours worked, overtime hours, regular pay, overtime bonus, and total weekly pay (one column per week, per metric)
- Summary rows (max, min, average, total) embedded directly in the same sheet, mixed in with employee records

This is a completely normal way for payroll systems to export data — and completely unusable for BI tooling as-is. A wide, repeating-column layout can't be filtered, sliced, related to a date table, or aggregated properly by a tool like Power BI, which expects one clean row per observation.

---

## 3. Working in Excel: From Raw Export to Clean Data

This was the first and, in some ways, most important stage — the dashboard is only as reliable as the structure underneath it.

### 3.1 Diagnosing the problem
Before touching anything, I identified three structural issues that would break any BI tool downstream:
- **No unique identifier** — employees were only distinguishable by first/last name, which is unsafe for relationships (duplicate names, typos, inconsistent spelling)
- **Wide format instead of long format** — five weeks of data spread across repeating columns rather than stacked as rows
- **Hardcoded summary statistics** embedded in the data itself (max/min/average/total rows), which would double-count or corrupt any aggregation done downstream

### 3.2 Reshaping the data
- **Unpivoted** the wide weekly columns into a long/tidy structure — one row per employee, per week (17 employees × 5 weeks = 85 rows), so each row is a single clean observation
- **Added an `Employee ID`** as a proper unique key, since names alone weren't a safe basis for relationships or filtering
- **Added a `Full Name`** field (concatenated from first/last) for clean display in visuals
- **Converted `Week Ending Date` into a true date type** (not text) — this one decision is what made proper time-based analysis possible later in Power BI
- **Added a `Month` label** for clarity, even though the dataset only spans one month, to make the structure extensible if more months are added later
- **Removed the embedded summary rows** (max, min, average, total) entirely — these are recalculated live in Power BI via DAX instead of being hardcoded, so they stay accurate if the underlying data ever changes

### 3.3 Formatting for reliability
- Structured the cleaned data as a formal **Excel Table** (not just a range) — this makes the dataset far more reliable to import, since Power BI reads table boundaries automatically rather than guessing where data starts/stops
- Applied consistent number formatting (currency, whole-number hours, proper date formatting) so data types would be interpreted correctly on import rather than defaulting to text

**Outcome:** a 12-column, 85-row tidy table — `Employee ID, Last Name, First Name, Full Name, Hourly Wage, Week Ending Date, Month, Hours Worked, Overtime Hours, Regular Pay, Overtime Bonus, Total Weekly Pay` — ready to import cleanly into Power BI.

---

## 4. Working in Power BI: Modeling & Analysis

### 4.1 Import and data type verification
On import, I re-verified every column's data type rather than trusting Power BI's auto-detection — dates as `Date`, currency fields as `Decimal Number`, hours as `Whole Number`. This step catches a common silent failure: a date or currency field imported as text will not filter, sort, or aggregate correctly, and the error often doesn't show up until a chart looks subtly wrong.

### 4.2 Building a proper date table
Even with only one month of data, I built a standalone `DateTable` using:

```
DateTable = CALENDAR(DATE(2026,1,1), DATE(2026,1,31))
```

and related it to the payroll data on `Week Ending Date` (one-to-many, single direction). This is standard BI modeling practice — it decouples "time" from the fact table, enabling clean time-based filtering and making the model extensible if more months of payroll are added later.

### 4.3 Core DAX measures
Rather than relying on raw column aggregation inside visuals, I built explicit measures so every number in the report is transparent, reusable, and auditable:

```
Total Payroll = SUM('PayrollData'[Total Weekly Pay])
Total OT Cost = SUM('PayrollData'[Overtime Bonus])
OT % of Payroll = DIVIDE([Total OT Cost], [Total Payroll], 0)
OT Dependency Ratio =
    DIVIDE(SUM('PayrollData'[Overtime Bonus]), SUM('PayrollData'[Regular Pay]), 0)
Total OT Hours = SUM('PayrollData'[Overtime Hours])
Avg OT Hours per Employee =
    DIVIDE([Total OT Hours], DISTINCTCOUNT('PayrollData'[Employee ID]), 0)
Headcount = DISTINCTCOUNT('PayrollData'[Employee ID])
```

`DIVIDE()` is used throughout instead of the `/` operator — it returns `0` instead of throwing an error when a filter selection produces no denominator (e.g. a slicer selection with no matching rows), which matters in a report other people will click around in.

### 4.4 Designing the dashboard — an iterative process
The dashboard went through several real rounds of correction before reaching its final state, which is worth documenting honestly rather than just showing the polished result:

- **Fixed misleading scale issues** — an early chart plotted Overtime Hours and Overtime Cost on the same axis; since cost values (in the thousands) dwarfed hour values (in the hundreds), the cost bars were rendered nearly invisible. Fixed by moving Overtime Cost to a secondary axis.
- **Removed a redundant chart** — an early version compared Regular Pay to Total Weekly Pay per employee; since Total Weekly Pay is just Regular Pay plus a small overtime addition, the two bars always looked nearly identical and added no insight. Replaced with the Overtime Dependency Ratio chart, which reveals a genuinely different pattern.
- **Corrected a broken scatter chart** — an early version of the wage-vs-overtime scatter collapsed all 17 employees into a single averaged point because no `Legend` field was assigned. Fixed by adding `Full Name` to the legend so each employee renders as an individual point.
- **Added a weekly trend line** — the dashboard initially had no time dimension at all, meaning it could show *who* was affected by overtime but not *when*. Adding a `Week Ending Date` line chart closed this gap and is what surfaced the mid-month spike.
- **Added conditional color formatting** on the dependency ranking chart, tied to hourly wage, so the wage-to-dependency relationship is visible at a glance rather than requiring a hover-to-check tooltip.

---

## 5. Explaining the KPIs

| KPI | What it measures | Why it matters |
|---|---|---|
| **Total Payroll** | Sum of all pay (regular + overtime) for the month | The baseline cost figure everything else is measured against |
| **Total Overtime Cost** | Sum of overtime bonus pay only | Isolates the cost specifically attributable to overtime |
| **OT % of Payroll** | Overtime cost as a share of total payroll | Turns a raw currency figure into a comparable, benchmarkable rate |
| **Headcount** | Distinct count of employees | Context for whether cost changes are about pay rates or staffing levels |
| **OT Dependency Ratio** | Overtime bonus as a % of an employee's regular pay | Identifies who is financially reliant on overtime, not just who logs the most hours |

---

## 6. Analytical Questions & Insights

**Q1 — How much of total payroll is overtime, and is that reasonable?**
Overtime accounted for **≈4.65% of total January payroll** (₦5,960 of ₦128,125). On its own, that's a modest figure — the real story is in how unevenly it's distributed.

**Q2 — Does overtime affect everyone equally, or is it concentrated?**
No. The top 5 employees by overtime hours (Mirabel Johnson, Tawa Muftau, Ramat Wale, Tolu Jamiu, Esther Patrick) account for a disproportionate share of total overtime cost, while several employees logged only 1–2 overtime hours all month.

**Q3 — Is overtime reliance related to base wage?**
Yes — this is the central finding. **Ramat Wale earns one of the lowest hourly wages in the company (≈₦10/hr) but ranks 4th in total monthly pay (₦9,233.40)**, driven almost entirely by 25 hours of overtime. Meanwhile, the highest hourly earner (₦45/hr) logs comparatively little overtime. Hourly wage alone is a poor predictor of who earns the most in a given month — overtime hours are the stronger predictor.

**Q4 — Is overtime a chronic pattern or a one-off spike?**
Weekly totals show overtime **more than doubled in weeks 2 and 3** (64 and 67 hours company-wide) compared to week 1 (25 hours), before returning to baseline in weeks 4–5 (33 hours each). This points to a specific mid-month event or staffing gap, not a constant planned level of overtime.

---

## 7. Key Findings

1. Overtime made up **~4.65%** of total January payroll.
2. Overtime cost is **concentrated among a small subset of employees**, not evenly spread.
3. **Low hourly-wage employees show the highest overtime dependency** — some of the lowest-paid staff by the hour end up among the highest-paid by the month, entirely through overtime.
4. Overtime volume **spiked mid-month** (weeks 2–3) and returned to baseline afterward, suggesting a temporary staffing or scheduling gap rather than a structural pattern.

---

## 8. Recommendations

Based on the findings above, three concrete actions follow:

1. **Investigate the weeks 2–3 spike operationally.** Since the increase is sharp and temporary rather than gradual, it's likely tied to a specific cause — a call-out, a scheduling gap, or a short-term demand increase. Identifying the cause would clarify whether it's a one-time event or a recurring seasonal pattern worth planning for in advance.
2. **Review scheduling for the highest-dependency, lowest-wage employees.** Employees like Ramat Wale are effectively being paid a de facto higher "blended" wage through overtime rather than a planned rate. It's worth assessing whether adding hours to their base schedule, or hiring additional part-time coverage, would be more cost-predictable than continued reliance on overtime.
3. **Set an overtime-dependency threshold to monitor going forward.** Rather than only tracking total overtime cost, tracking the *dependency ratio* per employee would flag early when someone's pay is becoming overtime-reliant — a leading indicator, not just a lagging cost report.

---

## 9. Explaining the Dashboard

The report is built as a single, story-driven page, structured to be read top to bottom:

- **Headline statement**: states the core finding in plain language before any chart is read
- **KPI cards**: Total Payroll, Headcount, Total Overtime Cost — the baseline numbers
- **Overtime Hours vs Overtime Cost by Employee**: ranks employees by overtime volume and cost side by side
- **Overtime Dependency by Employee**: ranks employees by what share of their pay comes from overtime — the chart that surfaces the wage-vs-reliance pattern
- **Wage vs. Overtime Hours (scatter)**: plots every employee by hourly wage against overtime hours, visually clustering the low-wage/high-OT group
- **Total Hours Worked vs Overtime Hours**: separates total workload from the overtime portion specifically
- **Weekly trend line**: shows overtime and payroll movement across the five weeks, surfacing the mid-month spike
- **Closing narrative**: the "so what" — translating the pattern into the recommendations above

---

## 10. Skills Demonstrated

- **Excel**: diagnosing structural data problems, unpivoting wide-format data, building unique keys, formal table structuring
- **Power BI data modeling**: date tables, relationships, type verification
- **DAX**: aggregation measures, safe ratio calculations (`DIVIDE`), distinct counts
- **Dashboard design & data storytelling**: headline-first structure, iterative chart correction, deliberate visual choices over default chart types
- **Business analysis**: translating a data pattern into specific, actionable recommendations

---

## 11. Files in This Repository

```
├── README.md
├── data/
│   └── Employee_Payroll_PowerBI.xlsx     (cleaned, tidy dataset)
├── report/
│   └── Employee_Payroll_Report.pbix      (Power BI report file)
├── screenshots/
│   └── (dashboard images)
└── docs/
    └── Employee_Payroll_Case_Study.docx  (this write-up as a standalone document)
```
