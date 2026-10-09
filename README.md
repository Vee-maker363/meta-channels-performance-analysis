# Meta Channels Performance Dashboard

A data quality audit and 4-KPI analysis of Facebook and Instagram performance data (January – April 2023), built end-to-end in Excel and Power BI from stacked, messy platform exports to a self-contained, interactive report.

## Overview

This project analyses Facebook Ads Manager and Page Insights exports for a four-month reporting window. The brief was scoped to 4 specific business questions rather than a general exploration, and that shaped every decision in the cleaning and analysis: fix what the KPIs need, document what they do not, and be explicit about what the data can and cannot support.

The exports arrived in a layout that is easy to misread: several tables stacked in a single sheet, summary blocks sitting above the daily data, repeated headers, and Facebook and Instagram series separated only by a blank row. There were also stretches of zero values in the ad data that could have been mistaken for a tracking fault. All of it is documented rather than fixed silently; see the audit trail in `Task2_Issues_Log.xlsx`.

## Why This Project Matters

Platform exports look authoritative, and a four-month trend with a sharp dip and a sharp recovery invites a tidy story: spend fell, reach fell, spend came back, reach came back. The harder skill is deciding how much of that story the data actually supports. This project treats the dip as something to investigate before explaining, and the final conclusions are deliberately narrower than the first draft of them.

The project also documents its own reasoning as it went, including several places where an earlier assumption or finding was found to be incomplete or inconsistent and corrected: a currency placeholder that disagreed with the chart formatting, a chart axis that exaggerated the dip, a supporting argument that turned out not to be independent evidence, and report wording that went further than the numbers. The process log keeps that history rather than quietly overwriting it.

## The 4 KPIs

1. Growth rate of Facebook and Instagram audience reach
2. Comparison of Facebook Page likes and Instagram followers
3. Trend of total ad spend
4. Percentage increase or decrease in ad spend

## Methodology

**Data quality audit (Excel, C.L.E.A.N.).** The data was reviewed against the C.L.E.A.N. framework: Consistency, Legitimacy, Exhaustiveness, Alignment and Notation. Each export was profiled for structure, completeness, duplicate dates and unusual values, and every issue was logged with its affected file, business impact, resolution and status.

**Structure and completeness.** Each raw export stacks multiple tables in one sheet. They were split into one clean table per metric, giving 120 daily rows each for 1 January to 30 April 2023, with no gaps and no duplicate dates. Row counts were checked against the expected 120 before anything was built on top of them.

**Legitimacy of the zero-value days.** Paid reach and ad spend fall to exactly zero on 7–8 March, taper down beforehand and recover afterwards, and spend then runs at roughly £2 a day on 14–21 March. A sudden drop-and-snap-back would suggest a logging fault; a gradual shape suggests something real happened to ad delivery. The values were therefore kept as recorded, with no imputation, and the cause is flagged as unconfirmed rather than assumed. A budget cut, a pause and an exhausted budget would all look alike in this data.

**Decisions made explicit.** Two ambiguities in the brief were resolved and documented rather than left implicit. Growth is measured month to month, built from the daily series summed by month. Spend is treated as GBP because the export states no currency, which is an assumption and is labelled as one.

**KPI calculation (Excel PivotTables).** Each KPI was built as a PivotTable on the cleaned tables, with month-on-month percentage changes calculated using `GETPIVOTDATA`, so every number in the report traces back to a formula. Expected values were predicted independently in pandas before the Excel build and then compared with it.

**Modelling and visualisation (Power BI).** The cleaning was rebuilt natively in Power Query (remove top rows, promote headers, keep the first 120 rows of each stacked block) so the `.pbix` is self-contained and does not depend on the Excel working copy staying in its cleaned state. The model, relationships and DAX measures are described below. Excel, Power BI and pandas produced identical values for all four KPIs.

**Review of the narrative.** Before the findings were written up, every figure was re-computed from the raw exports and each claim was checked against the numbers. Claims that went beyond the evidence were reworded, including a recovery described as triple-digit on both platforms (only Facebook was) and a rebound described as a return on spend. The reasoning for each change is recorded in the log.

## Data Model

![Power BI model view showing DateTable related to five fact tables](images/data-model.png)

The model is a star schema: one shared date dimension in the centre and five fact tables around it.

`DateTable` is a calculated table holding `Date`, `Month` (for example "Jan") and `MonthNum` (1 to 4), covering 1 January to 30 April 2023. `Month` is sorted by `MonthNum` so months display in calendar order instead of alphabetically.

The five fact tables each hold one row per day: `Ad trends` (paid reach and spend), `FB_Reach`, `IG_Reach`, the Facebook Page likes table and the Instagram followers table. Every one joins to `DateTable[Date]` as a one-to-many relationship (the `1` on the date table, the `*` on the fact table) with filtering flowing in a single direction, from the date table to the facts. Choosing a month on the date table therefore filters all five fact tables together, which is what lets one month axis drive every visual.

The raw exports carry their own date columns, and splitting the stacked sheets left the Instagram reach table's column named `Date2`; the relationship uses whichever column holds the date. A separate `Measures` table keeps the DAX in one findable place instead of attaching it to any single fact table.

## DAX

### Date table

```DAX
DateTable =
ADDCOLUMNS(
    CALENDAR(DATE(2023, 1, 1), DATE(2023, 4, 30)),
    "Month", FORMAT([Date], "MMM"),
    "MonthNum", MONTH([Date])
)
```

### Base measures

Each base measure sums the daily values, so on a visual with `DateTable[Month]` it returns the monthly total.

```DAX
FB Reach       = SUM(FB_Reach[Facebook reach])
IG Reach       = SUM(IG_Reach[Instagram reach])
Ad Spend       = SUM('Ad trends'[AD_ACCOUNT,SPEND,COUNT])
FB Page Likes  = SUM(FB_Page_Likes[Facebook Page likes])
IG Followers   = SUM(IG_Followers[Instgram followers])
```

### Month-over-month % change

Rather than relying on built-in time intelligence functions, each measure looks up its own previous month explicitly using the month number, which keeps the logic transparent and easy to check against the Excel calculation. It reads the current month number, evaluates the base measure with the date table's filter removed and restricted to the month before, and divides the difference by that previous value. January returns blank because there is no earlier month.

```DAX
FB Reach MoM % =
VAR CurrMonth = MAX(DateTable[MonthNum])
VAR CurrVal   = [FB Reach]
VAR PrevVal =
    CALCULATE(
        [FB Reach],
        FILTER(ALL(DateTable), DateTable[MonthNum] = CurrMonth - 1)
    )
RETURN
    DIVIDE(CurrVal - PrevVal, PrevVal)
```

`IG Reach MoM %` and `Ad Spend MoM %` follow the same structure with their own base measure. `ALL(DateTable)` removes the month filter coming from the visual so the measure can look back one month, and `DIVIDE` returns blank instead of an error when the previous value is blank or zero.

## Key Findings (Snapshot)

- Facebook reach fell by more than half in March (−53.70%) and rebounded in April (+140.10%). Instagram followed the same shape on a smaller scale (−35.74%, then +32.47%). Reach moved closely with ad delivery, especially on Facebook, where daily reach correlates strongly with both paid reach and ad spend. Instagram reach responded far less.
- April's large percentages are largely a rebound from a depressed March base. Measured against February instead, Facebook reach is only about 11% higher, Instagram reach is about 15% lower, and ad spend is about 39% higher.
- Audience growth did not follow spend month to month. In February spend rose by more than a third while both Facebook likes and Instagram followers fell, so spend explains reach better than it explains audience growth.
- Facebook and Instagram finished the period close to parity on new audience (3,089 likes against 3,188 followers), but Instagram overtook Facebook from March and held up better through the dip, with followers down about a third from January to March against about two thirds for Facebook likes.
- Total ad spend was £4,959, with March the lowest month (£642) and April the highest (£1,920.40).

(Full detail and the reasoning behind each finding are in the Power BI report and the process documentation.)

## Tools Used

- **Microsoft Excel**: data quality audit, cleaning, PivotTable KPI calculations and `GETPIVOTDATA` month-on-month changes
- **Microsoft Power BI**: self-contained data model, DAX measures and the interactive report
- **Power Query (M)**: used so every structural fix is a reproducible step, not a manual one-off edit
- **Python (pandas)**: independent prediction and verification of every KPI value

## Reproducing This Project

To explore the interactive report, open `Task2.pbix` in Power BI Desktop (free). To see the KPI calculations, open `Reach(2).xlsx` and review the `Pivot_summary` sheet for the PivotTables, month-on-month columns and charts. Open it in Excel and press F9 once so the `GETPIVOTDATA` cells refresh.

To see the full audit trail, open `Task2_Issues_Log.xlsx`: it documents every data quality issue found and its status, how each KPI was scoped and checked, and a step-by-step process log of how the analysis was built, including the places where an earlier finding was revisited and corrected.

## Limitations & Notes

- Four monthly data points show association, not proof of cause. The cause of the March drop in ad delivery is unconfirmed and is the one open issue in the log.
- Monthly reach is the sum of daily unique users, so repeat viewers are counted again. Summed daily paid reach is 848,467, while the de-duplicated paid reach in the export's own summary block is 339,230. Growth rates use the same method in every month, so they are not affected, but the totals should not be read as the number of distinct people reached.
- Facebook Page reach includes paid reach (daily correlation of 0.99 with paid reach), so organic reach cannot be isolated from these exports.
- Currency is assumed to be GBP because the export states none.
- The exports contain no placement or format breakdown, so nothing here says which ad formats worked.
