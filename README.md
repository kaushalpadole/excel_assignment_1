# Excel Practical Assignment --- Superstore Sales Analysis

## Project Overview

This project applies Excel features to the Sample Superstore sales
dataset and a small number of focused practice tables. The goal is to
demonstrate how Excel can be used to clean, organize, analyze,
summarize, and present business data.

The assignment is organized into 12 sections, covering foundational
formulas, formatting, sorting and filtering, date calculations, logical
formulas, data cleaning, PivotTables, lookup formulas, conditional
aggregation, conditional formatting, report consolidation, and VBA
macros. The accompanying dashboard provides a visual overview of sales
performance, including category, product/sub-category, regional, and
yearly views.

## Project Objectives

-   Calculate and validate sales metrics using Excel formulas.
-   Organize records with sorting, filtering, and multi-level sorting.
-   Apply consistent number formats and conditional formatting to
    improve readability.
-   Clean text fields, identify duplicates and errors, and split or
    combine values where needed.
-   Summarize sales data with PivotTables and conditional aggregation.
-   Retrieve related information using VLOOKUP and HLOOKUP.
-   Consolidate reports across multiple worksheets.
-   Explore VBA macros to automate repetitive formatting and workbook
    tasks.
-   Communicate results through a concise sales dashboard and documented
    answers.

## Dataset

**Primary dataset:** Sample Superstore.

The workbook uses Superstore order/sales data for practical tasks such
as sorting records, cleaning customer/order fields, applying logical
tests to sales and profit, creating PivotTables, using lookup formulas,
and analyzing conditional totals.

Some exercises use small practice tables rather than the full Superstore
dataset where a specific feature is easier to demonstrate clearly.
Examples include monthly sales figures, a five-year revenue series,
timestamp/duration examples, and quarterly reports for consolidation.
These are practice examples and should not be interpreted as additional
records from the Superstore dataset.

## Requirements

### Functional requirements

1.  Calculate totals, minimums, maximums, ranked values, medians,
    counts, averages, and growth rates.
2.  Transpose tables manually and with the TRANSPOSE formula.
3.  Apply currency, percentage, custom number, date, and time formats.
4.  Sort and filter by numbers, dates, text, cell color, and multiple
    levels.
5.  Calculate elapsed time and extract date components.
6.  Classify records with IF, AND, OR, and IFERROR.
7.  Identify duplicate values, errors, blanks, and text-cleaning issues.
8.  Build PivotTables and group fields where appropriate.
9.  Retrieve values with VLOOKUP/HLOOKUP and aggregate data with
    COUNTIF(S), SUMIF(S), and AVERAGEIF.
10. Consolidate values from multiple worksheets and calculate subtotals
    for filtered data.
11. Create and run simple VBA macros, including a formatting macro and a
    workbook-open message.
12. Present the key business measures and trends in a readable
    dashboard.

### Software and tools

-   **Microsoft Excel:** data preparation, formulas, formatting,
    sorting/filtering, PivotTables, consolidation, dashboard visuals,
    and macro execution.
-   **Excel formulas/functions:** SUM, MAX, MIN, SMALL, LARGE, MEDIAN,
    TRANSPOSE, COUNT, COUNTA, COUNTBLANK, AVERAGE, IF, AND, OR, IFERROR,
    date functions, text-cleaning functions, VLOOKUP, HLOOKUP,
    COUNTIF(S), SUMIF(S), AVERAGEIF, SUBTOTAL, and POWER.
-   **PivotTables and PivotCharts/charts:** summarization and visual
    analysis.
-   **Conditional Formatting:** highlight performance bands, duplicates,
    and values meeting rules.
-   **VBA Editor / Developer tab:** create, edit, run, and assign
    macros.
-   **File formats:** `.xlsx` for regular workbooks and `.xlsm` for
    workbooks that contain VBA macros.

## Methods and Workflow

1.  **Prepare the data:** open the Superstore workbook and identify the
    relevant columns for each exercise.
2.  **Calculate:** use formulas to create metrics and derived columns,
    then validate results against the source data or Excel Status Bar.
3.  **Clean:** inspect duplicates, errors, blanks, inconsistent text,
    and values stored as text.
4.  **Organize:** sort and filter records, including custom and
    multi-level sorts.
5.  **Summarize:** use PivotTables and conditional aggregation to
    compare categories, regions, products, and time periods.
6.  **Consolidate:** combine quarterly practice reports and use SUBTOTAL
    to summarize visible rows after filtering.
7.  **Present:** format tables and build visuals that make performance
    patterns easier to interpret.
8.  **Automate:** use simple VBA macros for repetitive formatting and
    workbook actions.
9.  **Document:** record the purpose, formula or steps, result, and
    explanation for each assignment task.

## Challenges Addressed

-   **Different data types:** Excel treats numeric values, text, blanks,
    dates, and timestamps differently. Functions such as COUNT, COUNTA,
    COUNTBLANK, VALUE, and date formulas were used to demonstrate those
    distinctions.
-   **Static versus dynamic outputs:** manual transpose produces a
    static copy, while TRANSPOSE can stay linked to the source data.
-   **Sorting complete records safely:** sorting must include the full
    table so that values from different columns remain aligned with the
    correct record.
-   **Conditional formatting and sorting:** color rules must be applied
    consistently before sorting records by cell color.
-   **Date duration interpretation:** subtraction returns elapsed days
    including fractional days from timestamps, whereas DATEDIF returns
    completed calendar units.
-   **Lookup and error handling:** exact-match lookups and IFERROR help
    make missing or invalid lookup results easier to handle.
-   **Visible-row totals:** ordinary SUM includes filtered-out rows;
    SUBTOTAL can calculate totals or averages for visible data.
-   **Macro security and saving:** VBA macros require a macro-enabled
    workbook (`.xlsm`) and should only be enabled for files from trusted
    sources.

## Key Insights

-   **Monthly sales practice table:** total sales are ₹199,200. July has
    the highest monthly sales at ₹27,500; March has the lowest at
    ₹9,800; median sales are ₹15,300.
-   **Revenue growth practice table:** revenue increases from ₹100,000
    in 2021 to ₹180,000 in 2025. YoY growth is 20.00% in 2022, -4.17% in
    2023, 26.09% in 2024, and 24.14% in 2025. The revenue-based CAGR
    over the four growth periods is approximately 15.83%.
-   **Sorting practice table:** the highest amount is ₹75,000 and the
    lowest is ₹23,000. Color bands help distinguish high (₹30,000 and
    above), medium (₹15,000--₹29,999), and low (below ₹15,000) sales in
    that practice dataset.
-   **Date/time calculations:** elapsed duration and completed calendar
    days/months/years are not always the same because timestamps can
    include partial days.
-   **Data quality checks:** COUNT, COUNTA, and COUNTBLANK answer
    different questions. Using them together helps reveal whether a
    range contains numbers, non-empty cells, or blanks.
-   **Dashboard purpose:** the workbook's dashboard is designed to make
    sales totals and comparisons across product/category, region, and
    year easier to scan than raw rows alone.

The monthly sales, revenue-growth, sorting, and timestamp examples are
dedicated practice datasets used for demonstrating Excel features; they
are not business findings calculated from every row in the Superstore
dataset.

## Recommendations for Improvement

1. **Improve Product Profitability:** Identify products and categories generating high sales and profits. Investigate products with high sales but low or negative profit, and review their pricing and discount strategies.

2. **Optimize Discount Strategies:** Analyze the impact of discounts on profitability. Reduce excessive discounting where it significantly affects profit margins and focus on promotions that generate profitable sales.

3. **Optimize Inventory Management:** Use historical sales trends to forecast demand, maintain sufficient stock of popular products, and reduce excess inventory of slow-moving items.

4. **Strengthen Regional Sales Performance:** Compare sales and profitability across regions to identify growth opportunities. Develop targeted sales strategies and promotions for regions with weaker performance.

5. **Focus on Profitable Sales Growth:** Track sales growth alongside profit margins rather than focusing on revenue alone. Prioritize products and strategies that increase both revenue and profitability.

6. **Monitor Key Business KPIs:** Regularly track total sales, total profit, profit margin, order volume, average order value, and discount levels. Use monthly performance reviews to identify trends and support data-driven decisions.

**Expected Outcome:** These recommendations can help the store improve profitability, optimize inventory, strengthen regional performance, and make more informed business decisions.
**

## Limitations

-   The assignment demonstrates Excel techniques; it is not a formal
    audit of the Superstore dataset.
-   Some sections use small custom practice tables because they are
    better suited to demonstrating a specific formula or feature.
-   The dashboard provides a summary view; any business recommendation
    should be supported by checking the underlying records, filters,
    date range, and metric definitions.
-   Macro behavior can depend on Excel version and security settings.

## Deliverables

-   Excel workbook containing the assignment worksheets and analyses.
-   Dashboard summarizing selected sales metrics and comparisons.
-   This README documenting project purpose, requirements, tools,
    workflow, challenges, insights, and recommendations.


------------------------------------------------------------------------

**Project type:** Excel data analysis and reporting practical
assignment\
**Primary tool:** Microsoft Excel\
**Dataset:** Sample Superstore, with clearly labelled custom practice
tables for selected exercises
