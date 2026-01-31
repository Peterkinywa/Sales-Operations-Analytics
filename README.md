# Sales-Operations-Analytics
This project analyzes sales operations data for a multi-regional electronics distributor using Microsoft Excel. The objective was to clean, enrich, analyze and visualize the data using Excel formulas, PivotTables and interactive dashboards to generate actionable business insights.

# Part A - Data cleaning & preparation
## 1. Staging Table Creation
A dedicated staging sheet was created to preserve raw data integrity and apply all cleaning logic.

### Duplicate removal
- Criteria used - Selected all data (ctrl+A), clicked Remove Duplicates, selected all columns and hit OK

### Data type fixes
- Date fields (OrderDate, RequiredDate) converted to Date format
- Numeric fields (UnitPrice, UnitCost, Quantity, DiscountPct) converted to Number

### Missing value handling
Missing values were handled using Excel formulas in the staging table. Blank City field were replaced with "Unknown", Channel fields were replaced with "Unspecified" while missing Salesperson values were replaced with "Unassigned". This ensured no records were excluded from analysis while preserving transparency of data quality issues.

### Flagging and correcting suspicious values
- UnitPrice (negative prices) - All negative prices were converted to their absolute values (-30.66 converted to 30.66)
- Discounts (>30%) - Discount percentages exceeding 30% were capped at 30% using Excel formulas.

### Date validation
Rule: RequiredDate ≥ OrderDate. If RequiredDate < OrderDate, it was reset to OrderDate + 7 days.

RequiredDate values earlier than OrderDate were identified as data errors. To correct these, RequiredDate was reset to OrderDate plus 7 days. LeadTimeDays was then computed using the corrected RequiredDate. This ensured consistency with observed operational performance.

## Calculated metrics
| Metric       | Formula in excel                                                       |
| ------------ | -----------------------------------------------------------------------|
| GrossRevenue | `UnitPrice × Quantity × (1 − DiscountPct)`, `=N2*Q2*(1-P2)`            |
| CostOfGoods  | `UnitCost × Quantity`, `=L2*Q2`                                        |
| GrossProfit  | `GrossRevenue − CostOfGoods`, `=U2-V2`                                 |
| MarginPct    | `IF(GrossRevenue=0,0,GrossProfit/GrossRevenue)`, `=IF(U2=0, 0, W2/U2)` |

## Create standardized dimensions
- Month (MMM-YYYY) from OrderDate - `=TEXT(B2,"MMM-YYYY")`
- Quarter (e.g., Q1-2024) - `="Q"&ROUNDUP(MONTH(B2)/3,0)&"-"&YEAR(B2)`
- PriceBand - 
