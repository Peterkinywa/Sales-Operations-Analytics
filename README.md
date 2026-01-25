# Sales-Operations-Analytics
This project analyzes sales operations data for a multi-regional electronics distributor using Microsoft Excel. The objective was to clean, enrich, analyze and visualize the data using Excel formulas, PivotTables and interactive dashboards to generate actionable business insights.

# Part A - Data cleaning & preparation
## 1. Staging Table Creation
A dedicated staging sheet was created to preserve raw data integrity and apply all cleaning logic.

### Duplicate removal
- Criteria used - OrderID should be unique.
- Rule - Rows with duplicate OrderID values were considered exact duplicates and removed using Remove Duplicates on Excel. 

<img width="1312" height="455" alt="image" src="https://github.com/user-attachments/assets/e6c42c0b-1e00-4ccf-aa2e-4bb0b421f7b0" />

<img width="1312" height="420" alt="image" src="https://github.com/user-attachments/assets/a3a46c2d-82f0-4e00-b9d5-f760b4f21b26" />

### Data type fixes
- Date fields (OrderDate, RequiredDate) converted to Date format
- Numeric fields (UnitPrice, UnitCost, Quantity, DiscountPct) converted to Number

<img width="1328" height="389" alt="image" src="https://github.com/user-attachments/assets/072b6673-7bb7-46ff-8de4-ebf5bdb4c74f" />

<img width="1325" height="388" alt="image" src="https://github.com/user-attachments/assets/c714c307-9832-4a76-b137-f19f2d0e64fb" />

### Missing value handling
Missing values were handled using Excel formulas in the staging table. Blank City and Channel fields were replaced with "Unknown" while missing Salesperson values were replaced with "Unassigned". This ensured no records were excluded from analysis while preserving transparency of data quality issues.

`=IF(OR(F2=" ", ISBLANK(F2)), "Unknown", F2)`

`=IF(OR(H2=" ", ISBLANK(H2)), "Unknown", H2)`

`=IF(OR(I2=" ", ISBLANK(I2)), "Unassigned", I2)`

<img width="1421" height="237" alt="image" src="https://github.com/user-attachments/assets/e08b013c-e9a6-44d6-9c33-c00da424f2cb" />

<img width="1679" height="255" alt="image" src="https://github.com/user-attachments/assets/5e224a93-bb66-4afa-80b7-fa66d0b70f9a" />

<img width="1559" height="254" alt="image" src="https://github.com/user-attachments/assets/2438f206-0304-4f1b-a5a1-9180e5192b9e" />

### Flagging and correcting suspicious values
- UnitPrice (negative prices) - All negative prices were converted to their absolute values (-30.66 converted to 30.66)
  
  <img width="1503" height="160" alt="image" src="https://github.com/user-attachments/assets/7ad60782-b610-4db4-949d-4a4a2b1dc297" />

  <img width="1501" height="125" alt="image" src="https://github.com/user-attachments/assets/bf903d48-4700-4e6b-b9f5-99505335fe72" />

- Discounts (>30%) - Discount percentages exceeding 30% were capped at 30% using Excel formulas.

<img width="1433" height="238" alt="image" src="https://github.com/user-attachments/assets/0d596654-1cd1-4c9c-ae3e-bddab39157db" />

<img width="1437" height="197" alt="image" src="https://github.com/user-attachments/assets/5b84cbcb-c6a0-434d-ae9c-43c6f5b145de" />

### Date validation
Rule: RequiredDate ≥ OrderDate. If RequiredDate < OrderDate, it was reset to OrderDate + 15 days.

RequiredDate values earlier than OrderDate were identified as data errors. To correct these, RequiredDate was reset to OrderDate plus 15 days. The 15-day adjustment was derived from the average lead time calculated across all valid records, making the correction data-driven rather than assuming. LeadTimeDays was then computed using the corrected RequiredDate. This ensured consistency with observed operational performance.

<img width="1759" height="297" alt="image" src="https://github.com/user-attachments/assets/f55998b8-64a6-4bdf-bb95-3af768b741e4" />

<img width="1865" height="261" alt="image" src="https://github.com/user-attachments/assets/9a5e0de4-200f-45ed-ae16-a23c8b880dcf" />

<img width="1892" height="315" alt="image" src="https://github.com/user-attachments/assets/118ae7f4-cdf8-4fa5-baa6-36c7c7d1bcee" />

<img width="1767" height="226" alt="image" src="https://github.com/user-attachments/assets/8836b28c-e2b0-4733-a719-b8635c49c3ac" />

## Calculated metrics
| Metric       | Formula in excel                                                     |
| ------------ | -------------------------------------------------------------------- |
| GrossRevenue | `UnitPrice × Quantity × (1 − DiscountPct)`, `=M2*O2*(1-N2)`          |
| CostOfGoods  | `UnitCost × Quantity`, `=L2*O2`                                      |
| GrossProfit  | `GrossRevenue − CostOfGoods`, `=P2-Q2`                               |
| MarginPct    | `IF(GrossRevenue=0,0,GrossProfit/GrossRevenue)`, `=IF(R2=0,0,R2/P2)` |

<img width="1650" height="221" alt="image" src="https://github.com/user-attachments/assets/862a6b95-f1e3-4d5c-9a67-b3edf4a70ca8" />










  



