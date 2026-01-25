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
Missing values were handled using Excel formulas in the staging table. Blank City field were replaced with "Unknown", Channel fields were replaced with "Unspecified" while missing Salesperson values were replaced with "Unassigned". This ensured no records were excluded from analysis while preserving transparency of data quality issues.

<img width="1316" height="210" alt="image" src="https://github.com/user-attachments/assets/d70b0f2d-7ee3-439a-a4ae-1d192548c5e1" />

<img width="1317" height="199" alt="image" src="https://github.com/user-attachments/assets/989ea60b-7894-4b8f-b42a-2f719f0b418c" />

<img width="1314" height="209" alt="image" src="https://github.com/user-attachments/assets/5c842c0f-fc7c-4b3d-b104-f821daa51710" />

### Flagging and correcting suspicious values
- UnitPrice (negative prices) - All negative prices were converted to their absolute values (-30.66 converted to 30.66)
  
<img width="1417" height="124" alt="image" src="https://github.com/user-attachments/assets/fcdc9d9c-c903-44e2-aa5d-ffe93fa3d057" />

- Discounts (>30%) - Discount percentages exceeding 30% were capped at 30% using Excel formulas.

<img width="1530" height="186" alt="image" src="https://github.com/user-attachments/assets/5b71aebd-fff2-4172-851a-58734877f2d7" />

### Date validation
Rule: RequiredDate ≥ OrderDate. If RequiredDate < OrderDate, it was reset to OrderDate + 7 days.

RequiredDate values earlier than OrderDate were identified as data errors. To correct these, RequiredDate was reset to OrderDate plus 7 days. LeadTimeDays was then computed using the corrected RequiredDate. This ensured consistency with observed operational performance.

<img width="1837" height="297" alt="image" src="https://github.com/user-attachments/assets/d32df47f-249b-4cc2-95fe-882359a2bdc2" />

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










  



