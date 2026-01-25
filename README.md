# Sales-Operations-Analytics
This project analyzes sales operations data for a multi-regional electronics distributor using Microsoft Excel. The objective was to clean, enrich, analyze and visualize the data using Excel formulas, PivotTables and interactive dashboards to generate actionable business insights.

# Part A - Data Cleaning & Preparation
## 1. Staging Table Creation
A dedicated staging sheet was created to preserve raw data integrity and apply all cleaning logic.

### Duplicate Removal
Criteria used - OrderID
Rule - Rows with duplicate OrderID values were considered exact duplicates
Action - Kept the first occurrence and removed subsequent duplicates using Excel’s remove duplicates tool.

<img width="1312" height="455" alt="image" src="https://github.com/user-attachments/assets/e6c42c0b-1e00-4ccf-aa2e-4bb0b421f7b0" />

<img width="1312" height="420" alt="image" src="https://github.com/user-attachments/assets/a3a46c2d-82f0-4e00-b9d5-f760b4f21b26" />

### Data Type Fixes
Date fields (OrderDate, RequiredDate) converted to Date format
Numeric fields (UnitPrice, UnitCost, Quantity, DiscountPct) converted to Number

<img width="1328" height="389" alt="image" src="https://github.com/user-attachments/assets/072b6673-7bb7-46ff-8de4-ebf5bdb4c74f" />

<img width="1325" height="388" alt="image" src="https://github.com/user-attachments/assets/c714c307-9832-4a76-b137-f19f2d0e64fb" />



