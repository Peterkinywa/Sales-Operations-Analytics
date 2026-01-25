# Sales-Operations-Analytics
This project analyzes sales operations data for a multi-regional electronics distributor using Microsoft Excel. The objective was to clean, enrich, analyze and visualize the data using Excel formulas, PivotTables and interactive dashboards to generate actionable business insights.

# Part A - Data cleaning & preparation
## 1. Staging Table Creation
A dedicated staging sheet was created to preserve raw data integrity and apply all cleaning logic.

### Duplicate removal
- Criteria used - OrderID
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

<img width="1421" height="237" alt="image" src="https://github.com/user-attachments/assets/e08b013c-e9a6-44d6-9c33-c00da424f2cb" />

<img width="1679" height="255" alt="image" src="https://github.com/user-attachments/assets/5e224a93-bb66-4afa-80b7-fa66d0b70f9a" />

<img width="1559" height="254" alt="image" src="https://github.com/user-attachments/assets/2438f206-0304-4f1b-a5a1-9180e5192b9e" />





