# Sales-Operations-Analytics
This project analyzes sales operations data for a multi-regional electronics distributor using Microsoft Excel. The objective was to clean, enrich, analyze and visualize the data using Excel formulas, PivotTables and interactive dashboards to generate actionable business insights.

# Part A - Data Cleaning & Preparation
## 1. Staging Table Creation
A dedicated staging sheet was created to preserve raw data integrity and apply all cleaning logic.

### Duplicate Removal
Criteria used - OrderID
Rule - Rows with duplicate OrderID values were considered exact duplicates
Action - Kept the first occurrence and removed subsequent duplicates using Excel’s remove duplicates tool.
<img width="1915" height="1096" alt="image" src="https://github.com/user-attachments/assets/829d6320-3a28-4d60-9424-8dfb2f1165e6" />
