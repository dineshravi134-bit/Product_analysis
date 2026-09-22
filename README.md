# Excel Data Exploration & Analysis

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Data Analytics](https://img.shields.io/badge/Data_Analytics-Module_1-blue?style=for-the-badge)

## 📌 Project Overview
This repository contains the completed project for **Module 1 - Excel (Assignment 1: Data Exploration)**. 

The objective of this project is to perform fundamental data exploration, logical categorization, conditional aggregation, and string manipulation on a product dataset containing **34 product records** with attributes including `Product ID`, `Product Name`, `Brand Name`, `Quantity`, `Category`, and `Price ($)`.

---

## 🎯 Objectives & Tasks Completed

1. **Basic Data Exploration**
   - Summary statistics using `SUM`, `COUNT`, and `AVERAGE`.
2. **Extreme Values Identification**
   - Finding boundary values using `MIN` and `MAX`.
3. **Logical Categorization**
   - Applying conditional categorization (`IF`) based on custom pricing thresholds.
4. **Conditional Aggregations**
   - Segmented summary calculations using `SUMIF` and `COUNTIF`.
5. **Text Formatting & Parsing**
   - Extracting embedded metadata from structured string IDs using `LEFT`, `RIGHT`, and `MID`.

---

## 📊 Summary of Findings & Key Metrics

| Metric / Task | Result | Excel Formula Used |
| :--- | :---: | :--- |
| **Total Inventory Price** | `$10,100` | `=SUM(D2:D35)` |
| **Total Product Count** | `34` | `=COUNT(D2:D35)` |
| **Average Product Price** | `$297.06` | `=AVERAGE(D2:D35)` |
| **Minimum Price** | `$30` | `=MIN(D2:D35)` |
| **Maximum Price** | `$1,000` | `=MAX(D2:D35)` |
| **Total Price (Electronics)** | `$8,050` | `=SUMIF(F2:F35, "Electronics", D2:D35)` |
| **Products Under $100** | `11` | `=COUNTIF(D2:D35, "<100")` |

---

## 🛠️ Step-by-Step Implementation

### 1. Logical Function – IF (`Price Range`)
Categorized products into pricing tiers based on a `$500` threshold:
* **Rule:** If $\text{Price} \ge \$500$, classify as **High Price**, otherwise **Standard Price**.
* **Formula (Cell G2):**
  ```excel
  =IF(D2>=500, "High Price", "Standard Price")
