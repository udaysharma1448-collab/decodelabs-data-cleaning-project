# Data Cleaning & Preparation | DecodeLabs Project 1

Data cleaning and validation of an e-commerce orders dataset (1,200 records, 14 columns), done as part of the **DecodeLabs Industrial Training Kit | Data Analytics | Batch 2026**.

**Author:** Manish Chahar

## Problem Statement

Prepare a raw orders dataset for analysis by auditing it for missing values, duplicates, invalid dates, inconsistent text and broken business rules, and document every change in a change log.

## Dataset

| Item | Detail |
|---|---|
| Source file | `Dataset_for_Data_Analytics.xlsx` |
| Records (before / after) | 1,200 / 1,200 |
| Date range | 2023-01-01 to 2025-06-30 |
| Columns (14) | OrderID, Date, CustomerID, Product, Quantity, UnitPrice, ShippingAddress, PaymentMethod, OrderStatus, TrackingNumber, ItemsInCart, CouponCode, ReferralSource, TotalPrice |

## Tools Used

Python (pandas, openpyxl, reportlab), Microsoft Excel

## What Was Done

- **Missing values:** 309 blanks found, all in `CouponCode`. Filled with `No Coupon` (blank means no coupon was used), so no rows were lost.
- **Duplicates:** 0 duplicate OrderID, TrackingNumber or full rows.
- **Dates:** all 1,200 are valid dates; ISO format `YYYY-MM-DD` applied.
- **Numeric precision:** `UnitPrice` and `TotalPrice` set to 2 decimals, `Quantity` and `ItemsInCart` to whole numbers.
- **Text standardisation:** spaces trimmed and Proper Case applied; no spelling variants found.
- **Integrity rules:** ID formats, `TotalPrice = Quantity x UnitPrice` and `ItemsInCart >= Quantity` hold on every row.
- **Validation sheet:** live Excel formulas in `Cleaned_Dataset.xlsx` re-prove every check; the raw data is kept in a `Raw_Data` sheet for traceability.

## Verification Gate

| Requirement | Result |
|---|---|
| 0% error rate on unique identifiers | PASS |
| 0% error rate on date formats | PASS |
| No missing values left | PASS |

## Files in This Repository

| File | Description |
|---|---|
| `Dataset_for_Data_Analytics.xlsx` | Original raw dataset |
| `Cleaned_Dataset.xlsx` | Cleaned data with Validation and Raw_Data sheets |
| `data_cleaning_script.py` | Reproducible Python cleaning script |
| `Change_Log.pdf` | Detailed change log (CR001 to CR007) |

## How to Run

```bash
pip install pandas openpyxl reportlab
python data_cleaning_script.py
```

Keep the raw dataset in the same folder as the script when running it.

## Assumption

A blank `CouponCode` is treated as "no coupon used". If the business confirms blanks mean "data not captured", the label can be changed to `Unknown` with one edit in the script.

## Contact

Manish Chahar | manishchahar310@gmail.com
