# Hospital Analysis Dashboard

An Excel-based data cleaning and dashboard project built on a fictional hospital patient dataset. It covers data cleaning, derived columns, pivot tables, KPIs and an interactive dashboard with slicers.

> **Note:** All data is fictional practice data. It does not describe real patients and must not be used for medical decisions.

![Dashboard](Hospital_Analysis_Dashboard.pdf)

---

## Project Files

| File | Description |
|------|-------------|
| `health_care_dataset.xlsx` | Raw dataset with intentional data-quality issues, 30 practice questions and cleaning notes |
| `Hospital_Analysis_Dashboard.xlsx` | Cleaned dataset, pivot tables, KPI sheet, answers and the final dashboard |
| `Hospital_Analysis_Dashboard.pdf` | Exported snapshot of the dashboard |
| `pandas_numpy_project.py` | Companion Pandas & NumPy practical (23 tasks) |

### Sheets in `Hospital_Analysis_Dashboard.xlsx`

| Sheet | Purpose |
|-------|---------|
| `dataset` | Cleaned data (40 records) plus derived columns |
| `pivot` / `Sheet7` | Pivot tables feeding the charts |
| `kpi` | KPI formulas (total patients, male/female, fee paid, follow-ups) |
| `answers` | Answers to the practice questions (LOOKUP, VLOOKUP, etc.) |
| `dashboard` | Final dashboard with KPI cards, charts and slicers |

---

## Dataset

**40 patient records, 14 original columns**

`Patient_ID`, `Patient_Name`, `Age`, `Gender`, `City`, `Visit_Date`, `Department`, `Diagnosis`, `Consultation_Fee`, `Amount_Paid`, `Payment_Status`, `Follow_Up_Required`, `Blood_Pressure`, `Blood_Group`

### Data cleaning performed

| Issue in raw data | Fix applied |
|-------------------|-------------|
| Negative age (-5), impossible age (150), one blank age | Corrected / replaced with valid values (final range: 19-73) |
| Inconsistent gender values (`Female`, `female`, `M`, `F`) | Standardised to `Male` / `Female` |
| Mixed date formats (`2026-01-05`, `05/01/2026`, `2026/01/07`, `Jan 09 2026`) | Converted to a single Excel date format |
| Inconsistent `Payment_Status` capitalisation (`Paid` / `paid`) | Standardised to `Paid`, `Pending`, `Unpaid` |
| Inconsistent text case in `Diagnosis`, blood group, BP prefix (`BP-120/80`) | Proper-cased and cleaned |

### Derived columns

| Column | Logic |
|--------|-------|
| `Total Bill` | `Consultation_Fee + Amount_Paid` |
| `Balance Due` | `Consultation_Fee - Amount_Paid` |
| `Fee Status` | `Expensive` if fee >= 1000, `Moderate` if fee > 500, otherwise `Cheap` |
| `systolic`, `diastolic` | Split from the `Blood_Pressure` text (e.g. `120/80`) |
| `Pressure Status` | `High` if systolic >= 140 or diastolic >= 90; `Normal` if systolic <= 120 and diastolic <= 80; otherwise `Elevated` |

---

## Dashboard

### KPI cards

| KPI | Value | Formula |
|-----|-------|---------|
| Total Patients | 40 | `COUNTA(Patient_ID)` |
| Male Patients | 20 | `COUNTIF(Gender, "Male")` |
| Female Patients | 20 | `COUNTIF(Gender, "Female")` |
| Total Fee Paid | 24,450 | `SUMIF(Payment_Status, "Paid", Total Bill)` |
| Follow-up Required | 9 | `COUNTIF(Follow_Up_Required, "Yes")` |

### Charts

1. **Gender Wise Diagnosis Count** - bar chart of diagnoses split by gender
2. **Patients Count By Months** - line chart of visits per month (Jan, Feb, May)
3. **Patient Count By Department** - column chart across 7 departments (General is highest with 12)
4. **Follow Up Required Count** - donut chart (77% No / 23% Yes)
5. **Patient Count By City** - pie chart across 9 Kerala cities
6. **Blood Pressure Status** - column chart (Elevated 16, High 9, Normal 15)

### Slicers (interactive filters)

`Fee Status` | `Age` | `Gender` | `Blood_Group` | `Payment_Status`

---

## Skills Demonstrated

- Data cleaning and standardisation in Excel
- Formulas: `IFS`, `IF`, `AND`/`OR`, `LEFT`, `MID`, `FIND`, `VALUE`, `SUMIF`, `COUNTIF`, `COUNTA`
- Lookup functions: `LOOKUP`, `VLOOKUP`
- Pivot tables and pivot charts
- Slicers and dashboard design
- Pandas & NumPy data analysis (see `pandas_numpy_project.py`)

## How to Use

1. Open `Hospital_Analysis_Dashboard.xlsx` in Microsoft Excel (2019 or later; slicers and `IFS` are required).
2. Go to the `dashboard` sheet and use the slicers on the left to filter all charts.
3. To refresh after editing `dataset`: **Data > Refresh All**.

To run the Python companion:

```bash
pip install pandas numpy openpyxl
python pandas_numpy_project.py
```

## Tools

Microsoft Excel (Pivot Tables, Slicers, Charts) | Python | Pandas | NumPy
