# 🏥 Transforming Secure Life General Hospital Patient Data into Actionable Insights
![](home.png)
[![Excel](https://img.shields.io/badge/Microsoft_Excel-Data_Analysis_&_Dashboard-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)]()
[![Domain](https://img.shields.io/badge/Domain-Healthcare_Analytics-blue?style=for-the-badge)]()
[![Contact](https://img.shields.io/badge/Email-elal.o.wilsons@gmail.com-red?style=for-the-badge&logo=gmail&logoColor=white)](mailto:elal.o.wilsons@gmail.com)

---

## 📋 Executive Summary

Secure Life General Hospital recorded **1,000 patient visits** between January 2024 and December 2025 and billed a total of **$8.05 million**. About **$3.48 million (43.3%)** of that amount belongs to bills that are still **Pending** or only **Partially Paid**. This report studies the hospital’s admission, billing, and insurance records and presents the results in an Excel dashboard so that management can see how the hospital is performing and where action is needed.

* **Organisation:** Secure Life General Hospital (6 clinical departments)
* **Role:** Data Analyst
* **Data Used:** 1,000 patient visits and 11 fields, with admissions from 1 January 2024 to 28 December 2025
* **Tools:** Microsoft Excel for data cleaning, modeling, statistical analysis, and interactive dashboarding
* **Deliverables:** An interactive Excel dashboard workbook (4 headline KPI cards and 7 visual charts) and this consolidated analytical report

### Key Headline Findings
* **Patients and Income Come From Different Places:** General treats the most patients (**252 visits**, 25.2%) but produces only **12.3%** of total billing. Conversely, Orthopedics and Cardiology treat **31%** of patients but produce **57.7%** of all billing ($4.65 million).
* **Unpaid Bills Are the Main Financial Risk:** Only **56.4%** of visits are fully paid. A further **25.7%** are pending and **17.9%** are partially paid.
* **Insurance Cover Correlates With Full Payment:** Where there is no insurance cover (0%), only **27%** of visits are fully paid. Where cover is 90% or more, **67%** are fully paid. *(Note: This is an observed pattern in the data and does not prove causation.)*
* **Longer Stays Mean Bigger Bills:** Length of stay and total bill amount move together very closely (**correlation = 0.89**), and Orthopedics has the longest average stay at **8.4 days**.
* **End of 2025 Needs Immediate Attention:** Pending bills rose to **33.8%** of visits admitted in Q4 2025, compared with **20.5%** in Q2 2025.

---

## 🏥 Project Background & Problem Statement

Secure Life General Hospital treats patients across six clinical departments: **Cardiology, Emergency, General, Neurology, Orthopedics, and Pediatrics**. Management observed shifts in patient admissions, a growing number of unpaid medical bills, and uneven insurance coverage across departments, lacking a single consolidated view of day-to-day operational and financial performance.

As the Data Analyst, my objective was to transform the raw admission, billing, and insurance dataset into actionable insights and build an interactive Excel dashboard to help management monitor performance, track payments, and make informed decisions.

> **Core Analytical Rule:** Each row in the dataset represents **one patient visit**, not necessarily one unique individual. A returning patient appears in multiple rows, so every count in this report reflects visit volume.

### The 10 Management Questions Addressed
1. How many patient visits has the hospital recorded?
2. Which departments receive the highest number of patients?
3. Which departments generate the highest total bills?
4. What is the average length of stay for patients across different departments?
5. What proportion of patient bills has been paid, remains pending, or has been partially paid?
6. How much has the hospital billed, and how much remains outstanding?
7. How does insurance coverage vary across departments and diagnoses?
8. Which age groups account for the highest number of patient visits?
9. How do admissions and billing amounts change over time?
10. Which areas may require management attention based on the analysis?

---

## 🗂️ About the Dataset & Data Quality Checks

The dataset comprises **1,000 patient visit records** and **11 core fields** with zero missing values.

### Data Quality Validation (4 Checks Passed)
* **Completeness:** 1,000 rows with zero missing values across all required fields.
* **Uniqueness:** All 1,000 Patient IDs (`PAT-3001` to `PAT-4000`) are unique visit identifiers. While 684 distinct patient names exist, names are not reliable unique person identifiers, so analysis strictly counts visits.
* **Date Logic:** Zero records show discharge before admission. Stays range from **1 to 14 days**, with admissions spanning **1 January 2024 to 28 December 2025** and the final discharge occurring on **3 January 2026**.
* **Clinical Consistency:** Every diagnosis belongs exclusively to one department. All Pediatrics patients are aged 17 or younger, and no other department treats patients under 18.

### Engineered Fields Created in Excel
Five calculated fields were built from the raw dataset to support dashboarding and financial analysis:
1. **Length of Stay (Days):** `Discharge Date - Admission Date`
2. **Insurance Covered ($):** `Total Bill Amount * Insurance Coverage (%)`
3. **Patient Share ($):** `Total Bill Amount - Insurance Covered ($)`
4. **Age Group:** Categorized into five bands (`Under 18`, `18 to 35`, `36 to 50`, `51 to 65`, `66 and above`)
5. **Admission Month & Quarter:** Extracted from `Admission Date` for time-series trend analysis

---

## ⚙️ Tools, Methodology & Key Definitions

### 4-Step Excel Workflow
1. **Extract & Validate:** Loaded raw records into Excel, verified structural integrity, paired alternating patient/billing blocks row-by-row, and confirmed diagnosis-to-department mapping.
2. **Prepare:** Built calculated columns for Length of Stay, Insurance Covered ($), Patient Share ($), Age Groups, and Date Periods.
3. **Analyze:** Used Excel Pivot Tables, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, and correlation formulas (`CORREL`) to compute all operational and financial metrics.
4. **Visualize:** Designed an interactive Excel Dashboard featuring 4 KPI headline cards, 7 dynamic charts, and slicers for Department, Admission Date, and Diagnosis.

### Key Terms & Measurement Notes
* **Weighted vs. Simple Insurance Coverage:** The **weighted coverage is 75.3%** (`$6.06M covered / $8.05M billed`) and is used for financial questions because high-value bills carry greater monetary weight. The **simple average is 74.9%** (displayed rounded as `75%` on the dashboard), treating every visit equally regardless of bill size.
* **Unsettled Billed Value vs. Outstanding Balance:** **$3.48 million** represents the *full billed value* of visits marked Pending or Partially Paid. Because the dataset does not record partial payment amounts already collected, $3.48M serves as the **upper limit** of true outstanding debt.
* **Mean vs. Median:** The mean bill is **$8,053**, while the median bill is **$5,296**—indicating that a smaller number of high-cost surgical/specialty stays pull the average upward.

---

## 📊 Detailed Analysis & Answers to Management Questions

### 1. Patient Visits Overview
* **Total Visits:** **1,000 visits** recorded between January 2024 and December 2025 (**512 female**, **488 male**).
* **Billing Distribution:** Average bill is **$8,053.14**, while the median bill is **$5,296.00**. Half of all hospital visits cost under $5,296.

### 2. & 3. Departmental Workload vs. Total Billing
General drives patient volume, whereas Orthopedics and Cardiology drive hospital revenue:

| Department | Patient Visits | Share of Visits | Total Billed ($) | Share of Billing | Average Bill ($) | Avg Length of Stay |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **General** | 252 | 25.2% | $993,026.92 | 12.3% | $3,941 | 3.9 days |
| **Emergency** | 190 | 19.0% | $622,525.28 | 7.7% | $3,276 | 2.0 days |
| **Orthopedics** | 159 | 15.9% | $2,553,862.32 | 31.7% | $16,062 | 8.4 days |
| **Pediatrics** | 153 | 15.3% | $572,987.96 | 7.1% | $3,745 | 3.0 days |
| **Cardiology** | 152 | 15.2% | $2,095,437.97 | 26.0% | $13,786 | 6.5 days |
| **Neurology** | 94 | 9.4% | $1,215,302.99 | 15.1% | $12,929 | 6.9 days |
| **Total / Overall** | **1,000** | **100.0%** | **$8,053,143.44** | **100.0%** | **$8,053** | **4.8 days** |

* **Key Takeaway:** An average Orthopedics visit ($16,062) or Cardiology visit ($13,786) generates roughly **4x more revenue** than a General ($3,941) or Emergency ($3,276) visit. A single late payment in Orthopedics or Cardiology carries far greater financial impact.

### 4. Average Length of Stay (LOS)
* **Hospital-Wide LOS:** **4.8 days** average (median: **4.0 days**).
* **Departmental Breakdown:** Orthopedics (**8.4 days**), Neurology (**6.9 days**), Cardiology (**6.5 days**), General (**3.9 days**), Pediatrics (**3.0 days**), and Emergency (**2.0 days**).
* **Stay-to-Bill Correlation ($r = 0.89$):** Length of stay and bill size rise together closely. Out of 101 visits lasting 10+ days, **66 were in Orthopedics**, 19 in Neurology, and 16 in Cardiology.

### 5. & 6. Payment Status & Unsettled Billed Value

| Payment Status | Number of Visits | Share of Visits | Total Billed Value ($) | Share of Billed Value |
| :--- | :---: | :---: | :---: | :---: |
| **Paid** | 564 | 56.4% | $4,568,527.32 | 56.7% |
| **Pending** | 257 | 25.7% | $2,082,786.44 | 25.9% |
| **Partially Paid** | 179 | 17.9% | $1,401,829.68 | 17.4% |
| **Total Unsettled (Pending + Partial)** | **436** | **43.6%** | **$3,484,616.12** | **43.3%** |

* **Departmental Unsettled Exposure:** Orthopedics holds the largest unsettled dollar value (**$1.10 million**), followed by Cardiology (**$897,000**), General (**$493,000**), Neurology (**$488,000**), Pediatrics (**$260,000**), and Emergency (**$251,000**).
* **Relative Collection Performance:** Measured against its own billing, **General** is the weakest performer with **49.6%** of its billed value unsettled and the lowest full-payment rate (**53.2%**). **Pediatrics** has the highest pending visit rate (**28.8%**), while **Emergency** has the strongest collection rate (**62.1%** fully paid).

### 7. Insurance Coverage Across Departments & Diagnoses
* **Overall Coverage:** Insurers cover **$6.06 million** (**75.3%** weighted average) of the $8.05 million billed, leaving **$1.99 million** as patient responsibility.
* **Departmental Uniformity:** Weighted coverage varies by less than 4.5 percentage points across departments, ranging from **72.7% in Pediatrics** to **77.1% in Cardiology**.
* **Coverage Tiers vs. Payment Completion:**
  * **0% Cover (44 visits):** Only **27%** fully paid (43% partially paid).
  * **50%–60% Cover (161 visits):** **39%** fully paid.
  * **70%–80% Cover (494 visits):** **59%** fully paid.
  * **90%–100% Cover (301 visits):** **67%** fully paid.
* **Diagnosis Highlights:**
  * *Lowest Weighted Coverage:* Asthma (**60.1%**) and Arrhythmia (**64.1%**).
  * *Highest Weighted Coverage:* RSV (**84.8%**), Coronary Artery Disease (**84.2%**), and Stroke (**84.1%**).
  * *Highest Unpaid Rates:* Neuropathy (**65%** not fully paid), Bronchitis (**61%**), and Dehydration (**61%**).
  * *Top 5 Revenue-Generating Diagnoses:* Carpal Tunnel ($589,740 across 37 visits), Atrial Fibrillation ($546,653 across 36 visits), Knee Replacement ($494,615 across 27 visits), Spinal Disc Herniation ($448,637 across 27 visits), and Hip Fracture ($410,608 across 24 visits). Notably, only **21%** of Hip Fracture bills remain unsettled compared to **57%** for Carpal Tunnel and **53%** for Atrial Fibrillation.

### 8. Patient Visits by Age Group

| Age Group | Patient Visits | Share of Visits | Average Bill ($) | Average Length of Stay |
| :--- | :---: | :---: | :---: | :---: |
| **51 to 65** | 235 | 23.5% | $10,028 | 5.4 days |
| **66 and above** | 232 | 23.2% | $8,715 | 5.1 days |
| **36 to 50** | 210 | 21.0% | $8,809 | 5.0 days |
| **18 to 35** | 170 | 17.0% | $7,364 | 4.5 days |
| **Under 18** | 153 | 15.3% | $3,745 | 3.0 days |

* **Demographic Insight:** Adults aged **51 and above** account for **46.7%** of all hospital visits. Patients aged 51–65 generate the highest average bill ($10,028) and longest stay (5.4 days). Under-18 visits reflect Pediatrics billing structure rather than an independent age effect.

### 9. Admissions & Billing Trends Over Time (2024 vs. 2025)
* **Year-over-Year Shift:** Visits dipped **3.1%** from **508 in 2024** to **492 in 2025**, while total billing fell **5.1%** from **$4.13 million** to **$3.92 million**.
* **Seasonality:** Across both years combined, **November (100 visits)** and **July (96 visits)** are the busiest months, while **February (66 visits)** and **June (69 visits)** are the quietest. Monthly volume ranged from a low of 29 visits (June & August 2025) to a peak of 57 visits ($510,000 billed) in November 2025.
* **Q4 2025 Payment Bottleneck:** Pending bills rose sharply to **33.8%** of visits admitted in Q4 2025, up from **20.5%** in Q2 2025 and **23.7%** in Q3 2025.

### 10. Priority Areas Requiring Management Attention

* **🔴 High Priority:**
  * **$3.48M Unsettled Billing:** 43.3% of total billed value is not yet fully settled, compounded by the Q4 2025 spike in pending bills (33.8%).
  * **Two-Department Revenue Dependence:** Orthopedics and Cardiology generate 57.7% of revenue, with Orthopedics alone holding $1.10M in unsettled bills.
  * **Low-Coverage Payment Risk:** Uninsured visits (0% cover) have only a 27% full payment rate versus 67% for patients with $\ge 90\%$ cover.
  * **Departmental Collection Gaps:** Pediatrics has the highest pending rate (28.8%), while General has the lowest paid share (53.2%) and highest unsettled value share (49.6%).
* **🟡 Medium Priority:**
  * **Extended Inpatient Stays:** Long stays in Orthopedics (8.4 days) and Neurology (6.9 days) drive high bills and increase revenue at risk.
  * **High-Value Invoices:** 85 bills of $20,000 or more represent **27% of total hospital revenue ($2.21M)**, and 35 of them are not fully paid.
* **🟢 Keep Watching:**
  * **2025 Volume Dip:** Slight year-over-year contraction in admissions (-3.1%) and billing (-5.1%).
  * **Underperforming Diagnoses:** Weak collection on Neuropathy, Bronchitis, and Dehydration; low coverage on Asthma and Arrhythmia.

---

## 📈 Excel Dashboard Overview

The interactive Excel dashboard consolidates all metrics onto one executive view, equipped with dynamic slicers for **Department**, **Admission Date**, and **Diagnosis**.

* **4 Headline KPI Cards:** `1,000 Total Visits` | `$8,053,143.44 Total Billed` | `4.81 Days Avg Length of Stay` | `75% Avg Insurance Coverage (Simple)`
* **7 Visual Charts:**
  1. Monthly Hospital Visit Trends (Jan 2024 – Dec 2025)
  2. Patient Visits by Department
  3. Total Billed Amount by Department
  4. Patient Visits by Age Group
  5. Hospital Payment Status Distribution (Paid / Pending / Partially Paid)
  6. Average Length of Stay by Department
  7. Average Insurance Coverage by Department *(Note: Axis starts at 72.5% to highlight inter-departmental variance, though actual spread is ~2.4 percentage points)*

---

## 🎯 Actionable Recommendations (Ranked by Financial Impact)

1. **Audit Aging Unpaid Bills:** Prioritize follow-up on the $3.48M unsettled balance, starting with Orthopedics and Cardiology invoices over $10,000 and the Q4 2025 backlog. *(Expected Benefit: Faster cash recovery)*
2. **Verify Insurance & Pre-Authorize at Admission:** Confirm coverage upfront, secure pre-authorization for high-cost procedures, and set up structured payment plans for patients with 0%–60% cover. *(Expected Benefit: Reduced unpaid patient shares)*
3. **Strengthen Discharge Billing in Pediatrics & General:** Issue clear itemized statements at discharge and automate payment reminders. *(Expected Benefit: Lower pending bill rates)*
4. **Standardize Care Pathways in Orthopedics & Neurology:** Implement structured discharge planning to optimize length of stay without compromising care quality. *(Expected Benefit: Lower inpatient cost & improved bed turnover)*
5. **Align Staffing & Elective Scheduling With Seasonality:** Scale clinical staffing for peak months (July & November) and schedule elective Orthopedics procedures during quieter months (February & June). *(Expected Benefit: Balanced workload & steadier monthly revenue)*
6. **Weekly Tracking of High-Value Bills ($\ge \$20,000$):** Assign dedicated billing oversight to the 85 bills over $20,000 until settled. *(Expected Benefit: Safeguards 27% of hospital revenue)*
7. **Upgrade Billing Data Capture:** Add fields for *Amount Paid to Date*, *Exact Balance Due*, *Invoice Date*, and *Insurance Claim Status* to track true receivables in real time. *(Expected Benefit: Exact monthly financial reporting)*
8. **Review Coding & Coverage for Weak Diagnoses:** Audit billing workflows and payer contracts for Neuropathy, Bronchitis, Dehydration, Asthma, and Arrhythmia. *(Expected Benefit: Targeted collection & coverage improvements)*

---

## ⚠️ Limitations & Lessons Learned

* **Upper-Limit Receivables:** Because partial payment amounts collected are not logged, the $3.48M unsettled figure represents the maximum potential unpaid balance.
* **Coverage vs. Claim Payout:** Recorded insurance coverage % reflects policy terms rather than verified insurer payouts.
* **Timing Bias in Q4 2025:** Newer admissions naturally have fewer settled bills; an accounts-receivable aging report is required to distinguish timing lag from collection deterioration.
* **Revenue vs. Profitability:** The dataset contains billing figures without operational cost data, so departmental comparisons reflect gross revenue rather than net profit.

---

## 📁 Suggested Repository Structure

```text
secure-life-hospital-analytics/
│
├── data/
│   ├── Hospital_Patients_Dataset.xlsx
│   └── HOSPITAL_DICTIONARY_DATA.pdf
│
├── dashboard/
│   └── Secure_Life_Hospital_Dashboard.xlsx
│
└── README.md
