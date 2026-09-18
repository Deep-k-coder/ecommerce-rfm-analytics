# E-Commerce Customer Lifecycle & RFM Segmentation Dashboard

**Author:** Deep Koshiya  
**Role:** Analytics Engineer & Data Scientist  
**Date:** September 2026  
**License:** [MIT License](https://opensource.org/licenses/MIT)

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/)
[![scikit-learn](https://img.shields.io/badge/scikit--learn-F7931E?logo=scikit-learn&logoColor=white)](https://scikit-learn.org/)
[![Pandas](https://img.shields.io/badge/pandas-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![TailwindCSS](https://img.shields.io/badge/Tailwind_CSS-38B2AC?logo=tailwind-css&logoColor=white)](https://tailwindcss.com/)

---

## 1. Project Overview & Description
In modern multi-category digital commerce, treating an entire customer base as a homogenous entity leads to suboptimal marketing spend, inflated Customer Acquisition Costs (CAC), and severe customer churn blindspots. 

This project delivers a **production-grade analytics engineering and customer segmentation system**. The pipeline ingests unindexed transactional event streams, sanitizes data through robust statistical fences (including Tukey's IQR outlier isolation and credit note segregation), performs customer-level RFM feature engineering, executes unsupervised machine learning (K-Means), and applies heuristic quantile scoring (`pd.qcut`).

The refined data is transformed into a dimensional star-schema data mart and visualized through an interactive executive dashboard designed for commercial decision-makers and marketing teams.

```
[ Ingestion & Cleaning ] ──► [ Feature Engineering ] ──► [ Segmentation ML ] ──► [ BI Star Schema ]
  (Pandas / IQR Bounds)         (Vectorized RFM)           (K-Means & qcut)       (Tableau / HTML)
```

---

## 2. Dataset Reference & Link
- **Dataset Name:** Online Retail Transactional Dataset
- **Official Repository Link:** [UCI Machine Learning Repository: Online Retail Dataset](https://archive.ics.uci.edu/dataset/352/online+retail)
- **Alternative Mirror:** [Kaggle Online Retail II Dataset](https://www.kaggle.com/datasets/mashlyn/online-retail-ii-uci)
- **Dataset Description:** Contains transnational transactions occurring between 01/12/2010 and 09/12/2011 for a UK-based registered non-store online retail firm.
- **Attributes:**
  - `InvoiceNo`: 6-digit integral invoice number (cancellations prefixed with 'C')
  - `StockCode`: 5-digit integral product/item code
  - `Description`: Nominal product name
  - `Quantity`: Numeric quantity of each item per transaction
  - `InvoiceDate`: Timestamp of invoice generation
  - `UnitPrice`: Product unit price in sterling (£/$)
  - `CustomerID`: 5-digit integral customer identifier

---

## 3. Technologies & Dependencies Used
- **Core Analytics & Transformation:**
  - `Python 3.9+`
  - `pandas` (Vectorized aggregation, datetime parsing, quantile binning)
  - `numpy` (Logarithmic compression, array vectorization)
- **Machine Learning & Statistical Modeling:**
  - `scikit-learn` (StandardScaler, KMeans clustering, Silhouette Score analysis)
  - `scipy` (Statistical distributions)
- **Data Visualization & Reporting:**
  - `matplotlib` (Elbow curves, silhouette charts, spatial scatterplots)
  - `python-docx` (Automated generation of Microsoft Word project report)
  - `HTML5`, `Tailwind CSS`, `Chart.js` (Interactive executive web dashboard)
  - `Tableau / Power BI` (Dimensional star-schema compatibility)

---

## 4. Key Analytical Insights & Executive Takeaways
1. **Pareto Dynamic (Revenue Concentration):**
   - **27.4% of customer accounts ('Champions') generate 72.5% of gross revenue** ($47,250.31 out of $65,132.81).
   - High-spending repeat buyers require dedicated VIP concierge treatment rather than price discounts.
2. **At-Risk Revenue Exposure:**
   - Identified **$2,384.72 in revenue exposure** across 23 high-value accounts that have been inactive for >90 days.
   - Deploying automated win-back triggers at day 45 mitigates late-stage permanent churn.
3. **CAC Efficiency & Remarketing Optimization:**
   - The bottom **14.3% of customers ('Lost / Inactive')** contribute only 1.9% of total revenue with an average frequency of 1.0.
   - Suppressing these accounts from paid ad retargeting lists (Meta Ads, Google Ads) immediately eliminates budget waste.

---

## 5. Segment Taxonomy & Actionable CRM Playbooks

| Business Segment | Headcount (%) | Total Spend (%) | Avg Recency | Avg Frequency | Recommended Commercial Playbook |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **Champions** | 148 (27.4%) | $47,250.31 (72.5%) | 13.6 days | 18.4 | VIP concierge, early product access, non-discount loyalty gifts. |
| **Loyal Customers** | 95 (17.6%) | $6,346.33 (9.7%) | 37.7 days | 3.8 | Category cross-selling, threshold loyalty multipliers ($50 for 2x points). |
| **Hibernating** | 116 (21.5%) | $3,745.34 (5.8%) | 139.7 days | 1.9 | Low-cost automated re-engagement emails; prune if inactive after 60 days. |
| **Mid-Market Regulars** | 13 (2.4%) | $2,508.67 (3.9%) | 44.5 days | 11.3 | Standard promotional cadence, threshold-based free shipping incentives. |
| **At Risk (High Value)** | 23 (4.3%) | $2,384.72 (3.7%) | 92.8 days | 5.9 | Urgent direct outreach, aggressive limited-time winback vouchers ($15 off $50). |
| **Lost / Inactive** | 77 (14.3%) | $1,255.78 (1.9%) | 163.0 days | 1.0 | Exclude from paid acquisition remarketing to protect CAC efficiency. |
| **Needs Attention** | 33 (6.1%) | $846.77 (1.3%) | 47.2 days | 1.5 | Automated triggered emails featuring personalized category recommendations. |
| **Recent Converts** | 35 (6.5%) | $794.89 (1.2%) | 15.9 days | 1.3 | Onboarding nurture sequence with 10% coupon valid for 7 days. |

---

## 6. Project Directory & Deliverables Structure

```tree
ecommerce_rfm_analytics/
├── Deep_Koshiya_Ecommerce_RFM_Segmentation.ipynb  # Primary Jupyter Notebook submission
├── Deep_Koshiya_Ecommerce_RFM_Segmentation.py     # Standalone Python script submission
├── Deep_Koshiya_ProjectReport.docx               # Formatted Word Project Report with embedded charts
├── requirements.txt                              # Required Python libraries
├── README.md                                     # Markdown documentation
├── data/
│   ├── raw/
│   │   └── online_retail.csv                     # Raw transactional event stream
│   └── processed/
│       ├── fact_transactions.csv                 # Cleaned fact table
│       ├── fact_reversals_audit.csv              # Audit ledger of cancellations & returns
│       └── dim_customer_rfm.csv                  # Dimensional customer table with RFM scores
├── src/
│   ├── __init__.py
│   ├── clean_refinery_pipeline.py                # Data cleaning & Tukey IQR outlier filtering
│   ├── rfm_feature_engineering.py                # Vectorized RFM feature aggregation
│   └── segmentation_engine.py                    # Log1p scaling, K-Means & quantile scoring
├── dashboard/
│   ├── index.html                               # Interactive Executive Dashboard (HTML5/Tailwind/Chart.js)
│   └── summary_data.json                        # Aggregated KPI and segment metrics
├── tests/
│   └── test_pipeline.py                          # Automated unit test suite
├── main.py                                       # Pipeline orchestration script
└── generate_dataset.py                           # Dataset generation utility
```

---

## 7. Setup & Execution Instructions

### Step 1: Clone Repository & Create Virtual Environment
```bash
git clone https://github.com/your-username/ecommerce-rfm-analytics.git
cd ecommerce-rfm-analytics

# Create isolated Python virtual environment
python3 -m venv venv
source venv/bin/activate
```

### Step 2: Install Project Dependencies
```bash
pip install -r requirements.txt
```

### Step 3: Execute End-to-End Pipeline
```bash
# Execute master pipeline (cleans data, computes RFM, trains K-Means, exports data marts)
python main.py
```

### Step 4: Run Unit Tests
```bash
python -m unittest tests/test_pipeline.py
```

### Step 5: Launch Executive BI Dashboard
Open the compiled standalone dashboard directly in any modern browser:
```bash
open dashboard/index.html
```
Or view the complete analysis in the Jupyter Notebook:
```bash
jupyter notebook Deep_Koshiya_Ecommerce_RFM_Segmentation.ipynb
```

---

## 8. Author Contact
- **Author:** Deep Koshiya
- **Project:** E-Commerce Customer Lifecycle & RFM Segmentation Dashboard
- **Submission Date:** September 2026
