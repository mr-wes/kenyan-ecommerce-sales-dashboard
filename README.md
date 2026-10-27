# Kenyan E-commerce Sales Intelligence Dashboard

An interactive sales analytics dashboard built with Python and Streamlit to explore e-commerce performance, identify business trends, and support data-driven decision-making.

## Project Overview

This project transforms e-commerce transaction data into an interactive business intelligence dashboard. It enables users to explore sales performance, analyze products, monitor returns and financial metrics, assess data quality, and investigate individual transactions.

## Dashboard Features

The dashboard contains six analytical pages:

### 1. Sales Overview

* Monitor key sales performance indicators.
* Explore sales trends over time.
* Review overall business performance.

### 2. Executive Insights

* Summarize important business performance indicators.
* Highlight patterns and trends relevant to decision-makers.
* Support data-driven business discussions.

### 3. Product Analysis

* Analyze product performance.
* Compare products using sales-related metrics.
* Identify products that contribute to business performance.

### 4. Returns & Finance

* Examine product returns.
* Explore relevant financial metrics.
* Investigate the relationship between sales and returns.

### 5. Data Quality

* Inspect the dataset for potential data quality issues.
* Review missing or inconsistent values.
* Support more reliable analysis.

### 6. Transactions

* Explore transaction-level records.
* Filter and investigate individual transactions.
* Examine the underlying data behind business metrics.

## Technology Stack

* **Python** — application logic and data processing
* **Streamlit** — interactive dashboard interface
* **Pandas** — data manipulation and analysis
* **Plotly** — interactive data visualizations
* **OpenPyXL** — Excel file support

## Project Structure

```text
04_sales_dashboard/
├── app.py
├── requirements.txt
├── .gitignore
├── README.md
└── kenyan_ecommerce_sales.xlsx
```

The Excel dataset should only be included if its redistribution is permitted. If it is excluded, the application must have another way to access the required data.

## Getting Started

### Prerequisites

* Python 3.10 or a compatible version supported by the installed dependencies
* pip

### Installation

1. Clone this repository:

   ```bash
   git clone YOUR_GITHUB_REPOSITORY_URL
   ```

2. Navigate to the project directory:

   ```bash
   cd 04_sales_dashboard
   ```

3. Create a virtual environment:

   ```bash
   python -m venv .venv
   ```

4. Activate the virtual environment.

   **Windows PowerShell:**

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

   **macOS/Linux:**

   ```bash
   source .venv/bin/activate
   ```

5. Install the dependencies:

   ```bash
   pip install -r requirements.txt
   ```

6. Start the dashboard:

   ```bash
   streamlit run app.py
   ```

7. Open the local address displayed in your terminal.

## Dataset

The project uses e-commerce sales transaction data stored in an Excel workbook.

The dashboard is designed to support exploratory analysis of sales performance, product trends, returns, financial indicators, data quality, and transaction records.

Dataset attribution and redistribution permissions should be confirmed before publishing the data publicly.

## Future Improvements

Potential future enhancements include:

* Cloud deployment
* Automated data refresh
* Expanded sales forecasting
* Additional customer and product segmentation
* Automated data quality reporting

## Author

[mr-wes](https://github.com/mr-wes)
