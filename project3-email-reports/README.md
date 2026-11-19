# Automated Sales Report & Email Sender (Python)

Reads raw sales data from Excel, calculates KPIs (total revenue, order count, top 5 best-selling products), generates a high-resolution bar chart, and emails an HTML summary report with attachments automatically.

## Features
- **Automated KPI Analysis:** Aggregates revenue and order volume using `pandas`.
- **Data Visualization:** Generates a clean, labeled bar chart using `matplotlib`.
- **HTML Email Formatting:** Sends a styled summary table via SMTP (`smtplib`) with the chart and Excel file attached.
- **Credential Security:** Uses environment variables (`.env`) to keep email credentials safe.

## How to run
1. Install dependencies:
   ```bash
   pip install pandas openpyxl matplotlib python-dotenv