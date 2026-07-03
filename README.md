# Tender Monitoring System

Python/Pandas based Tender Monitoring and SLA Tracking System for automated MIS reporting.

## Project Overview

This project automates tender lifecycle monitoring from Excel data. It processes 500 tender records and generates management-level reports for tender status, SLA delay, pending responsibility, GM attention cases, officer workload, completed tenders, cancelled tenders, and exception records.

## Dataset

The input file contains 500 tender records across:

- Services
- Materials
- Works

Tender status distribution:

- Completed tenders
- Under process tenders
- Cancelled tenders

The dataset includes 20 tender officers and finance/GM approval routing.

## Key Features

- Reads tender data from Excel
- Calculates current tender stage
- Identifies pending officer or approver
- Calculates planned date and delay days
- Marks SLA status as Delayed, On Timeline, or Before Timeline
- Generates GM attention report
- Generates officer-wise running tender summary
- Generates exception report for data-quality issues
- Exports final MIS report to Excel

## Reports Generated

The output Excel report contains:

- dashboard
- gm_attention
- ongoing_tenders
- completed_tenders
- cancelled_tenders
- officer_summary
- exception_report

## Tech Stack

- Python
- Pandas
- OpenPyXL
- Excel

## How to Run

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the project:

```bash
python tender_monitor.py
```

The output report will be generated at:

```text
reports/daily_monitoring_report_500.xlsx
```
