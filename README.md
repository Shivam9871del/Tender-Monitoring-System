# Tender Monitoring System

## Project Overview

This project automates the monitoring of tender activities and provides daily management reports for tracking tender progress, SLA compliance, officer workload, and data quality issues.

The system reads tender data from an Excel workbook, processes the information using Python and Pandas, and generates multiple management reports in Excel format.

---

## Features

### Daily Dashboard

Provides overall tender statistics:

- Total tenders
- Under process
- Completed
- Cancelled
- GM attention required
- Delayed tenders
- On timeline tenders

---

### Live Tender Monitoring

Tracks:

- Current tender stage
- Pending officer
- Planned completion date
- Delay days
- SLA status
- Tender remarks

---

### GM Attention Report

Highlights critical tenders based on:

- Delay greater than 30 days
- High value delayed tenders
- Approval bottlenecks

---

### Officer Workload

Shows officer-wise:

- Running tenders
- Delayed tenders
- Before timeline tenders
- On timeline tenders

---

### Exception Report

Automatically identifies data quality issues such as:

- Missing previous activity dates
- Incorrect sequence of tender activities
- Out-of-order milestone completion

---

## Input

Excel workbook containing:

- Tender master
- Officer master
- Activity configuration

---

## Output

Excel report containing:

- Dashboard
- GM Attention
- Ongoing Tenders
- Completed Tenders
- Cancelled Tenders
- Officer Workload
- Exception Report

---

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Excel
- DateTime

---

## Business Value

The system helps management:

- Monitor tender progress.
- Identify delayed activities.
- Track officer workload.
- Detect data entry issues.
- Improve SLA compliance.

---

## Future Enhancements

- Exception severity classification.
- Officer-wise exception dashboard.
- Automatic email alerts.
- Interactive Power BI dashboard.
