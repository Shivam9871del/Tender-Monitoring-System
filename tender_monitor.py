# ==============================================================
# Project: Tender Lifecycle Monitoring Automation - Version 2
# Purpose: Read raw tender Excel data, calculate current status,
#          pending owner, delay, SLA status, GM attention cases,
#          officer workload, and exception report.
# Input  : tender_input_master_final.xlsx
# Output : reports/daily_monitoring_report_500.xlsx
# ============================================================== 

import os
from datetime import datetime

import pandas as pd


# --------------------------------------------------------------
# 1. File paths
# --------------------------------------------------------------
# INPUT_FILE should be the corrected 500-tender raw Excel file.
# Keep the Excel file in the same folder where this Python file is kept.
INPUT_FILE = "tender_input_master_final.xlsx"

# Report will be generated inside the reports folder.
OUTPUT_FILE = "reports/daily_monitoring_report_500.xlsx"

# Create reports folder automatically if it does not already exist.
os.makedirs("reports", exist_ok=True)


# --------------------------------------------------------------
# 2. Report date
# --------------------------------------------------------------
# For project testing/demo, we are using fixed date 30-06-2026
# because the raw data is created from 01-01-2026 to 30-06-2026.
# Later, when connecting with Airflow, replace this line with:
# today = pd.Timestamp.today().normalize()
today = pd.to_datetime("2026-06-30")


# --------------------------------------------------------------
# 3. Read all required sheets from Excel
# --------------------------------------------------------------
# tender_master = main transaction data of all tenders
# officers      = officer master data, used to map Officer ID to Officer Name
# approvers     = approver master data, used for finance/GM approval owners
# config        = planned activity days, kept for project reference

df = pd.read_excel(INPUT_FILE, sheet_name="tender_master")
officers = pd.read_excel(INPUT_FILE, sheet_name="officers")
approvers = pd.read_excel(INPUT_FILE, sheet_name="approvers")
config = pd.read_excel(INPUT_FILE, sheet_name="config")


# --------------------------------------------------------------
# 4. Create lookup dictionaries from master sheets
# --------------------------------------------------------------
# Example: O-01 -> Aarav Mehta
# Example: A-01 -> Vivek Arora

officer_name_dict = dict(zip(officers["officer_id"], officers["officer_name"]))
approver_name_dict = dict(zip(approvers["approver_id"], approvers["approver_name"]))


# --------------------------------------------------------------
# 5. Define tender lifecycle stage columns
# --------------------------------------------------------------
# These are the date columns used to identify where a tender is currently stuck.
# The first blank stage date becomes the Current Stage for Under Process tenders.

stage_columns = [
    "NIT Date",
    "Tender Floating Date",
    "Technical Bid Opening Date",
    "TBA Date",
    "CBA Date",
    "PBO Date",
    "Price Comparison Date",
    "Workability/Reasonability Date",
    "Negotiation Date",
    "Award Approval Date",
    "Order Issued Date",
]

# File received date + all stage dates should be converted into datetime format.
date_columns = ["File/Tender Received Date"] + stage_columns


# --------------------------------------------------------------
# 6. Convert all date columns safely
# --------------------------------------------------------------
# errors="coerce" means invalid/blank dates become NaT.
# NaT is pandas' datetime version of NULL.

for col in date_columns:
    df[col] = pd.to_datetime(df[col], errors="coerce")


# --------------------------------------------------------------
# 7. Planned cumulative days from file received date
# --------------------------------------------------------------
# These values are used to calculate planned date and delay.
# Example: If current stage is TBA Date, expected TBA completion =
# File Received Date + 30 days.

cumulative_days = {
    "NIT Date": 3,
    "Tender Floating Date": 6,
    "Technical Bid Opening Date": 23,
    "TBA Date": 30,
    "CBA Date": 35,
    "PBO Date": 38,
    "Price Comparison Date": 39,
    "Workability/Reasonability Date": 42,
    "Negotiation Date": 47,
    "Award Approval Date": 50,
    "Order Issued Date": 51,
}


# --------------------------------------------------------------
# 8. Category-wise GM mapping
# --------------------------------------------------------------
# SER = Services  -> GM Contracts
# MAT = Materials -> GM Materials
# WRK = Works     -> GM Works

category_gm_approver_id = {
    "SER": "A-04",
    "MAT": "A-05",
    "WRK": "A-06",
}


def get_category_gm(row):
    """
    Return the GM approver name based on tender category.
    This is used mainly during Award Approval stage.
    """
    approver_id = category_gm_approver_id.get(row["Category"])
    return approver_name_dict.get(approver_id, "GM Not Mapped")


def get_finance_officer(row):
    """
    Return finance officer based on tender value.
    Rule:
    - Up to 5 crore  = Assistant Manager Finance
    - Above 5 crore  = Senior Manager Finance

    Note: Tender value is in Lakhs, so 5 crore = 500 lakhs.
    """
    if row["Estimate Value (Lakhs)"] > 500:
        return approver_name_dict.get("A-02", "Senior Manager Finance")

    return approver_name_dict.get("A-01", "Assistant Manager Finance")


def get_gm_finance():
    """
    Return GM Finance approver name.
    We are keeping GM Finance as one single role.
    """
    return approver_name_dict.get("A-03", "GM Finance")


# --------------------------------------------------------------
# 9. Current stage calculation
# --------------------------------------------------------------
# Logic:
# - Cancelled tender -> Cancelled
# - Order Issued Date present -> Completed
# - Otherwise, first blank stage date is the current pending stage


def get_current_stage(row):
    """Identify the current lifecycle stage of each tender."""

    if row["Tender Status"] == "Cancelled":
        return "Cancelled"

    if pd.notna(row["Order Issued Date"]):
        return "Completed"

    for stage in stage_columns:
        if pd.isna(row[stage]):
            return stage

    return "Completed"


# --------------------------------------------------------------
# 10. Pending-with calculation
# --------------------------------------------------------------
# This function decides who is responsible for the current pending stage.
# Most operational stages are pending with the assigned tender officer.
# Finance stages are pending with finance officer / GM Finance.
# Award approval is pending with GM Finance + category GM.


def get_pending_with(row):
    """Return the person/authority with whom the tender is currently pending."""

    stage = row["Current Stage"]

    if stage == "Completed":
        return "Completed"

    if stage == "Cancelled":
        return "Cancelled"

    officer_stages = [
        "NIT Date",
        "Tender Floating Date",
        "Technical Bid Opening Date",
        "TBA Date",
        "CBA Date",
        "PBO Date",
        "Price Comparison Date",
        "Negotiation Date",
        "Order Issued Date",
    ]

    if stage in officer_stages:
        return officer_name_dict.get(row["Officer ID"], row["Officer ID"])

    if stage == "Workability/Reasonability Date":
        return get_finance_officer(row)

    if stage == "Award Approval Date":
        return get_gm_finance() + " & " + get_category_gm(row)

    return officer_name_dict.get(row["Officer ID"], row["Officer ID"])


# --------------------------------------------------------------
# 11. Planned date calculation
# --------------------------------------------------------------
# Planned Date = File/Tender Received Date + planned cumulative days
# Cancelled tenders do not need planned date in monitoring report.


def get_planned_date(row):
    """Calculate planned date for the current pending stage."""

    stage = row["Current Stage"]

    if stage == "Completed":
        return row["Order Issued Date"]

    if stage == "Cancelled":
        return ""

    file_received_date = pd.to_datetime(row["File/Tender Received Date"])
    days_to_add = cumulative_days.get(stage, 0)
    planned_date = file_received_date + pd.Timedelta(days=days_to_add)

    return planned_date.strftime("%d-%m-%Y")


# --------------------------------------------------------------
# 12. Delay calculation
# --------------------------------------------------------------
# For completed tenders:
# Delay = Actual Order Issued Date - Planned Order Issued Date
#
# For under-process tenders:
# Delay = Report Date - Planned Date of current stage
#
# Negative delay means tender is before timeline.
# Zero delay means tender is on timeline.
# Positive delay means tender is delayed.


def get_delay_days(row):
    """Calculate delay days for completed and under-process tenders."""

    if row["Tender Status"] == "Cancelled":
        return ""

    if row["Current Stage"] == "Completed":
        planned_order_date = (
            pd.to_datetime(row["File/Tender Received Date"])
            + pd.Timedelta(days=51)
        )

        actual_order_date = pd.to_datetime(row["Order Issued Date"])
        return (actual_order_date - planned_order_date).days

    planned_date = pd.to_datetime(
        row["Planned Date"],
        format="%d-%m-%Y",
        errors="coerce",
    )

    return (today - planned_date).days


# --------------------------------------------------------------
# 13. SLA status calculation
# --------------------------------------------------------------
# SLA Status is derived from Delay Days.


def get_sla_status(row):
    """Classify each tender as Before Timeline, On Timeline, Delayed, or Cancelled."""

    if row["Tender Status"] == "Cancelled":
        return "Cancelled"

    if row["Delay Days"] < 0:
        return "Before Timeline"

    if row["Delay Days"] == 0:
        return "On Timeline"

    return "Delayed"


# --------------------------------------------------------------
# 14. Exception report logic
# --------------------------------------------------------------
# This does not mean exception data is stored in raw Excel.
# Raw Excel remains clean transaction data.
# This function only checks the raw data and generates exception output.
# If data is clean, exception_report may be blank.


def build_exception_report(df):
    """Find business-rule violations and return them as an exception report."""

    exception_rows = []

    def add_exception(mask, issue):
        """Small helper function to add matching rows into exception list."""

        temp = df.loc[
            mask,
            [
                "Tender ID",
                "Category",
                "Officer ID",
                "Estimate Value (Lakhs)",
                "Tender Status",
                "Remarks",
            ],
        ].copy()

        if len(temp) > 0:
            temp["Exception Issue"] = issue
            exception_rows.append(temp)

    # Rule 1: Commercial Bid Approval cannot happen before Technical Bid Approval.
    add_exception(
        df["CBA Date"].notna() & df["TBA Date"].isna(),
        "CBA Date present but TBA Date blank",
    )

    # Rule 2: If later stages are present after PBO, Price Comparison should not be blank.
    add_exception(
        df["PBO Date"].notna()
        & df["Price Comparison Date"].isna()
        & df[
            [
                "Workability/Reasonability Date",
                "Negotiation Date",
                "Award Approval Date",
                "Order Issued Date",
            ]
        ].notna().any(axis=1),
        "PBO/later stage present but Price Comparison Date blank",
    )

    # Rule 3: Order cannot be issued before award approval.
    add_exception(
        df["Order Issued Date"].notna() & df["Award Approval Date"].isna(),
        "Order Issued Date present but Award Approval Date blank",
    )

    # Rule 4: Remarks should not say tender floated if Tender Floating Date is blank.
    add_exception(
        df["Remarks"].astype(str).str.contains("tender floated", case=False, na=False)
        & df["Tender Floating Date"].isna(),
        "Remarks say tender floated but Tender Floating Date is blank",
    )

    # Rule 5: Award Approval Date should not be before PBO Date.
    add_exception(
        df["Award Approval Date"].notna()
        & df["PBO Date"].notna()
        & (df["Award Approval Date"] < df["PBO Date"]),
        "Award Approval Date is before PBO Date",
    )

    # Rule 6: Technical Bid Opening should not exist if Tender Floating Date is blank.
    add_exception(
        df["Technical Bid Opening Date"].notna()
        & df["Tender Floating Date"].isna(),
        "Technical Bid Opening Date present but Tender Floating Date blank",
    )

    if len(exception_rows) == 0:
        return pd.DataFrame(
            columns=[
                "Tender ID",
                "Category",
                "Officer ID",
                "Officer Name",
                "Estimate Value (Lakhs)",
                "Tender Status",
                "Exception Issue",
                "Remarks",
            ]
        )

    exception_report = pd.concat(exception_rows, ignore_index=True)

    # Add officer name for easier business reading.
    exception_report["Officer Name"] = exception_report["Officer ID"].map(officer_name_dict)

    exception_report = exception_report[
        [
            "Tender ID",
            "Category",
            "Officer ID",
            "Officer Name",
            "Estimate Value (Lakhs)",
            "Tender Status",
            "Exception Issue",
            "Remarks",
        ]
    ]

    return exception_report.drop_duplicates()


# --------------------------------------------------------------
# 15. Apply all business functions on main dataframe
# --------------------------------------------------------------
# axis=1 means function is applied row by row.

df["Current Stage"] = df.apply(get_current_stage, axis=1)
df["Pending With"] = df.apply(get_pending_with, axis=1)
df["Planned Date"] = df.apply(get_planned_date, axis=1)
df["Delay Days"] = df.apply(get_delay_days, axis=1)
df["SLA Status"] = df.apply(get_sla_status, axis=1)


# --------------------------------------------------------------
# 16. Report generation date and time
# --------------------------------------------------------------
# This is actual system time when script runs.

report_datetime = datetime.now()
report_date = report_datetime.strftime("%d-%m-%Y")
report_time = report_datetime.strftime("%H:%M:%S")


# --------------------------------------------------------------
# 17. Main live monitoring report
# --------------------------------------------------------------
# This is the key business output used by users/managers.

live_monitoring = df[
    [
        "Tender ID",
        "Category",
        "Officer ID",
        "Estimate Value (Lakhs)",
        "File/Tender Received Date",
        "Current Stage",
        "Pending With",
        "Planned Date",
        "Delay Days",
        "SLA Status",
        "Tender Status",
        "Remarks",
    ]
]


# --------------------------------------------------------------
# 18. Split report into status-wise datasets
# --------------------------------------------------------------
# These sheets make the final Excel easier to read.

ongoing_tenders = live_monitoring[live_monitoring["Tender Status"] == "Under Process"]
completed_tenders = live_monitoring[live_monitoring["Tender Status"] == "Completed"]
cancelled_tenders = live_monitoring[live_monitoring["Tender Status"] == "Cancelled"]


# --------------------------------------------------------------
# 19. GM attention report
# --------------------------------------------------------------
# Business logic:
# - Any ongoing tender delayed by more than 30 days
# - Any high-value tender above 10 crore delayed by even 1 day
# - Any tender pending at Award Approval stage
#
# Note: 10 crore = 1000 lakhs.

gm_attention = ongoing_tenders[
    (ongoing_tenders["Delay Days"] > 30)
    |
    (
        (ongoing_tenders["Estimate Value (Lakhs)"] > 1000)
        & (ongoing_tenders["Delay Days"] > 0)
    )
    |
    (ongoing_tenders["Current Stage"].isin(["Award Approval Date"]))
]

gm_attention = gm_attention.sort_values(by="Delay Days", ascending=False)


# --------------------------------------------------------------
# 20. Officer-wise running tender summary
# --------------------------------------------------------------
# This is dynamic, so it works for 10 officers, 20 officers, or more.
# No hardcoding of O-01 to O-20 is needed.

officer_summary = (
    ongoing_tenders
    .groupby("Officer ID")
    .agg(
        Total_Running_Tenders=("Tender ID", "count"),
        Delayed_Tenders=("SLA Status", lambda x: (x == "Delayed").sum()),
        Before_Timeline=("SLA Status", lambda x: (x == "Before Timeline").sum()),
        On_Timeline=("SLA Status", lambda x: (x == "On Timeline").sum()),
    )
    .reset_index()
)

# Add officer name beside officer ID.
officer_summary["Officer Name"] = officer_summary["Officer ID"].map(officer_name_dict)

officer_summary = officer_summary[
    [
        "Officer ID",
        "Officer Name",
        "Total_Running_Tenders",
        "Delayed_Tenders",
        "Before_Timeline",
        "On_Timeline",
    ]
]


# --------------------------------------------------------------
# 21. Exception report output
# --------------------------------------------------------------
# This report is generated by automation. It is not stored in raw data.

exception_report = build_exception_report(df)


# --------------------------------------------------------------
# 22. Dashboard summary
# --------------------------------------------------------------
# This gives a quick management-level count summary.

dashboard_data = [
    ["Report Generated On", report_date],
    ["Report Time", report_time],
    ["Total Tenders", len(live_monitoring)],
    ["Under Process", len(ongoing_tenders)],
    ["Completed", len(completed_tenders)],
    ["Cancelled", len(cancelled_tenders)],
    ["GM Attention Required", len(gm_attention)],
    ["Delayed Ongoing Tenders", len(ongoing_tenders[ongoing_tenders["Delay Days"] > 0])],
    ["Ongoing Before Timeline", len(ongoing_tenders[ongoing_tenders["Delay Days"] < 0])],
    ["Ongoing On Timeline", len(ongoing_tenders[ongoing_tenders["Delay Days"] == 0])],
    ["Exception Records", len(exception_report)],
]

dashboard = pd.DataFrame(dashboard_data, columns=["Metric", "Value"])


# --------------------------------------------------------------
# 23. Write all reports into one Excel workbook
# --------------------------------------------------------------
# Each dataframe becomes a separate sheet.

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    dashboard.to_excel(writer, sheet_name="dashboard", index=False)
    gm_attention.to_excel(writer, sheet_name="gm_attention", index=False)
    ongoing_tenders.to_excel(writer, sheet_name="ongoing_tenders", index=False)
    completed_tenders.to_excel(writer, sheet_name="completed_tenders", index=False)
    cancelled_tenders.to_excel(writer, sheet_name="cancelled_tenders", index=False)
    officer_summary.to_excel(writer, sheet_name="officer_summary", index=False)
    exception_report.to_excel(writer, sheet_name="exception_report", index=False)


# --------------------------------------------------------------
# 24. Terminal output for quick validation
# --------------------------------------------------------------
# This helps us verify if the run was successful without opening Excel.

print()
print("DAILY TENDER MONITORING REPORT")
print()
print("Report Generated On :", report_date)
print("Report Time         :", report_time)
print()
print("Total Tenders       :", len(live_monitoring))
print("Under Process       :", len(ongoing_tenders))
print("Completed           :", len(completed_tenders))
print("Cancelled           :", len(cancelled_tenders))
print("GM Attention        :", len(gm_attention))
print("Exception Records   :", len(exception_report))
print()
print("Daily monitoring report generated successfully.")
