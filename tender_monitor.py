import pandas as pd
from datetime import datetime

INPUT_FILE = "tender_input_master_final.xlsx"
OUTPUT_FILE = "reports/daily_monitoring_report.xlsx"

today = pd.to_datetime("2026-05-15")

df = pd.read_excel(INPUT_FILE, sheet_name="tender_master")
officers = pd.read_excel(INPUT_FILE, sheet_name="officers")
config = pd.read_excel(INPUT_FILE, sheet_name="config")

officer_name_dict = dict(zip(officers["officer_id"], officers["officer_name"]))

approver_dict = {
    "A2": "Sunil Mehta",
    "A3": "Rakesh Srivastava"
}

planned_days_dict = dict(zip(config["activity_name"], config["planned_days"]))

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
    "Order Issued Date"
]

stage_to_activity = {
    "NIT Date": "NIT Preparation",
    "Tender Floating Date": "Tender Floating",
    "Technical Bid Opening Date": "Technical Bid Opening",
    "TBA Date": "TBA",
    "CBA Date": "CBA",
    "PBO Date": "PBO",
    "Price Comparison Date": "Price Comparison",
    "Workability/Reasonability Date": "Workability / Reasonability",
    "Negotiation Date": "Negotiation",
    "Award Approval Date": "Award Approval",
    "Order Issued Date": "Order Issued"
}

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
    "Order Issued Date": 51
}

def get_current_stage(row):

    if pd.notna(row["Order Issued Date"]):
        return "Completed"

    for stage in stage_columns:
        if pd.isna(row[stage]):
            return stage

    return "Completed"

def get_pending_with(row):
    stage = row["Current Stage"]

    if stage == "Completed":
        return "Completed"

    if row["Tender Status"] == "Cancelled":
        return "Cancelled"

    if stage in [
        "NIT Date",
        "Tender Floating Date",
        "Technical Bid Opening Date",
        "TBA Date",
        "CBA Date",
        "PBO Date",
        "Price Comparison Date",
        "Negotiation Date",
        "Order Issued Date"
    ]:
        return officer_name_dict.get(
            row["Officer ID"],
            row["Officer ID"]
        )

    if stage == "Workability/Reasonability Date":
        return approver_dict["A2"]

    if stage == "Award Approval Date":
        return (
            approver_dict["A2"]
            + " & " +
            approver_dict["A3"]
        )

    return row["Officer ID"]

def get_planned_date(row):
    stage = row["Current Stage"]

    if stage == "Completed":
        return row["Order Issued Date"]

    if row["Tender Status"] == "Cancelled":
        return ""

    file_received_date = pd.to_datetime(row["File/Tender Received Date"])
    days_to_add = cumulative_days[stage]

    planned_date = file_received_date + pd.Timedelta(days=days_to_add)

    return planned_date.strftime("%d-%m-%Y")

def get_delay_days(row):

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
        format="%d-%m-%Y"
    )

    return (today - planned_date).days

def get_sla_status(row):

    if row["Tender Status"] == "Cancelled":
        return "Cancelled"

    if row["Delay Days"] < 0:
        return "Before Timeline"

    if row["Delay Days"] == 0:
        return "On Timeline"

    return "Delayed"

df["Current Stage"] = df.apply(get_current_stage, axis=1)
df["Pending With"] = df.apply(get_pending_with, axis=1)
df["Planned Date"] = df.apply(get_planned_date, axis=1)
df["Delay Days"] = df.apply(get_delay_days, axis=1)
df["SLA Status"] = df.apply(get_sla_status, axis=1)

report_datetime = datetime.now()

report_date = report_datetime.strftime("%d-%m-%Y")
report_time = report_datetime.strftime("%H:%M:%S")

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
        "Remarks"
    ]
]

ongoing_tenders = live_monitoring[
    live_monitoring["Tender Status"] == "Under Process"
]

completed_tenders = live_monitoring[
    live_monitoring["Tender Status"] == "Completed"
]

cancelled_tenders = live_monitoring[
    live_monitoring["Tender Status"] == "Cancelled"
]

gm_attention = ongoing_tenders[
    (ongoing_tenders["Delay Days"] > 30) |
    (
        (ongoing_tenders["Estimate Value (Lakhs)"] > 1000) &
        (ongoing_tenders["Delay Days"] > 0)
    ) |
    (ongoing_tenders["Current Stage"].isin(["Award Approval Date"]))
]

gm_attention = gm_attention.sort_values(
    by="Delay Days",
    ascending=False
)

officer_summary = (
    ongoing_tenders
    .groupby("Officer ID")
    .agg(
        Total_Running_Tenders=("Tender ID", "count"),
        Delayed_Tenders=("SLA Status", lambda x: (x == "Delayed").sum()),
        Before_Timeline=("SLA Status", lambda x: (x == "Before Timeline").sum()),
        On_Timeline=("SLA Status", lambda x: (x == "On Timeline").sum())
    )
    .reset_index()
)

officer_summary["Officer Name"] = officer_summary["Officer ID"].map(officer_name_dict)

officer_summary = officer_summary[
    [
        "Officer ID",
        "Officer Name",
        "Total_Running_Tenders",
        "Delayed_Tenders",
        "Before_Timeline",
        "On_Timeline"
    ]
]

mandatory_stage_columns = [
    "NIT Date",
    "Tender Floating Date",
    "Technical Bid Opening Date",
    "TBA Date",
    "CBA Date",
    "PBO Date",
    "Price Comparison Date",
    "Award Approval Date",
    "Order Issued Date"
]

def get_exception_reason(row):
    issues = []

    # Check missing previous mandatory date
    for i in range(1, len(mandatory_stage_columns)):
        current_col = mandatory_stage_columns[i]
        previous_cols = mandatory_stage_columns[:i]

        if pd.notna(row[current_col]):
            for prev_col in previous_cols:
                if pd.isna(row[prev_col]):
                    issues.append(
                        f"{current_col} entered but {prev_col} is missing"
                    )

    # Check date sequence
    previous_date = None
    previous_col = None

    for col in stage_columns:
        if pd.notna(row[col]):
            current_date = pd.to_datetime(row[col])

            if previous_date is not None and current_date < previous_date:
                issues.append(
                    f"{col} date is earlier than {previous_col}"
                )

            previous_date = current_date
            previous_col = col

    return "; ".join(issues)

df["Exception Reason"] = df.apply(get_exception_reason, axis=1)

exception_report = df[
    df["Exception Reason"] != ""
][
    [
        "Tender ID",
        "Officer ID",
        "Tender Status",
        "Current Stage",
        "Exception Reason",
        "Remarks"
    ]
]

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
    ["Ongoing On Timeline", len(ongoing_tenders[ongoing_tenders["Delay Days"] == 0])]
]

dashboard = pd.DataFrame(
    dashboard_data,
    columns=["Metric", "Value"]
)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    dashboard.to_excel(writer, sheet_name="Dashboard", index=False)
    gm_attention.to_excel(writer, sheet_name="GM Attention", index=False)
    ongoing_tenders.to_excel(writer, sheet_name="ongoing Tenders", index=False)
    completed_tenders.to_excel(writer, sheet_name="Completed Tenders", index=False)
    cancelled_tenders.to_excel(writer, sheet_name="Cancelled Tenders", index=False)
    officer_summary.to_excel(writer, sheet_name="Officer Workload", index=False)
    exception_report.to_excel(writer, sheet_name="Exception Report", index=False)

print()
print("DAILY TENDER MONITORING REPORT")
print()
print("Report Generated On :", report_date)
print("Report Time         :", report_time)
print()
print("Total Tenders :", len(live_monitoring))
print("Under Process :", len(ongoing_tenders))
print("Completed :", len(completed_tenders))
print("Cancelled :", len(cancelled_tenders))
print("GM Attention :", len(gm_attention))
print()
print("Daily monitoring report generated successfully.")
