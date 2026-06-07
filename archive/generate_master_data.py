import random
from datetime import datetime, timedelta

import pandas as pd


random.seed(42)

# -----------------------------
# 1. Officers
# -----------------------------
officers = [
    ["O1", "Priya Sharma", "F", "Asst. Mgr.", "MAT"],
    ["O2", "Neha Gupta", "F", "Mgr.", "MAT"],
    ["O3", "Rahul Verma", "M", "Sr. Mgr.", "MAT"],
    ["O4", "Amit Singh", "M", "Asst. Mgr.", "SER-WRK"],
    ["O5", "Rajesh Kumar", "M", "Mgr.", "SER-WRK"],
    ["O6", "Vikas Sharma", "M", "Sr. Mgr.", "SER-WRK"],
    ["O7", "Sandeep Gupta", "M", "Asst. Mgr.", "SER-WRK"],
    ["O8", "Manish Verma", "M", "Mgr.", "SER-WRK"],
    ["O9", "Arjun Yadav", "M", "Sr. Mgr.", "SER-WRK"],
    ["O10", "Pooja Agarwal", "F", "Mgr.", "SER"],
]

officer_df = pd.DataFrame(
    officers,
    columns=["officer_id", "officer_name", "gender", "designation", "category"]
)

# -----------------------------
# 2. Approvers
# -----------------------------
approvers = [
    ["A1", "Deepak Jain", "Mgr.", "FIN", "Finance Vetting"],
    ["A2", "Sunil Mehta", "GM", "FIN", "PBO / Award / Reasonability / Workability"],
    ["A3", "Rakesh Srivastava", "GM", "CONT", "Tender Floating / PBO / Award Approval"],
]

approver_df = pd.DataFrame(
    approvers,
    columns=["approver_id", "approver_name", "designation", "department", "approval_role"]
)

# -----------------------------
# 3. Config - planned days
# -----------------------------
config = [
    ["File Received", 0],
    ["NIT Preparation", 3],
    ["Tender Floating Approval", 2],
    ["Tender Floating", 1],
    ["Bid Submission Period", 15],
    ["Technical Bid Opening", 2],
    ["TBA", 7],
    ["CBA", 5],
    ["PBO Approval", 2],
    ["PBO", 1],
    ["L1 Rates Offered", 1],
    ["Workability / Reasonability", 3],
    ["Negotiation", 5],
    ["Award Approval", 3],
    ["Order Issued", 1],
]

config_df = pd.DataFrame(config, columns=["activity_name", "planned_days"])

print("Base master data created successfully.")

# -----------------------------
# 4. Tender Master Generation
# -----------------------------

officer_tender_count = {
    "O1": 18,
    "O2": 23,
    "O3": 19,
    "O4": 24,
    "O5": 21,
    "O6": 19,
    "O7": 16,
    "O8": 19,
    "O9": 18,
    "O10": 23,
}

officer_category = {
    "O1": "MAT",
    "O2": "MAT",
    "O3": "MAT",
    "O4": "SER-WRK",
    "O5": "SER-WRK",
    "O6": "SER-WRK",
    "O7": "SER-WRK",
    "O8": "SER-WRK",
    "O9": "SER-WRK",
    "O10": "SER",
}

# Exact category distribution
categories = (
    ["MAT"] * 60 +
    ["SER"] * 120 +
    ["WRK"] * 20
)
random.shuffle(categories)

# Tender status distribution
statuses = (
    ["Completed"] * 90 +
    ["Under Process"] * 60 +
    ["Delayed"] * 25 +
    ["Yet To Start"] * 15 +
    ["Cancelled"] * 10
)
random.shuffle(statuses)

# Estimate value distribution
value_bands = (
    ["Low"] * 52 +
    ["Medium"] * 116 +
    ["High"] * 32
)
random.shuffle(value_bands)

def generate_estimate_value(band):
    if band == "Low":
        return round(random.uniform(10, 100), 2)
    elif band == "Medium":
        return round(random.uniform(100, 1000), 2)
    else:
        return round(random.uniform(1000, 10000), 2)

def random_file_received_date():
    start_date = datetime(2026, 1, 1)
    end_date = datetime(2026, 5, 31)
    days_gap = (end_date - start_date).days
    return start_date + timedelta(days=random.randint(0, days_gap))

# Create officer list as per tender count
officer_allocation = []
for officer_id, count in officer_tender_count.items():
    officer_allocation.extend([officer_id] * count)

random.shuffle(officer_allocation)

tender_rows = []

for i in range(1, 201):
    tender_id = f"IOC/BD/26-27/{i:03d}"
    officer_id = officer_allocation[i - 1]
    officer_cat = officer_category[officer_id]

    # Category rule
    if officer_cat == "MAT":
        category = "MAT"
    elif officer_cat == "SER":
        category = "SER"
    else:
        # O4-O9 can handle SER and WRK
        category = categories.pop()

        # If MAT accidentally comes here, convert to SER
        if category == "MAT":
            category = "SER"

    value_band = value_bands[i - 1]
    estimate_value_lakhs = generate_estimate_value(value_band)

    file_received_date = random_file_received_date()
    target_completion_date = file_received_date + timedelta(days=60)

    tender_status = statuses[i - 1]

    bidders_participated = random.randint(2, 20)

    tender_rows.append([
        tender_id,
        category,
        officer_id,
        estimate_value_lakhs,
        value_band,
        file_received_date.strftime("%Y-%m-%d"),
        target_completion_date.strftime("%Y-%m-%d"),
        tender_status,
        bidders_participated,
        "",  # tba_rejected
        "",  # cba_rejected
        "",  # qualified_bidders
        "",  # l1_value_lakhs
        "",  # post_pbo_review_type
        "",  # negotiation_required
        "",  # cancellation_reason
        ""   # remarks
    ])

tender_master_df = pd.DataFrame(
    tender_rows,
    columns=[
        "tender_id",
        "category",
        "officer_id",
        "estimate_value_lakhs",
        "value_band",
        "file_received_date",
        "target_completion_date",
        "tender_status",
        "bidders_participated",
        "tba_rejected",
        "cba_rejected",
        "qualified_bidders",
        "l1_value_lakhs",
        "post_pbo_review_type",
        "negotiation_required",
        "cancellation_reason",
        "remarks"
    ]
)

# -----------------------------
# 5. Export to Excel
# -----------------------------
output_file = "tender_input_master.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    officer_df.to_excel(writer, sheet_name="officers", index=False)
    approver_df.to_excel(writer, sheet_name="approvers", index=False)
    config_df.to_excel(writer, sheet_name="config", index=False)
    tender_master_df.to_excel(writer, sheet_name="tender_master", index=False)

print(f"{output_file} created successfully.")
