import pandas as pd
from docx import Document
from math import ceil
from datetime import datetime

# Load Excel
df = pd.read_excel("data/master_projects.xlsx")

# Take first row
data = df.iloc[0]

# Extract values
start_date = pd.to_datetime(data["Start_Date"])
end_date = pd.to_datetime(data["End_Date"])
total_value = float(data["Total_Value"])

# Calculate duration
duration_days = (end_date - start_date).days
months = ceil(duration_days / 30)

# Monthly value
monthly_value = total_value / months

# Load Word
doc = Document("templates/business_case_template.docx")

# Replace normal placeholders
mapping = {
    "{{Contract_Number}}": str(data["Contract_Number"]),
    "{{Work_Order_Number}}": str(data["Work_Order_Number"]),
    "{{Project_Name}}": str(data["Project_Name"]),
    "{{Date}}": str(data["Date"]),
    "{{Total_Value}}": f"AED {total_value:,.2f}"
}

for p in doc.paragraphs:
    for key, val in mapping.items():
        if key in p.text:
            p.text = p.text.replace(key, val)

# 🔥 Create Cashflow Table
for p in doc.paragraphs:
    if "{{CASHFLOW_TABLE}}" in p.text:

        p.text = p.text.replace("{{CASHFLOW_TABLE}}", "")

        table = doc.add_table(rows=2, cols=months+1)

        # Header row
        table.cell(0, 0).text = "Roles"
        for i in range(months):
            table.cell(0, i+1).text = f"Month {i+1}"

        # Data row
        table.cell(1, 0).text = "Contractor Fees"
        for i in range(months):
            table.cell(1, i+1).text = f"{monthly_value:,.2f}"

# Save
doc.save("output/output.docx")

print("✅ Business Case Generated with Dynamic Cashflow!")
