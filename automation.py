import pandas as pd
import os
input_file = "data/raw_data.xlsx"
output_file = "data/final_report.xlsx"
print("Starting automation...")
print(f"Reading input file: {input_file}")


# File paths
input_file = "data/raw_data.xlsx"
output_file = "data/final_report.xlsx"

# Check if input exists
if not os.path.exists(input_file):
    print("Error: Input file not found.")
    exit()

# Read data
df = pd.read_excel(input_file)

# Transform data
df["Productivity"] = df["Tasks Completed"] / df["Hours Worked"]
df["Error Rate"] = df["Errors"] / df["Tasks Completed"]
df["Low Performer"] = df["Productivity"] < 5

# Summary metrics
summary = pd.DataFrame({
    "Metric": [
        "Total Tasks Completed",
        "Average Productivity",
        "Average Error Rate",
        "Low Performers Count"
    ],
    "Value": [
        df["Tasks Completed"].sum(),
        df["Productivity"].mean(),
        df["Error Rate"].mean(),
        df["Low Performer"].sum()
    ]
})

# Export report
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Detailed Report", index=False)
    summary.to_excel(writer, sheet_name="Summary", index=False)

print(f"Report generated at: {output_file}")
print("Automation completed successfully.")
