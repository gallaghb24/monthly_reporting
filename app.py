import pandas as pd

# Load the Excel file
file_path = "Artwork_Versions_Client-13.xlsx"  # Update this path if needed
xls = pd.ExcelFile(file_path)

# Load the sheet
df = pd.read_excel(xls, sheet_name="general_report")

# Step 1: Sort by column 8 (Client Versions) in descending order
df = df.sort_values(by="Client Versions", ascending=False)

# Step 2: Dedupe based on column 6 (POS Code)
df = df.drop_duplicates(subset=["POS Code"], keep="first")

# Step 3: Remove any rows where Client Versions = 0
df = df[df["Client Versions"] != 0]

# Step 4: Create column 10 ("Amends") as Client Versions - 1
df["Amends"] = df["Client Versions"] - 1

# Step 5: Create column 11 ("Right First Time")
df["Right First Time"] = df["Client Versions"].apply(lambda x: 1 if x == 1 else 0)

# Step 6: If "ROI" in Project Description, set Category to "ROI"
df.loc[df["Project Description"].str.contains("ROI", na=False), "Category"] = "ROI"

# Step 7: Change Members/Starbuys category to "Main Event"
df.loc[df["Category"].isin(["Members", "Starbuys"]), "Category"] = "Main Event"

# Step 8: Change Loyalty / CRM and Mobile to "Other"
df.loc[df["Category"].isin(["Loyalty / CRM", "Mobile"]), "Category"] = "Other"

# Step 9: Calculate stats
num_new_artworks = len(df)
total_amends = df["Amends"].sum()
num_right_first_time = df["Right First Time"].sum()
right_first_time_percentage = (num_right_first_time / num_new_artworks) * 100
average_amend_rate = round(df["Amends"].mean(), 2)

# Print results
print("New artworks created:", num_new_artworks)
print("Total rounds of amends:", total_amends)
print("Right first time:", num_right_first_time)
print("Right first time %:", f"{right_first_time_percentage:.2f}%")
print("Average amend rate:", average_amend_rate)
