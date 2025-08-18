import pandas as pd

# Load your original Excel file (replace with your filename)
file_path = "StkGrpSum.xlsx"   # keep this file in the same folder as script
df = pd.read_excel(file_path, sheet_name="Stock Group Summary", header=None)
df.columns = ["Item"]

# Define broad customer-friendly categories
categories = {
    "Wires & Cables": ["wire", "cable", "flexible", "cctv", "pair"],
    "Switchgear & Protection": ["mcb", "rccb", "isolator", "contactor", "relay", "db", "distribution",
                                "mccb", "acb", "changeover", "fuse", "connector", "neutral link"],
    "Conduits & Fittings": ["conduit", "bend", "coupler", "junction", "box", "pipe"],
    "Fans & Appliances": ["fan", "exhaust"],
    "Lighting": ["led", "bulb", "tube", "light", "panel", "flood"],
    "Hardware & Accessories": ["screw", "tape", "clamp", "saddle", "pop"]
}

# Function to categorize each item
def categorize_item(item):
    item_lower = str(item).lower()
    for cat, keywords in categories.items():
        if any(keyword in item_lower for keyword in keywords):
            return cat
    return "Others"

# Apply categorization
df["Category"] = df["Item"].apply(categorize_item)

# Sort by category for catalog style
df_sorted = df.sort_values(by=["Category", "Item"]).reset_index(drop=True)

# Save new cleaned file
output_path = "Customer_Friendly_Catalog.xlsx"
df_sorted.to_excel(output_path, index=False)

print(f"✅ Catalog saved as {output_path}")