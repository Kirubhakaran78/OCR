import pandas as pd
import json

# Path to your CSV (from Sonnet4 output folder)
csv_path = "sonnet4_output/pg1_structured_data.csv"

# Read CSV with proper multi-line handling
df = pd.read_csv(
    csv_path,
    sep=",",
    quotechar='"',
    engine='python',      # Handles multi-line quoted fields
    keep_default_na=False
)

# Optional: If you want, parse the structured_data JSON into separate columns
if 'structured_data' in df.columns:
    # Convert JSON string to dict, then expand
    json_cols = df['structured_data'].apply(lambda x: json.loads(x) if x else {})
    json_df = pd.json_normalize(json_cols)
    df = df.drop(columns=['structured_data']).join(json_df)

# Save Excel file
excel_path = "pg1_structured_data.xlsx"
df.to_excel(excel_path, index=False)

print(f"✅ CSV converted to Excel successfully! File saved as: {excel_path}")
