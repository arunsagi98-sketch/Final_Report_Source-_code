import io
import sys

from reference_db import load_app_master_df

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

df = load_app_master_df()
print("Columns:", list(df.columns))
print(f"Total rows: {len(df)}")
print("\nAll unique values in Language col:")
if "Language or Line Item" in df.columns and df.shape[1] >= 3:
    print(df.iloc[:, 2].dropna().unique()[:30])
else:
    print([])
print("\nFirst 20 rows:")
print(df.head(20).to_string())
