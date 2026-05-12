import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd

xl = pd.ExcelFile("backend/data/App_Url Data base.xlsx")
df = xl.parse('Sheet1', header=1)  # row 1 is the real header: ID, URL/App Name, Language or Line Item
print("Columns:", list(df.columns))
print(f"Total rows: {len(df)}")
print("\nAll unique values in Language col:")
print(df.iloc[:, 2].dropna().unique()[:30])
print("\nFirst 20 rows:")
print(df.head(20).to_string())
