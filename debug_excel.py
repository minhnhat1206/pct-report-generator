
import pandas as pd
import os

files = ['StudentList10.xlsx', 'StudentList11.xlsx']
for f in files:
    if os.path.exists(f):
        df = pd.read_excel(f)
        print(f"File: {f}")
        for col in df.columns:
            print(f"  - [{col}]")
        print("-" * 20)
    else:
        print(f"File {f} not found")
