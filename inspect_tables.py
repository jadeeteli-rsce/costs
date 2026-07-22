import pandas as pd
import glob

pd.set_option("display.max_columns", None)
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", None)

files = glob.glob("data/*.xlsx") + glob.glob("data/**/*.xlsx", recursive=True)
files = sorted(set(files))

for path in files:
    print(f"\n{'='*80}\n{path}\n{'='*80}")
    xls = pd.ExcelFile(path)
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, header=None)
        print(f"\n--- Sheet: {sheet} ({df.shape[0]} rows x {df.shape[1]} cols) ---")
        with pd.option_context("display.max_rows", 200):
            print(df.to_string())
