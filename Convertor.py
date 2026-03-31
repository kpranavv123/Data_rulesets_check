
# import csv

# filename = "C:/Users/SW526XH/Downloads/Data Quality Check/tabfiles/Part (FG).tab"

# with open(filename, newline="", encoding="utf-8") as f:
#     reader = csv.reader(f, delimiter="\t")
#     next(reader, None)  # skip header
#     row_count = sum(1 for _ in reader)

# print(row_count)


import pandas as pd

# input and output files
tab_file = "C:/Users/SW526XH/Downloads/Data Quality Check/tabfiles/HistoricalDemandActuals.tab"
excel_file = "HistoricalDemandActuals.xlsx"

# read tab file
df = pd.read_csv(tab_file, sep="\t")

# write to Excel
df.to_excel(excel_file, index=False)

print("Conversion completed")

