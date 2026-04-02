
# import csv

# filename = "C:/Users/SW526XH/Downloads/Data Quality Check/tabfiles/Part (FG).tab"

# with open(filename, newline="", encoding="utf-8") as f:
#     reader = csv.reader(f, delimiter="\t")
#     next(reader, None)  # skip header
#     row_count = sum(1 for _ in reader)

# print(row_count)


# import pandas as pd

# # input and output files
# tab_file = "C:/Users/SW526XH/Downloads/Data Quality Check/tabfiles/Cutomer.tab"
# excel_file = "Customer.xlsx"

# # read tab file
# df = pd.read_csv(tab_file, sep="\t")

# # write to Excel
# df.to_excel(excel_file, index=False)

# print("Conversion completed")

import pandas as pd

# Read only the 4 required columns from the .tab file
cols = ["MATERIAL", "PLANT", "SOLDTOPART", "BILLING_WEEK_START"]
df = pd.read_csv('your_file.tab', sep='\t', usecols=cols)

# Split into chunks of 10 lakh rows (1,000,000) per sheet
chunk_size = 1_000_000

with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    for i, start in enumerate(range(0, len(df), chunk_size)):
        chunk = df[start:start + chunk_size]
        chunk.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)

print("Done!")
