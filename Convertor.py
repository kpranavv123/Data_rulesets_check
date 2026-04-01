
# import csv

# filename = "C:/Users/SW526XH/Downloads/Data Quality Check/tabfiles/Part (FG).tab"

# with open(filename, newline="", encoding="utf-8") as f:
#     reader = csv.reader(f, delimiter="\t")
#     next(reader, None)  # skip header
#     row_count = sum(1 for _ in reader)

# print(row_count)


import pandas as pd

# input and output files
tab_file = "C:/Users/SW526XH/Downloads/Data Quality Check/tabfiles/Cutomer.tab"
excel_file = "Customer.xlsx"

# read tab file
df = pd.read_csv(tab_file, sep="\t")

# write to Excel
df.to_excel(excel_file, index=False)

print("Conversion completed")

# import pandas as pd
# from openpyxl import load_workbook

# # ---------- CONFIG ----------
# INPUT_TAB_FILE = "C:/Users/SW526XH/Downloads/Data Quality Check/tabfiles/HistoricalDemandActuals.tab"     # your TAB file
# OUTPUT_EXCEL_FILE = "HistoricalDemandActuals.xlsx"
# DELIMITER = "\t"

# EXCEL_MAX_ROWS = 1048576               # Excel limit (including header)
# CHUNK_SIZE = 100000                    # safe chunk size (adjust if needed)
# # ----------------------------

# def tab_to_excel_multi_sheet():
#     sheet_num = 1
#     start_row = 0  # track rows written in current sheet

#     writer = pd.ExcelWriter(
#         OUTPUT_EXCEL_FILE,
#         engine="openpyxl"
#     )

#     for chunk in pd.read_csv(
#         INPUT_TAB_FILE,
#         sep=DELIMITER,
#         dtype=str,          # prevents datatype issues
#         chunksize=CHUNK_SIZE
#     ):
#         remaining_rows = EXCEL_MAX_ROWS - start_row - 1  # -1 for header

#         # If current chunk fits in current sheet
#         if len(chunk) <= remaining_rows:
#             chunk.to_excel(
#                 writer,
#                 sheet_name=f"Sheet{sheet_num}",
#                 index=False,
#                 startrow=start_row,
#                 header=(start_row == 0)
#             )
#             start_row += len(chunk)

#         else:
#             # Split chunk across sheets
#             first_part = chunk.iloc[:remaining_rows]
#             second_part = chunk.iloc[remaining_rows:]

#             # Write first part
#             first_part.to_excel(
#                 writer,
#                 sheet_name=f"Sheet{sheet_num}",
#                 index=False,
#                 startrow=start_row,
#                 header=(start_row == 0)
#             )

#             # Move to next sheet
#             sheet_num += 1
#             start_row = 0

#             # Write remaining part into new sheet
#             second_part.to_excel(
#                 writer,
#                 sheet_name=f"Sheet{sheet_num}",
#                 index=False,
#                 startrow=start_row,
#                 header=True
#             )
#             start_row = len(second_part)

#     writer.close()
#     print("✅ TAB file successfully converted to Excel with multiple sheets")

# if __name__ == "__main__":
#     tab_to_excel_multi_sheet()

