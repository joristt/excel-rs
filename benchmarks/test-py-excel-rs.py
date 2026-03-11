from py_excel_rs import WorkBook

wb = WorkBook()

f = open("organizations-1000000.csv", "rb")

wb.write_csv_to_sheet(
    "Organizations",
    f.read(),
)

with open("report.xlsx", "wb") as f:
    wb.finish(f)
