import lansweeper_update
import xlrd
import re

# Regex for Dell serial or Surface tablet serial
regex = r"(^[a-zA-Z0-9]{7}$)|(^[0-9]{12}$)"

# Open and read ExpenseReports.xlsx. Read from 1st worksheet
workbook = xlrd.open_workbook(r"\\InaGalaxyFarFarAway.xls")
worksheet = workbook.sheet_by_index(0)
total_rows = worksheet.nrows


for i in range(6, total_rows):

    asset = worksheet.cell(i, 0).value
    po = worksheet.cell(i, 1).value
    serial = worksheet.cell(i, 6).value.strip()

    try:
        if re.search(regex, serial):
            asset = int(asset)
            po = int(po)
            print(asset, po, serial)
            lansweeper_update.Sweep(asset, po, serial)

    except Exception as e:
        print(e)
