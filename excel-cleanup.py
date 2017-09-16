from openpyxl import load_workbook
wb = load_workbook('data.xlsx')
site_sheet = wb['Site']

for cell in site_sheet['D']:
    print(cell.value)
