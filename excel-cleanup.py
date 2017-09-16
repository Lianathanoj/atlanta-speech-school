from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('data.xlsx')
site_sheet = wb['Site']

def find_address_column():
    for row in site_sheet.iter_rows(min_row=0, max_row=1):
        for index, cell in enumerate(row):
            if cell.value == "Site Address":
                return get_column_letter(index)

def cell_operations():
    address_column = find_address_column()
    for cell in site_sheet[address_column]:
        print(cell.value)

if __name__ == '__main__':
    cell_operations()