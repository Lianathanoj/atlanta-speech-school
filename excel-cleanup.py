# import pandas as pd
from geocoder import google
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('data - Copy.xlsx')
site_sheet = wb['Site']

def find_address_column():
    for row in site_sheet.iter_rows(min_row=0, max_row=1):
        for index, cell in enumerate(row):
            if cell.value == "Site Address":
                return get_column_letter(index)

cell_errors = []

def cell_operations():
    address_column = find_address_column()
    new_column_letter = get_column_letter(site_sheet.max_column + 1)
    for index, cell in enumerate(site_sheet[address_column]):
        # print(cell.value)
        try:
            # print(cell, google(cell.value).latlng)
            # print(new_column_letter + str(index + 1))
            latlng = google(cell.value).latlng
            lat, long = latlng[0], latlng[1]
            site_sheet[new_column_letter + str(index + 1)] = str(lat) + ', ' + str(long)
        except ValueError:
            print('test')
            cell_errors.append((cell, cell.value))
    site_sheet[new_column_letter + '1'] = "Latitude/Longitude"

    print("These addresses could not be converted to latitude-longitude coordinates.")
    print(cell_errors)

    wb.save('test.xlsx')

if __name__ == '__main__':
    cell_operations()