# import pandas as pd
from tkinter import Tk
from os import getcwd, path
from geocoder import google
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from tkinter.filedialog import askopenfilename

root = Tk()
root.withdraw()
file_path = askopenfilename(initialdir=getcwd(), title="Select an Excel file.",
                            filetypes=[("Microsoft Excel Worksheet", ".xlsx")])
wb = load_workbook(file_path)
site_sheet = wb['Site']
cell_errors = []

def find_address_column():
    for row in site_sheet.iter_rows(min_row=0, max_row=1):
        for index, cell in enumerate(row):
            if cell.value == "Site Address":
                return get_column_letter(index)

def cell_operations():
    address_column = find_address_column()
    latitude_column_letter = get_column_letter(site_sheet.max_column + 1)
    longitude_column_letter = get_column_letter(site_sheet.max_column + 2)
    for index, cell in enumerate(site_sheet[address_column]):
        try:
            latlng = google(cell.value).latlng
            lat, long = latlng[0], latlng[1]
            site_sheet[latitude_column_letter + str(index + 1)] = lat
            site_sheet[longitude_column_letter + str(index + 1)] = long
        except ValueError:
            print('Could not read address correctly.')
            cell_errors.append((cell, cell.value))
    site_sheet[latitude_column_letter + '1'] = "Latitude"
    site_sheet[longitude_column_letter + '1'] = "Longitude"

    print("These addresses could not be converted to latitude-longitude coordinates.")
    print(cell_errors)

    file_name = path.basename(file_path)
    index = file_name.find(".xlsx")
    wb.save(file_name[:index] + '_MODIFIED' + file_name[index:])

if __name__ == '__main__':
    cell_operations()