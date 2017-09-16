from openpyxl import load_workbook
wb = load_workbook('data.xlsx')
site_sheet = wb['Site']

for cell in site_sheet['D']: #hard-code find address
    address = cell.value
    addressArray = address.split()
    print(address)
    for index, word in enumerate(addressArray):
    	word = word.strip(',')
    	word = word.strip('.')
    	addressArray[index] = word

for i in addressArray[::-1]:
	print(addressArray[i]) #1: zipcode 2: state 3: city