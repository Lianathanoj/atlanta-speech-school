import numbers
import decimal
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
wb = load_workbook('data.xlsx')
site_sheet = wb['Site']

zipCol = get_column_letter(site_sheet.max_column + 4)
stateCol = get_column_letter(site_sheet.max_column + 3)
cityCol = get_column_letter(site_sheet.max_column + 2)
addressCol = get_column_letter(site_sheet.max_column + 1)

for cellindex, cell in enumerate(site_sheet['D']): #hard-code find address
	address = cell.value
	addressArray = address.split()
	# print(address)
	for index, word in enumerate(addressArray):
		word = word.strip(',')
		word = word.strip('.')
		addressArray[index] = word
	if addressArray[len(addressArray) - 1].upper() != "GA" and addressArray[len(addressArray) - 1].upper() != "GEORGIA":
		for i, partOfAdd in enumerate(addressArray): #site_sheet.max_column
			if i==len(addressArray) - 1:
				#partOfAdd will have updated to be zip
				site_sheet[zipCol + str(cellindex + 1)] = partOfAdd
			# print(zipCol + str(cellindex + 1))
			# print(partOfAdd)
			# print("zipCol " + str(zipCol))
			elif i==len(addressArray) - 2:
				#will be state
				#converts "Georgia" to "GA"
				if partOfAdd == "Georgia":
					site_sheet[stateCol + str(cellindex + 1)] = "GA"
				else:
					site_sheet[stateCol + str(cellindex + 1)] = partOfAdd.upper()
				# print(stateCol + str(cellindex + 1))
				# print(partOfAdd)
				# print("stateCol " + str(stateCol))
			elif i==len(addressArray) - 3:
				#will be city
				site_sheet[cityCol + str(cellindex + 1)] = partOfAdd
			# print(cityCol + str(cellindex + 1))
			#print(partOfAdd)
			#print("cityCol " + str(cityCol))
		streetAd = addressArray[0]
		for word in addressArray[1:len(addressArray) - 3]:
			streetAd += " " + word
			site_sheet[addressCol + str(cellindex + 1)] = streetAd
			# print(addressCol + str(cellindex + 1))
	else:
		if addressArray[len(addressArray) - 1] == "Georgia":
			site_sheet[stateCol + str(cellindex + 1)] = "GA"
		else:
			site_sheet[stateCol + str(cellindex + 1)] = addressArray[len(addressArray) - 1].upper()
		site_sheet[cityCol + str(cellindex + 1)] = addressArray[len(addressArray) - 2]
		streetAd = addressArray[0]
		for word in addressArray[1:len(addressArray) - 3]:
			streetAd += " " + word
			site_sheet[addressCol + str(cellindex + 1)] = streetAd

wb.save('ndata.xlsx')
