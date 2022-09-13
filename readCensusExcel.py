# Reading data from a excel.
import openpyxl,pprint
print('opening workbook...')

wb = openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb['Population by Census Tract']

countData = {}
print('reading rows...')
for row in range(2,sheet.max_row+1):
	state = sheet['B'+str(row)].value
	country = sheet['C'+str(row)].value
	pop = sheet['D'+str(row)].value

	# make sure the key for this state exists.
	countData.setdefault(state,{})
	# make sure the key for this country in this state exists.
	countData[state].setdefault(country,{'tract':0,'pop':0})
	countData[state][country]['tract'] += 1
	countData[state][country]['pop'] += int(pop)

# Open a new file and write data to it
print('Writing results...')
resultFile = open('census.py','w')
resultFile.write('allData =' + pprint.pformat(countData))
resultFile.close()
print('done')
# print(country)