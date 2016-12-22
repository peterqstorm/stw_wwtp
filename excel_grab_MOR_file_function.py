import xlrd, os, pandas, calendar

def MOR_grab_data(year, month):
	"""
	This will go into MOR files to grab the data and returns the data in a set of libraries
	that will be accessible. Inputs required are year and month, has to have specific folder location.
	"""
	file_name = "{0:} MOR/MOR-{1:} {0:}.xlsx".format(year, calendar.month_name[month].upper())	#setting this up for a loop
	#file_name = '2015 MOR/MOR-JANUARY 2015.xlsx' #original file name
	if not os.path.isfile(file_name):
		print("The file does not exist.")
	x,ndays = calendar.monthrange(year, month)	#gets length of month so we know how many rows of data to import
		
	#opens the workbook	
	reader = xlrd.open_workbook(file_name)
	#lists the sheet names, selects one
	sheet_names = reader.sheet_names()

		
	if not 'MOR' in sheet_names:
		print("no MOR header for sheet names, for ", month, ", ", year)
		raise

	MOR = reader.sheet_by_name('MOR')

		
	row1 = MOR.row(0)	
	row2 = MOR.row(1)
	columns = []

	#importing the data values from the tables into a list of lists that will later be added to dictionary
	for j in range(len(row2)):
		columns.append([])
		for i in range(ndays):
			if not MOR.cell(i+2,j).value == '':
				columns[-1].append(MOR.cell(i+2,j).value)		#this does the last thing being added
			else:
				columns[-1].append(None)
	MOR_data = {}

	current = ''
	MOR_data[''] = {}
	for i,j,col in zip(row1,row2, columns):	
		unit = i.value
		parameter = j.value
		if  unit != "":
			current = unit
			MOR_data[current]= {} 	#creates a new list if it doesn't exit to store info
		MOR_data[current][parameter]= col	#adding a dictionary spot in a dictionary
	return MOR_data



# testing outputs
# year = 2015
# month = 1
# output = MOR_grab_data(year, month)
# Get = output['DIGESTER C']['Digester - C   PH']
# print(Get)

