# This program  is custom made for Prof. Wimmer's research project.
# Run from terminal. 

# 	$ python wimmer.py
# 	What ROW are you starting from? <Enter the number of the ROW on the main excel database>.
#	What ROW do you want to stop at? <Enter the number of the ROW on the main excel database.>

# After the browser opens, hit <ENTER>.
# The link from #ROW will open up with the relevant information displaying in terminal.
# Enter the information for the corresponding questions that show up. 
# This information is written to the master excel database. 

# NOTE: I've not included the following columns:
# Tranlation Problem
# Countries Where Term Appears 
# Source Lang
# Partially Coded by Wimmer
# Name of Coder


from splinter import Browser
from xlutils.display import quoted_sheet_name
from os.path import join
from xlrd import open_workbook
from xlutils.copy import copy

browser = Browser()

book = open_workbook(join('googleterms.xls'), formatting_info=True, on_demand=True)
#testcell = book.sheet_by_index(0).cell(0,15).value
workbook = copy(book)
#print ('The test link is: %s' % (testcell))
start = int(raw_input('What ROW are you starting from? '))
end = int(raw_input('What ROW do you want to stop at? '))
raw_input("Press ENTER to continue")


#with open('googleurls') as myfile:
with book as myfile:

	for i in range(start, end):
		line = book.sheet_by_index(0).cell(i,16).value
		
		browser.visit(line)
		print ("=============================================")
		print ('The link is: %s' % (line)) 
		print ('The translated search term is: %s' % (book.sheet_by_index(0).cell(i, 1)))
		print ('The original search term is: %s' % (book.sheet_by_index(0).cell(i, 2)))
		print ('Country: %s' % (book.sheet_by_index(0).cell(i, 3)))
		print ("---------------------------------------------")

		print ('The current catergory is: %s' % (book.sheet_by_index(0).cell(i, 9)))
		category = raw_input('Give me a category: ')
		if category != "":
			print ('Your category is now %s' % (category))
			#put this in the excel sheet in the right column
			workbook.get_sheet(0).write(i,9, (category))

		fusion = raw_input('Fuse with another term? ')
		if fusion != "":
			print('Term was fused with %s' % (fusion))
			workbook.get_sheet(0).write(i,7, (fusion))

		translate = raw_input('Change translation to -? ')
		if translate != "" :
			print('Term was translated to %s' % (translate))
			workbook.get_sheet(0).write(i,8, (translate))

		eliminate = raw_input('Eliminate the term? Enter 1 for YES. ')
		if eliminate == "1":
			print ('Term was eliminated.')
			workbook.get_sheet(0).write(i,6, ("1"))

		comment = raw_input('Comments? ')
		if comment != "":
			workbook.get_sheet(0).write(i,12, (comment))

		workbook.save(join('googleurls.xls'))

		raw_input("Press ENTER to continue") #moves onto next link

			





