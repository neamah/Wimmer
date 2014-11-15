Wimmer
======
Iteration of predefined Google Search terms with data entry to a master database for each one. 

This program  is custom made for Prof. Wimmer's research project.

==================
NOTE on the excel file you're using

This script can only take xls files. Since the database we have is xlsx, I just Saved As the same file as .xls 
This shouldn't change anything in the contents of the file.

When entering the name of the file in the program, don't enter .xls. JUST the name of the file.

==================
How to run: 

Download this repo and save the excel database in the same folder. 

Run from terminal. 

 	$ python wimmer.py
 	Name of the excel file (without .xls): <Enter the name of the file as you saved it Eg: googleterms>
 	What ROW are you starting from? <Enter the number of the ROW on the main excel database>.
	What ROW do you want to stop at? <Enter the number of the ROW on the main excel database>.

After the browser opens, hit <ENTER>.
The link from #ROW will open up with the relevant information displaying in terminal.
Enter the information for the corresponding questions that show up. 
This information is written to the master excel database. 

NOTE: I've not included the following columns:
1)Tranlation Problem
2)Countries Where Term Appears 
3)Source Lang
4)Partially Coded by Wimmer
5)Name of Coder
==================




