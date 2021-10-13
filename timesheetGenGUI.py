#!/usr/bin/python3
"""CATS Timesheet Generator

This script reads in a Google Calendar file and outputs a MSWord document (.docx) 
in CATS timesheet format. It is intended for use by CATS tutors to quickly generate
timesheets based on meetings scheduled in their Google Calendar. To be included in the
output timesheet, the meetings should have their summary (meeting name) formatted as:
	tutorLastName-studentLastName-Course-Sport
If the session is with multiple students, use '/' to separate names.

command line usage:
	python3 timesheetGen.py inputICS -s [startDate] -e [endDate] -n [namesFile] -c
	- where inputICS is the user provided (.ics) file
	- [startDate] and [endDate] are optional arguments to provide the starting
	  and ending dates of the window to be parsed
	- date format: MM/DD/YYYY
	- If no dates are provided, the full calendar will be parsed
	- If only start date provided, then 2 weeks from date will be parsed
	- If only end date provided, then full calendar up to date will be parsed
	- [namesFile] is an optional argument to provide a file (.txt) of names with
	  the format:
	  	tutor:
		[tutor's last name], [tutor's first name]
		students:
		[student's last name], [student's first name]
		...
	- The optional flag -c can be included to keep the generated (.csv) file that
	  is otherwise deleted after use

The output file will be named dependent on the input dates and the file of names
if included. If the namesFile is included, the output file will be named:
	[tutor's last name]_timesheet_[startDate]_to_[endDate].docx
Otherwise the output file will be named:
	timesheet_[startDate]_to_[endDate].docx
**Note: the [namesFile] option won't work properly if students share a last name

example use:
	python3 timesheetGen.py infile.ics -s 08/15/2021
This finds all the CATS tutor meetings in infile.ics within a window of two weeks 
from 08/15/2021. The only output file generated will be:
	timesheet_08_15_to_08_29.docx

requires:
	* the following files to be in the runpath of timesheetGen.py:
		* calendar2csv.py
		* csv2timesheet.py
		* timesheetTemplate.docx
"""

import os
import argparse
from calendar2csv import calendar2csv as cal2csv
from calendar2csv import dateStr2Obj
from csv2timesheet import csv2timesheet as csv2ts
from csv2timesheet import checkFormat
from datetime import *
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkcalendar import DateEntry

def timesheetGen(inputICS, startDate='01/01/1970', endDate='12/31/9999', namesFile='none', keepCSV=False):
	"""Given an input Google Calendar file, returns a (.docx) file in CATS timesheet format
	
	Parameters
	~~~~~~~~~~
	inputICS: str
		The input Google calendar (.ics) file
	startDate : str, optional
		The starting date of window to extract meetings from, formatted as: MM/DD/YYYY
	endDate : str, optional
		The ending date of window to extract meetings from, formatted as: MM/DD/YYYY
	namesFile: str, optional
		The input (.txt) file of tutor's and students' names in the format:
			tutor:
			[tutor's last name], [tutor's first name]
			students:
			[student's last name], [student's first name]
			...
	keepCSV: bool, optional
		An optional flag that when True, the generated (.csv) file will not be deleted
	
	Returns
	~~~~~~~
	file(.docx)
		A MSWord document of the meetings in CATS timesheet format is created in the
		run directory. The name of the output document is returned.
	"""

	inCSV = cal2csv(inputICS, startDate, endDate)
	outDoc = csv2ts(inCSV, namesFile)
	if not keepCSV:
		if os.path.exists(inCSV): 
			os.remove(inCSV)
	else:
		print("Output (.csv) file created: ", inCSV)
	return outDoc # return name of output document

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
def _chooseFile(tkEntry, ftype):
	"""opens user dialogue asking for file, and stores name of file into given Tkinter entry"""
	if ftype == ".ics":
		fname = askopenfilename(filetypes=[("Calendar File", ".ics")])
		txt = "calendar (.ics) file"
	else:
		fname = askopenfilename()
		txt = "names file"
	tkEntry.delete(0,END)
	tkEntry.insert(0,fname)
	if os.path.exists(fname):
		message.configure(text="") # reset notification message
		if not ftype == ".ics":
			if not checkFormat(fname):
				message.configure(text="The given namesFile doesn't have the correct format and won't be used.")
	else:
		message.configure(text="Please choose a valid "+txt)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
def _setEndDate(startDate, endCal):
	"""sets the Ending Date Entry to 14 days past the Starting Date"""
	endCal.set_date(startDate + timedelta(days=13)) # 13 to exclude 3rd instance of starting day)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
def _gatherVars():
	"""gathers all the variables that will be input to timesheetGen() and returns them as list"""
	try:
		inputICS = inFile.get()
	except:
		inputICS = ""
	sDateObj = startCal.get_date()
	startDate = str(sDateObj.month)+'/'+str(sDateObj.day)+'/'+str(sDateObj.year)
	eDateObj = endCal.get_date()
	endDate = str(eDateObj.month)+'/'+str(eDateObj.day)+'/'+str(eDateObj.year)
	try:
		namesFile = names.get()
	except:
		namesFile = "none"

	return [inputICS, startDate, endDate, namesFile]

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
def _generateTimeSheet():
	"""gathers variables and generates time sheet with timesheetGen()"""
	args = _gatherVars()
	if not os.path.exists(args[0]):
		message.configure(text="Please choose a valid calendar (.ics) file")
	else:
		outputDoc = timesheetGen(args[0],args[1],args[2],args[3])
		message.configure(text="Output file created:   "+outputDoc)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Main GUI loop
# arguments and defaults
defaults = ["", "01/01/1970", "12/31/9999", "none"]
inputICS, startDate, endDate, namesFile = defaults[0], defaults[1], defaults[2], defaults[3]

# create Tkinter GUI
root = Tk()
root.geometry('800x270')
root.title("CATS Timesheet Generator")

# first row
row1 = Frame(root)
lab1 = Label(row1, width=18, text="Input Calendar File", anchor='w')
inFile = StringVar()
ent1 = Entry(row1, textvariable=inFile)
fileBut = Button(row1, text="Choose File", anchor='e')
fileBut.bind('<Button>', lambda fButHandler: _chooseFile(ent1,".ics")) # widget.bind(event, handler)
# when using bind, tkinter needs an argument containing the event with lambda
row1.pack(side=TOP, fill=X, padx=5, pady=5)
lab1.pack(side=LEFT)
ent1.pack(side=LEFT, expand=YES, fill=X)
fileBut.pack(side=RIGHT)

# second row
row2 = Frame(root)
lab2 = Label(row2, width=18, text="Starting Date", anchor='w')
startCal = DateEntry(row2, width=10)
startCal.bind('<<DateEntrySelected>>', lambda dateHandler: _setEndDate(startCal.get_date(), endCal))
row2.pack(side=TOP, fill=X, padx=5, pady=5)
lab2.pack(side=LEFT)
startCal.pack(side=LEFT, padx=10, pady=10)

# third row
row3 = Frame(root)
lab3 = Label(row3, width=18, text="Ending Date", anchor='w')
endCal = DateEntry(row3, width=10)
row3.pack(side=TOP, fill=X, padx=5, pady=5)
lab3.pack(side=LEFT)
endCal.pack(side=LEFT, padx=10, pady=10)	

# fourth row
row4 = Frame(root)
lab4 = Label(row4, width=18, text="Input Names File", anchor='w')
names = StringVar(value="none")
ent4 = Entry(row4, textvariable=names)
file2But = Button(row4, text="Choose File", anchor='e')
file2But.bind('<Button>', lambda fButHandler: _chooseFile(ent4, ".txt"))
row4.pack(side=TOP, fill=X, padx=5, pady=5)
lab4.pack(side=LEFT)
ent4.pack(side=LEFT, expand=YES, fill=X)
file2But.pack(side=RIGHT)

# message
message = Label(root, text="")
message.pack(side=BOTTOM, padx=5, pady=10)

# timesheetGen button - will show up in GUI above message
genBut = Button(root, text="Generate Timesheet", anchor='n')
genBut.bind('<Button>', lambda genButHandler: _generateTimeSheet())
genBut.pack(side=BOTTOM, padx=5, pady=5)

root.mainloop()