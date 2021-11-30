#!/usr/bin/python3
"""CATS Timesheet Generator

This script reads in a Google Calendar file and outputs a MSWord document (.docx) 
in CATS timesheet format. It is intended for use by CATS tutors to quickly generate
timesheets based on meetings scheduled in their Google Calendar. To be included in the
output timesheet, the meetings should have their summary (meeting name) formatted as:
	tutorLastName-studentLastName-Course-Sport

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
from datetime import *

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
		A MSWord document of the meetings in CATS timesheet format
	"""

	inCSV = cal2csv(inputICS, startDate, endDate)
	csv2ts(inCSV, namesFile)
	if not keepCSV:
		if os.path.exists(inCSV):
			os.remove(inCSV)
	else:
		print("Output file created: ", inCSV)

def main():
	argParser = argparse.ArgumentParser()
	argParser.add_argument(
		"inputICS",
		type=str,
		help="The input Google Calendar (.ics) file"
	)
	argParser.add_argument(
		"-s", "--startDate",
		nargs='?', # 0 or 1 argument
		type=str,
		default="01/01/1970",
		help="The starting date of window to extract meetings from, formatted as: MM/DD/YYYY"
	)
	argParser.add_argument(
		"-e", "--endDate",
		nargs='?',
		type=str,
		default="12/31/9999",
		help="The ending date of window to extract meetings from, formatted as: MM/DD/YYYY"
	)
	argParser.add_argument(
		"-n", "--namesFile",
		type=str,
		default='none',
		help="""The input (.txt) file of tutor's name followed by students' names"""
	)
	argParser.add_argument(
		"-c", "--csv",
		action='store_true',
		help="""An option flag to save meeting info to a (.csv) file"""
	)
	#TODO: option to delete (.csv) file after run
	args = argParser.parse_args()

	startDate = dateStr2Obj(args.startDate)
	# if startDate is set and not endDate, set endDate to 2 weeks past the startDate
	if (not args.startDate == '01/01/1970' and args.endDate == '12/31/9999'):
		endDate = startDate + timedelta(days=13) # 13 to exclude 3rd instance of starting day
	else:
		endDate = dateStr2Obj(args.endDate)
	endDate = str(endDate.month)+'/'+str(endDate.day)+'/'+str(endDate.year)

	timesheetGen(args.inputICS, args.startDate, endDate, args.namesFile, args.csv)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if __name__ == "__main__":
	main()