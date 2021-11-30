#!/usr/bin/python3
"""Calendar to CSV Parser

This script parses a google calendar file (.ics) based on input dates to produce
a (.csv) file of meetings. It is intended for use with CATS tutor meetings with
the summary (meeting name) formatted as:
	tutorLastName-studentLastName-Course-Sport

command line usage:
	python3 calendar2csv.py inputICS -s [startDate] -e [endDate]
	- where inputICS is the user provided (.ics) file
	- [startDate] and [endDate] are optional arguments to provide the starting
	  and ending dates of the window to be parsed
	- date format: MM/DD/YYYY
	- If no dates are provided, the full calendar will be parsed
	- If only start date provided, then 2 weeks from date will be parsed
	- If only end date provided, then full calendar up to date will be parsed

This file can be used as a standalone script or imported as a module. If used as
a script, an output file (.csv) will be created containing the found CATS tutor
meetings following the headers:
	Date, Student, Sport, Course, StartTime, EndTime

The output file will be named dependent on the input dates and will have the format:
	meetings_[startDate]_[endDate].csv
For example, running:
	python3 calendar2csv.py infile.ics -s 01/28/1989 -e 06/10/2012
produces the output file "meetings_1_28_to_6_10.csv".

If imported as a module, the following functions are also available:
	* date2dayNtime - given a datetime object, returns a list of the date and time
	* dateStr2Obj - given a date string, returns a datetime object
"""

import sys
import argparse
import csv
import pytz
from dateutil import rrule
from dateutil.parser import parse
from datetime import *

def calendar2csv(inputICS, startDate='01/01/1970', endDate='12/31/9999'):
	"""Given a Google calendar (.ics) file, returns a list or (.csv) file of meetings
	
	Parameters
	~~~~~~~~~~
	inputICS : str
		The input Google calendar (.ics) file
	startDate : str, optional
		The starting date of window to extract meetings from, formatted as: MM/DD/YYYY
	endDate : str, optional
		The ending date of window to extract meetings from, formatted as: MM/DD/YYYY

	Returns
	~~~~~~~
	file(.csv)
		A (.csv) file of meetings with headers
	"""

	# name output file with given dates
	sDateSplit = [s.lstrip('0') for s in startDate.split('/')] # remove any leading 0's
	eDateSplit = [e.lstrip('0') for e in endDate.split('/')]
	outFname = "meetings_"+'_'.join(sDateSplit[:2])+'_to_'+'_'.join(eDateSplit[:2])+'.csv'

	# list of meetings where
	# 	meeting = [date, student, sport, course, sTime, eTime]
	calndrList = []

	# convert input dates to datetime.date objects
	startDate = dateStr2Obj(startDate)
	endDate = dateStr2Obj(endDate)

	with open(inputICS, 'r') as calndr:
		for line in calndr:
			# every meeting begins with "BEGIN:VEVENT"
			if "BEGIN:VEVENT" in line:
				meeting = {'dtStart': '', 'eTime': '', 'rrule': '', 'exDate': [], 'summ': '', 'recID': ''}
				dtstart = calndr.readline().strip()
				meeting['dtStart'] = 'DTSTART:' + dtstart.split(':')[1].rstrip('Z') # making sure to remove timezone info
				line = calndr.readline() # advance to next line

				# find meeting end date/time
				if 'T' in line.split(':')[1]: # check that times are included
					dtEndSplit = line.split(':')[1].split('T')
					etUnfmtd = dtEndSplit[1].strip().rstrip('Z')
					meeting['eTime'] = etUnfmtd[0:2]+':'+etUnfmtd[2:4]
				else:
					# print('WARNING: Times not found for meeting:\n',)
					meeting['eTime'] = 'NaN'
				
				line = calndr.readline()

				# check if meeting is recurring
				if "RRULE" in line:
					meeting['rrule'] = line.split(':')[1].replace('Z','').strip()
					line = calndr.readline()

				# check if any days are excluded
				while "EXDATE" in line:
					exdate = line.strip()
					exdate = 'EXDATE:' + exdate.split(':')[1].rstrip('Z') # making sure to remove timezone info
					meeting['exDate'].append(exdate)
					line = calndr.readline()

				# skip to meeting summary, or "RECURRENCE-ID" if present
				while not (line.startswith("SUMMARY") or line.startswith("RECURRENCE-ID")):
					line = calndr.readline()
				if line.startswith("RECURRENCE-ID"):
					meeting['recID'] = line.split(':')[1].strip()
					while not (line.startswith("SUMMARY")):
						line = calndr.readline()
				if line.startswith("SUMMARY"):
					meeting['summ'] = line.split(':')[1].strip()

				calndrList.append(meeting)

	# sort calendar list by summary
	sortedCalndrList = sorted(calndrList, key = lambda i: i['summ'])

	# generate list of sessions to output
	mtgList = []

	for mtgSet in sortedCalndrList:
		smrySplit = mtgSet['summ'].strip().split('-')
		# check if it's a recurring meeting
		if mtgSet['rrule'].strip(): # if rrule isn't empty string
			mtgDays = rrule.rruleset()
			ruleString = mtgSet['dtStart']+'\n'+mtgSet['rrule']+'\n'
			for xdt in mtgSet['exDate']:
				ruleString += xdt + '\n'				
			if mtgSet['exDate']:
				ruleString.rstrip('\n')
			ruleString.rstrip('\n')
			mtgDays.rrule(rrule.rrulestr(ruleString))

			for mtgday in mtgDays:
				# check if meeting date is within given starting and ending dates
				if (startDate <= mtgday.date() <= endDate):
					dateTimeStr = date2dayNtime(mtgday)
					sDate = dateTimeStr[0]
					sTime = dateTimeStr[1] 
					if len(smrySplit) > 3:
						mtgList.append([sDate, smrySplit[1].strip(), smrySplit[2].strip(), smrySplit[3].strip(), sTime, mtgSet['eTime'], mtgSet['recID']])
					else:
						# handle cases of incorrectly formatted summary
						print("WARNING: Meeting summary incorrectly formatted and was not included in output file: ",smrySplit)
						print(" Correct format: tutorLastName-studentLastName-Course-Sport")
		# if no rrules, then just a single meeting
		else:
			ruleString = mtgSet['dtStart']+'\n'+'RRULE:FREQ=DAILY;COUNT=1' # freq still required for single session using rrule
			mtgday = rrule.rruleset()
			mtgday.rrule(rrule.rrulestr(ruleString)) # mtgday should be list with single datetime object
			# check if meeting date is within given starting and ending dates
			if (startDate <= mtgday[0].date() <= endDate):
				dateTimeStr = date2dayNtime(mtgday[0])
				sDate = dateTimeStr[0]
				sTime = dateTimeStr[1] 
				if len(smrySplit) > 3:
					mtgList.append([sDate, smrySplit[1].strip(), smrySplit[2].strip(), smrySplit[3].strip(), sTime, mtgSet['eTime'], mtgSet['recID']])
				else:
					# handle cases of incorrectly formatted summary
					print("WARNING: Meeting summary incorrectly formatted and was not included in output file: ",smrySplit)
					print(" Correct format: tutorLastName-studentLastName-Course-Sport")

	# replace meetings referenced by recID's; these are meetings that had their times moved
	mtg2replace = []
	newMtgs = []
	for mtg in mtgList:
		if mtg[6]:
			mtgDayNTime = date2dayNtime(parse(mtg[6]))
			newMtgs.append([mtgDayNTime[0],mtg[1],mtg[2],mtg[3],mtg[4],mtg[5]])
			mtg2replace.append([mtgDayNTime[0],mtg[1],mtg[2],mtg[3],mtgDayNTime[1]])

	outList = []
	for mtg in mtgList:
		if mtg[:5] not in mtg2replace:
			outList.append(mtg[:6])
	
	for mtg in newMtgs:
		if mtg not in outList:
			outList.append(mtg)

	# sort output rows by date and time
	outList.sort(key=lambda i: (dateStr2Obj(i[0]),i[4]))

	# generate output csv
	with open(outFname, 'w') as outputCSV:
		# create header fields for csv
		fieldnames = ['Date', 'Student', 'Sport', 'Course', 'StartTime', 'EndTime']
		csvwriter = csv.DictWriter(outputCSV, fieldnames=fieldnames)
		csvwriter.writeheader()
		count = 0
		print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
		for mtg in outList:
			count += 1
			print(mtg)
			csvwriter.writerow({'Date': mtg[0], 'Student': mtg[1], 'Sport': mtg[2], 'Course': mtg[3], 'StartTime': mtg[4], 'EndTime': mtg[5]})
		print("Total sessions: ",count)

		if __name__ == '__main__':
			print("Output file created: ", outFname)
		else:
			return outFname

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
def date2dayNtime(datetimeObj):
	"""Given a datetime object, returns a list of strings [date, time]"""
	datetimeObj.replace(tzinfo=None) # remove timezone info which complicates things
	dateStr = str(datetimeObj.month)+'/'+str(datetimeObj.day)+'/'+str(datetimeObj.year)
	if datetimeObj.hour < 10:
		hourStr = '0'+str(datetimeObj.hour)
	else:
		hourStr = str(datetimeObj.hour)
	if datetimeObj.minute < 10:
		minStr = '0'+str(datetimeObj.minute)
	else:
		minStr = str(datetimeObj.minute)	
	timeStr = hourStr+':'+minStr
	return [dateStr, timeStr]

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
def dateStr2Obj(datestring):
	"""Given a date string formatted as "MM/DD/YYYY", returns a date object"""
	if not datestring.count('/') == 2:
		print("Date must be formatted as MM/DD/YYYY\n")
	else:
		dateSplit = datestring.split('/')
		month = int(dateSplit[0])
		day = int(dateSplit[1])
		year = int(dateSplit[2])
		try:
			dateObj = date(month=month, day=day, year=year)
		except ValueError:
			print("Undefined date given: ",datestring,"\n")
		return dateObj

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
	args = argParser.parse_args()

	startDate = dateStr2Obj(args.startDate)
	# if startDate is set and not endDate, set endDate to 2 weeks past the startDate
	if (not args.startDate == '01/01/1970' and args.endDate == '12/31/9999'):
		endDate = startDate + timedelta(days=13) # 13 to exclude 3rd instance of starting day
	else:
		endDate = dateStr2Obj(args.endDate)
	endDate = str(endDate.month)+'/'+str(endDate.day)+'/'+str(endDate.year)

	calendar2csv(args.inputICS, args.startDate, endDate)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if __name__ == "__main__":
	main()