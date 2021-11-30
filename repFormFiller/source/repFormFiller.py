#!/usr/bin/python3
""" CATS Tutor Report Form Filler
	
	This script reads in an input calendar file (.ics) and textfile of student names to
	output a Javascript file to copy/paste into a browser code insertion tool such as
	Tampermonkey. This allows a tutor to easily fill out the majority of the Tutor Report 
	Form with the use of dropdown selections. 

	requires:
		* the following files to be in the runpath of timesheetGen.py:
			* calendar2csv.py
			* csv2timesheet.py
"""
import sys
import csv
import argparse
from datetime import date
from calendar2csv import calendar2csv as cal2csv
from csv2timesheet import namesFile2list as nfile2list

def _printList(studentList):
	for stu in studentList:
		print(stu)

def repFormFiller(inputICS, namesFile='none'):
	"""Given an input calendar (.ics) file, returns a string of Javascript for filling in CATS tutor report forms
	
	Parameters
	~~~~~~~~~~
	inputICS : str
		The input Google calendar (.ics) file

	namesFile : str, optional
		The input (.txt) file of tutor's and students' names in the format:
			tutor:
			[tutor's last name], [tutor's first name]
			students:
			[student's last name], [student's first name]
			...
	Returns
	~~~~~~~
	file(.js)
		Javascript file to copy/paste into a browser code insertion tool such as Tampermonkey
	"""

	# students = [
	# 	{student}
	# ]
	# student = {
	# 	'lastName': lastName
	# 	'sport': sport
	# 	'classNames': {className}
	# 	'startTimes': {startTime}
	# }
	students = []
	fullNames = []
	if not namesFile == 'none':
		namesList = nfile2list(namesFile)
		fullNamesStrings = namesList[1]
		for name in fullNamesStrings:
			fullNames.append(name.split(','))

	# generate csv file of meetings from input calendar file
	todayDate = date.today()
	todayDateStr = todayDate.strftime("%m/%d/%y")
	print(todayDateStr)
	inputCSV = cal2csv(inputICS, todayDateStr)
	
	with open(inputCSV, 'r') as inCSV, open('outputJS.js', 'w') as outputText:
		csvReader = csv.reader(inCSV)
		for row in csvReader:
			if not len(row) == 0 and not row[0] == "Date":	# skip past csv header and blank rows
				lastName, sport, className, startTime = [row[1], row[2], row[3], row[4]]

				# replace any "M/W" in sport with "Men's/Women's"
				if sport.startswith("W ") or sport.startswith("M "):
					sportSplit = sport.split(' ')
					if sportSplit[0] == 'W':
						sport = "Women's " + sportSplit[1]
					else:
						sport = "Men's " + sportSplit[1]

				# if not already there, add student to list of students
				if next((stu for stu in students if stu['lastName'] == lastName), None) == None:
					newStudent = {'lastName': lastName, 'sport': sport, 'classNames': [], 'startTimes': []}
					students.append(newStudent)

				# add startTime of session
				for stu in students:
					if stu['lastName'] == lastName:
						# append startTime to list of start times for student
						if startTime not in stu['startTimes']:
							stu['startTimes'].append(startTime)
						# append className to list of classes for student
						if className not in stu['classNames']:
							stu['classNames'].append(className)

		#_printList(students)

		# create javascript string
		outputJS = """
			// ==UserScript==
			// @name         CATS Tutor Report Form Filler
			// @namespace    http://tampermonkey.net/
			// @version      0.1
			// @description  fill in form with dropdown options from schedule
			// @author       Dimitri Mojsejenko
			// @match        https://uky.az1.qualtrics.com/jfe/form/SV_1B7bgZ8gYY0WOuF
			// @grant        none
			// @require      http://code.jquery.com/ui/1.12.1/jquery-ui.js
			// ==/UserScript==
			""".replace('\t','')

		outputJS += "var students = " + str(students) + ';\n'
		outputJS += "var fullNames = " + str(fullNames) + ';'
		outputJS += """
			var pageBody = document.querySelector("body");

			var sbmtButton = document.createElement("DIV");
			sbmtButton.innerHTML = '<input type="button" value="Submit" id="submit">';
			pageBody.prepend(sbmtButton);
			document.getElementById('submit').onclick = function() {
				document.getElementById("NextButton").click();
			}

			var popButton = document.createElement("DIV");
			popButton.innerHTML = '<input type="button" value="Populate" id="populate">';
			pageBody.prepend(popButton);
			document.getElementById('populate').onclick = function() {populateFields()};

			var classesDrpdwn = document.createElement("DIV");
			classesDrpdwn.innerHTML = '<select name="classes" id="classesDrpdwn">';
			pageBody.prepend(classesDrpdwn)

			var sTimesDrpdwn = document.createElement("DIV");
			sTimesDrpdwn.innerHTML = '<select name="startTimes" id="sTimesDrpdwn">';
			pageBody.prepend(sTimesDrpdwn);

			var studentFnameDrpdwn = document.createElement("DIV");
			studentFnameDrpdwn.innerHTML = '<select name="studentFname" id="studentFnameDrpdwn">';
			pageBody.prepend(studentFnameDrpdwn);

			var studentLnameDrpdwn = document.createElement("DIV");
			studentLnameDrpdwn.innerHTML = '<select name="studentLname" id="studentLnameDrpdwn">';
			pageBody.prepend(studentLnameDrpdwn);

			var image = new Image();
		    // previously saved base64 png data of signature; use getDataURL() method
		    image.src='';
		    
			window.onload = function() {
				console.log(students);
				var studentLnameSel = document.getElementById("studentLnameDrpdwn");
				var studentFnameSel = document.getElementById("studentFnameDrpdwn");
				var sTimesSel = document.getElementById("sTimesDrpdwn");
				var classSel = document.getElementById("classesDrpdwn");
				// use 'of' for traversing arrays in JS; 'in' for dictionaries(objects)
				for (var stu of students) {
					var lastName = stu['lastName'];
					studentLnameSel.options[studentLnameSel.options.length] = new Option(lastName, lastName);
				}
				studentLnameSel.onchange = function() {
					// empty dropdowns
					sTimesSel.length = 1;
					classSel.length = 1;
					studentFnameSel.length = 1;
					for (var name of fullNames) {
						if (name[0] == this.value) {
							studentFnameSel.options[studentFnameSel.options.length] = new Option(name[1], name[1]);
						}
					}
					// display start times and classes for chosen student
					for (var stud of students) {
						if (stud['lastName'] == this.value) {
							for (var t of stud['startTimes']) {
								sTimesSel.options[sTimesSel.options.length] = new Option(t, t);
							}
							for (var c of stud['classNames']) {
								classSel.options[classSel.options.length] = new Option(c, c);
							}
						}
					}
				} //studentLnameSel.onchange
			} //window.onload

			function populateFields() {
				// create date object to get today's date in proper format
	    		const date = new Date()
			    var dateFmtd = new Intl.DateTimeFormat('en-GB', {month: '2-digit', day: '2-digit', year: 'numeric'}).format(date);
			    var timeFmtd = new Intl.DateTimeFormat('en-GB', {hour: '2-digit', minute: '2-digit'}).format(date);

			    // check if first page is displaying
			  	if (document.getElementById("QR~QID1~4")) {
				    var stuLastName = document.getElementById("studentLnameDrpdwn").value;
				    var stuFirstName = document.getElementById("studentFnameDrpdwn").value;
				    var startTime = document.getElementById("sTimesDrpdwn").value;
				    var className = document.getElementById("classesDrpdwn").value;
				    // format startTime
				    var curYear = date.getYear();
				    var curMonth = date.getMonth();
				    var curDay = date.getDate();
				    var startTimeHour = startTime.slice(0,2);
				    const sTimeDate = new Date(curYear, curMonth, curDay, startTimeHour, 0, 0);
				    var startTimeFmtd = new Intl.DateTimeFormat('en-GB', {hour: '2-digit', minute: '2-digit'}).format(sTimeDate);
				    var startTimeAMPM = (startTimeHour >= 12) ? "PM" : "AM";

				    // fill fields for chosen student and time
				    document.getElementById("QR~QID1~5").value = stuLastName;
				    for (var stu of students) {
				    	if (stu['lastName'] == stuLastName) {
				    		document.getElementById("QR~QID1~4").value = stuFirstName;
				    		document.getElementById("QR~QID15").value = className;
						    document.getElementById("QR~QID13").value = dateFmtd;
						    document.getElementById("appt-time").value = startTimeFmtd;
						    document.getElementById("QR~QID17").value = startTimeAMPM;
						    // search dropdown option for chosen student's sport
						    var sportDrpdwn = document.getElementById("QR~QID3");
						    selDrpdwnTxt(sportDrpdwn, stu['sport']);
				    	}
				    }
				    document.getElementById("QR~QID18~1").value = "Dimitri";
				    document.getElementById("QR~QID18~2").value = "Mojsejenko";  
			  	} 
			  	else if (document.getElementById("QR~QID23~1")) {
			  	    document.getElementById("QR~QID23~1").checked = true;
			  	    document.getElementById("appt-time").value = timeFmtd;
			  	    document.getElementById("QR~QID10~1").checked = true;
			  	    document.getElementById("QR~QID19~4").checked = true;
			  	}
			  	else if (document.getElementById("QR~QID25~1")) {
			  		document.getElementById("QR~QID25~1").checked = true;
				    document.getElementById("QR~QID26~1").checked = true;
				    document.getElementById("QR~QID20").value = "hw";
				    document.getElementById("QR~QID21").value = "no";
				    var canvas = document.getElementById("QID22-Signature");
				    var ctx = canvas.getContext("2d");
				    ctx.drawImage(image, 0, 0, image.width, image.height, 0, 0, canvas.width, canvas.height);
				    //--- Simulate a natural mouse-click sequence.
				    triggerMouseEvent(canvas, "mouseover");
				    triggerMouseEvent(canvas, "mousedown");
				    triggerMouseEvent(canvas, "mouseup");
				    triggerMouseEvent(canvas, "click");
				}
					
			}
			function triggerMouseEvent(node, eventType) {
			    var clickEvent = document.createEvent ('MouseEvents');
			    clickEvent.initEvent (eventType, true, true);
			    node.dispatchEvent (clickEvent);
			}
			function selDrpdwnTxt(docElement, txt) {
				Array.from(docElement.options).forEach(function(optionElement) {
					if (optionElement.text == txt) {
						optionElement.selected = true;
					}
				})
			}
		""".replace('\t','')


		print("outputJS:")
		print(outputJS)
		outputText.write(outputJS)

if __name__ == "__main__":
	argParser = argparse.ArgumentParser()
	argParser.add_argument(
		"inputICS",
		type=str,
		help="The input Google Calendar (.ics) file"
	)
	argParser.add_argument(
		"-n", "--namesFile",
		type=str,
		default='none',
		help="""The input (.txt) file of tutor's name followed by students' names"""
	)
	args = argParser.parse_args()


	repFormFiller(args.inputICS, args.namesFile)