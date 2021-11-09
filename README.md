# CATStutorTools
Useful tools for CATS tutors.

## Timesheet Generator
Reads in a calendar file (.ics) and textfile of student names to output a MS Word document (.docx) in CATS timesheet format.

### Contents:

- timesheetGenGUI.exe: executable which opens up a GUI where the user can choose files and dates from which a timesheet is generated

- timesheetTemplate.docx: timesheet template document which must be in same location that timesheetGenGUI.exe is run

- source: Python source code used in timesheetGenGUI.exe

### Guide
For users who don't care to mess with the Python source code, only the following two files need to be downloaded and saved in the 
same folder:
  - timesheetGenGUI.exe
  - timesheetTemplate.docx

The Timesheet Generator requires a calendar file (.ics) as input. With Google Calendar, this can be accessed in the Settings menu 
with the cog symbol in the top right. Then navigate to the Import/Export tab and hit export to save a zipfile containing the 
needed calendar file, named with the user's email address. The Timesheet Generator sorts through only the meetings in this file, 
so it will need to be re-downloaded before each use if any changes have been made to the schedule. 

The Timesheet Generator is intended for use with CATS tutor meetings with the summary (meeting name) formatted as:
	tutorLastName-studentLastName-Course-Sport
If the session is with multiple students, use '/' to separate names. The same sport will be used if there are multiple students in
a session.

Also neccessary to fully run the Timesheet Generator is a textfile of student names, from which first names are matched to last names 
listed in meetings and included in the generated timesheet. This textfile should have the following format:
	  tutor:
		tutor's last name, tutor's first name
		students:
		student's last name, student's first name
		...

The output file will be named dependent on the input dates and the file of names if included. If the namesFile is included, the output
file will be named:
	<tutor's last name>_timesheet_<startDate>_to_<endDate>.docx
Otherwise the output file will be named:
	timesheet_<startDate>_to_<endDate>.docx

**Note that if student's share a last name, their session may not be properly recorded in the timesheet and will need to be manually
  checked and edited if needed.
  
  
## Report Form Filler
Reads in a calendar file (.ics) and textfile of student names to
output a Javascript file to copy/paste into a browser code insertion tool such as
Tampermonkey. This allows a tutor to easily fill out the majority of the Tutor Report 
Form with the use of dropdown selections.

### Contents:

- source: Python source code

TODO: executable
