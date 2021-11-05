# CATStutorTools
Useful tools for CATS tutors.

## Timesheet Generator
Reads in a calendar file (.ics) and textfile of student names to output a MS Word document (.docx) in CATS timesheet format.
Contains:
	- timesheetGenGUI.exe
	Executable which opens up a GUI where the user can choose files and dates from which at timesheet is generated
	- timesheetTemplate.docx
	Timesheet template document which must be in same location that timesheetGenGUI.exe is run
	- source
	Python source code used in timesheetGenGUI.exe


## Report Form Filler
Reads in an input calendar file (.ics) and textfile of student names to
output a Javascript file to copy/paste into a browser code insertion tool such as
Tampermonkey. This allows a tutor to easily fill out the majority of the Tutor Report 
Form with the use of dropdown selections.
Contains:
	- source
	Python source code
	TODO: executable
