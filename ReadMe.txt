
Line Numbering Tool

'====================================================

To add a right click menu to visual basic projects that allows them to be line numbered then
run the file AddLineNumberingOption.Reg, the contents of this file are shown below....

'====================================================
REGEDIT4

[HKEY_CLASSES_ROOT\.vbp]
@="VisualBasic.Project"

[HKEY_CLASSES_ROOT\VisualBasic.Project\Shell\Line Number VB Project]

[HKEY_CLASSES_ROOT\VisualBasic.Project\Shell\Line Number VB Project\command]
@="C:\\Program Files\\Line Numbering\\linenumbering.exe /P%1 /C /M /W /I10 /T"

'====================================================

Usage: LineNumbering.exe /Pproject /Odirectory  [/C[directory]] [/W] [/Lincrement]

/P -     Project to generate line numbers
/O -     Output directory for new source code (default of \LN)
/C -     Compile the project with line numbers and place out in specified directory
/W -     Wipe the Output directory before starting
/T - 	 Conditional Compilation Contants (e.g. TEST=1:DEBUG=2)
/M -     Maintain the same Path32 (build path) in the new line numbered project 
/L -     Line increment to use (default of 1)
/? -     Display this help text



