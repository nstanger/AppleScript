set theFolder to choose folder with prompt "Choose the folder that contains your Excel files"
tell application "Finder" to set theFiles to (files of theFolder)
set fileCount to count theFiles
repeat with i from 1 to fileCount
	set fName to text 1 thru -5 of ((name of item i of theFiles) as text)
	if ((name of item i of theFiles) as text) ends with ".xls" then
		set tName to (theFolder as text) & fName & ".xlsx"
		tell application "Microsoft Excel"
			activate
			
			open (item i of theFiles) as text
			
			tell active workbook
				save workbook as filename tName file format Excel XML file format with overwrite
			end tell
			
			close active workbook without saving
		end tell
	end if
end repeat
activate