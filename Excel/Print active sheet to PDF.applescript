tell application "Microsoft Excel"
	set workbookName to (name of active workbook)
	set workbookPath to (full name of active workbook)
	save as active sheet filename "temp.pdf" file format PDF file format with overwrite
end tell

-- Lop off the workbook name.
set workbookPath to text 1 thru -((length of workbookName) + 1) of workbookPath
-- Lop off the extension.
set workbookName to text 1 thru -6 of workbookName

tell application "Finder"
	try
		delete file (workbookPath & workbookName & ".pdf")
	end try
	set name of file (workbookPath & "temp Sheet1.pdf") to (workbookName & ".pdf")
end tell
