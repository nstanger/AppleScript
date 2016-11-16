set tempFolder to POSIX path of (path to temporary items from user domain)

tell application "Microsoft Excel"
	
	set wkBook to active workbook
	set inputFolder to (path of wkBook) & "/"
	set inputName to (name of wkBook)
	
	tell application "System Events" to tell disk item (inputFolder & inputName) to set {fName, fExtension} to {name, name extension}
	
	if (fExtension is not "") then set fName to text 1 thru -((count fExtension) + 2) of fName -- the name part
	
	set tempFile to (tempFolder & fName & ".xlsx")
	
	save workbook as wkBook filename tempFile file format Excel XML file format with overwrite
	
	close wkBook without saving
	
end tell

do shell script "mv " & quoted form of tempFile & " " & quoted form of inputFolder
