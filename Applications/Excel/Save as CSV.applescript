tell application "Microsoft Excel"
	set tempFolder to POSIX path of (path to temporary items from user domain)
	
	set wkBook to active workbook
	set inputFolder to (path of wkBook) & "/"
	set inputName to (name of wkBook)
	
	tell application "System Events" to tell disk item (inputFolder & inputName) to set {fName, fExtension} to {name, name extension}
	
	if (fExtension is not "") then set fName to text 1 thru -((count fExtension) + 2) of fName -- the name part
	
	set tempFile to tempFolder & fName & ".csv"
	
	save as active sheet filename tempFile file format CSV file format with overwrite
end tell

do shell script "/opt/local/bin/flip -u " & quoted form of tempFile
do shell script "mv " & quoted form of tempFile & " " & quoted form of inputFolder
