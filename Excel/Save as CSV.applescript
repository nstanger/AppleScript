tell application "Microsoft Excel"
	
	set wFolder to (path of active workbook) & ":"
	set wName to (name of active workbook)
	
	tell application "System Events" to tell disk item (wFolder & wName) to set {fName, fExtension} to {name, name extension}
	
	if (fExtension is not "") then set fName to text 1 thru -((count fExtension) + 2) of fName -- the name part
	
	set tName to wFolder & fName & ".csv"
	
	tell active workbook
		save workbook as filename tName file format CSV Mac file format with overwrite
	end tell
	
end tell

tell application "System Events"
	do shell script "/opt/local/bin/flip -u '" & (POSIX path of disk item tName) & "'"
end tell