set inputFolder to choose folder with prompt "Choose the folder that contains your .xls files"
tell application "Microsoft Excel" to set tempFolder to POSIX path of (path to temporary items from user domain)

tell application "Finder" to set theFiles to (files of inputFolder whose name extension is "xls")

repeat with inputFile in theFiles
	set fName to text 1 thru -5 of ((name of inputFile) as text)
	set tempFile to tempFolder & fName & ".xlsx"
	tell application "Microsoft Excel"
		open inputFile as alias
		
		save as active sheet filename tempFile file format workbook normal file format with overwrite
		
		close active workbook without saving
	end tell
	
	do shell script "mv " & quoted form of tempFile & " " & quoted form of POSIX path of inputFolder
end repeat

