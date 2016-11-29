tell application "Microsoft Word"
	set tempFolder to POSIX path of (path to temporary items from user domain)
	
	set doc to active document
	set inputFolder to (path of doc) & ":"
	set inputName to (name of doc)
	alias inputFolder
	
	tell application "System Events" to tell disk item (inputFolder & inputName) to set {fName, fExtension} to {name, name extension}
	
	if (fExtension is not "") then set fName to text 1 thru -((count fExtension) + 2) of fName -- the name part
	
	set tempFile to (tempFolder & fName & ".docx")
	
	save as doc file name tempFile file format format document with overwrite
end tell

do shell script "mv " & quoted form of tempFile & " " & quoted form of POSIX path of inputFolder
