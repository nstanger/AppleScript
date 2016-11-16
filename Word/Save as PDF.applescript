tell application "Microsoft Word"
	
	set wFolder to (path of active document)
	set oldDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to {":"}
	set pathItems to text items of wFolder
	set wName to last item of pathItems
	set wFolder to ((reverse of the rest of reverse of pathItems) as string) & ":"
	set AppleScript's text item delimiters to oldDelimiters
	
	tell application "System Events" to tell disk item (wFolder & wName) to set {fName, fExtension} to {name, name extension}
	
	if (fExtension is not "") then set fName to text 1 thru -((count fExtension) + 2) of fName -- the name part
	
	set tName to wFolder & fName & ".pdf"
	
	tell active document
		save as file name tName file format format PDF with overwrite
	end tell
	
end tell
