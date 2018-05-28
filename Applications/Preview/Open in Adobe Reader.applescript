tell application "Preview" to set p to path of document 1

tell application "Adobe Acrobat Reader DC"
	open p
	activate
end tell