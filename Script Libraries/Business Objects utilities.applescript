global applicationName

on init(appName)
	set applicationName to appName
end init

-- Get the year from the user.
on getYear()
	tell application applicationName to Â
		return text returned of (display dialog "Enter year:" default answer year of (current date))
end getYear

-- Get a teaching period code from the user.
on getTeachingPeriod(validPeriods)
	if validPeriods is missing value then
		set validPeriods to {"All", "SS", "S1", "S2", "FY", "N"}
	end if
	set saveDelimiters to text item delimiters
	set text item delimiters to ", "
	set periodString to validPeriods as text
	set text item delimiters to saveDelimiters
	
	tell application applicationName
		set periodCode to text returned of (display dialog "Enter teaching period (" & periodString & "):" default answer "")
		set periodCode to do shell script "echo " & periodCode & " | tr a-z A-Z"
		if validPeriods does not contain periodCode then
			display alert "Invalid period code Ò" & periodCode & "Ó entered" message "The period code must be one of " & periodString & "." as critical
			return {missing value, missing value}
		end if
	end tell
	tell script "List utilities" to set periodIndex to (listPosition(periodCode, validPeriods) - 1)
	
	return {periodCode, periodIndex}
end getTeachingPeriod

-- Get a results data loading mode from the user.
on getMode()
	tell application applicationName to Â
		return button returned of (display dialog "Choose list mode" buttons {"Initial", "Cutoff", "Final"})
end getMode

-- Get a paper code from the user.
on getPaperCode()
	tell application applicationName to Â
		set paperCode to text returned of (display dialog "Enter paper code:" default answer "INFOxxx")
	set paperCode to do shell script "echo " & paperCode & " | tr a-z A-Z"
	return paperCode
end getPaperCode

-- Get results report type (data or proof) from the user.
on getResultsReportType()
	tell application applicationName
		set reportType to button returned of (display dialog Â
			"Generate data (Excel), or proof (PDF)?" buttons {"Cancel", "Proof", "Data"} Â
			default button "Data" cancel button "Cancel")
	end tell
	if reportType is "Data" then
		set reportType to "data"
		set reportTypeIndex to 2
	else
		set reportType to "proof"
		set reportTypeIndex to 4
	end if
	return {reportType, reportTypeIndex}
end getResultsReportType

-- Execute a chunk of JavaScript in a browser.
-- Note:
-- Vivaldi requires Tools > Enable JavaScript from Apple Events.
-- Safari requires Develop > Allow JavaScript from Apple Events.
-- The latter will require turning off things like MagicPrefs!
on runJavaScript(jsCode)
	if (applicationName is "Safari") then
		tell application "Safari" to return (do JavaScript jsCode in front document)
	else if (applicationName is "Vivaldi") then
		tell application "Vivaldi" to return (execute active tab of window 1 javascript jsCode)
	else
		display alert "Invalid browser" message "Ò" & applicationName & "Ó is not a recognised browser." as critical
	end if
end runJavaScript

-- Get title of current document.
on getDocumentTitle()
	return runJavaScript("document.title")
end getDocumentTitle

-- Check that this is a Business Objects report parameters form.
on checkBusinessObjects()
	if getDocumentTitle() is not "UO Business Objects: Report Scheduling" then
		display alert "This doesnÕt appear to be a Business Objects report parameters form." as critical
		return false
	else
		return true
	end if
end checkBusinessObjects

-- Set value of a generic named element.
on setNamedElement(elementName, subIndex, elementValue, callback)
	set jsCode to "document.getElementsByName('" & elementName & "')[" & subIndex & "].value = '" & elementValue & "';"
	runJavaScript(jsCode)
	if callback is not "" then runJavaScript(callback & ";")
end setNamedElement

-- Enable check box or radio button.
on checkButton(fieldName, index)
	runJavaScript("document.getElementsByName('" & fieldName & "')[" & index & "].checked = true;")
end checkButton

-- Disable check box or radio button.
on uncheckButton(fieldName, index)
	runJavaScript("document.getElementsByName('" & fieldName & "')[" & index & "].checked = false;")
end uncheckButton

-- Select a radio button in a group.
on selectRadio(fieldNum, index)
	checkButton("RadioDefault_" & fieldNum, index)
	setNamedElement("RadioDefault_" & fieldNum, index, index - 1, Â
		"window.paramEd[" & fieldNum & "].LoadFromRadioDefault(" & index - 1 & ")")
end selectRadio

-- Editable text field.
-- Technically these have a .Changed() callback with an argument of the field label, but in practice
-- this seems to be unnecessary (and breaks completely if the label happens to contain an apostrophe).
on setTextField(fieldNum, fieldValue)
	setNamedElement("editField" & fieldNum, 0, fieldValue, "")
end setTextField

-- Input parameters from keyboard.
on pickKeyboard(fieldNum)
	checkButton("PickKeyboard_" & fieldNum, 0)
	setNamedElement("PickKeyboard_" & fieldNum, 0, "on", "window.PickKeyboard(paramEd[" & fieldNum & "])")
end pickKeyboard

-- Input parameters from a file.
on pickFile(fieldNum, filePath)
	checkButton("PickKeyboard_" & fieldNum, 1)
	setNamedElement("PickKeyboard_" & fieldNum, 1, "on", "window.PickFile(paramEd[" & fieldNum & "])")
	setNamedElement("ReportParameterFile_" & fieldNum & "_", 0, filePath, "")
end pickFile

-- Single-selection pick list.
on chooseList(fieldNum, fieldValue)
	setNamedElement("paramEd_defaults" & fieldNum, 0, fieldValue, "")
end chooseList

-- Multiple-selection pick list.
on chooseMultipleList(fieldNum, fieldValue)
	setNamedElement("paramEd_defaults" & fieldNum, 0, fieldValue, "window.paramEd[" & fieldNum & "].StoreDefault()")
end chooseMultipleList

-- Add a user-provided value to a list.
on addTextToList(fieldNum, theText)
	setTextField(fieldNum, theText)
	runJavaScript("window.paramEd[" & fieldNum & "].StoreDiscrete();")
end addTextToList

-- Load values into a multiple option pick list from a file.
-- Needed because we can't set the value of a file picker input element from JavaScript :(.
on loadMultipleListFromFile(fieldNum, filePath)
	set itemList to paragraphs of (read filePath)
	repeat with theItem in items 1 thru -2 of itemList -- skip empty line at end
		addTextToList(fieldNum, theItem)
	end repeat
end loadMultipleListFromFile

on setFileFormat(formatIndex)
	setNamedElement("ExportFormat", 0, formatIndex, "")
end setFileFormat

on emailReport(emailAddress, emailSubject, attachmentFilename)
	setNamedElement("Destination", 0, 1, "")
	runJavaScript("window.SubmitDestinations();")
	
	if not waitForWindow("UO Business Objects: Destination Options") then return false
	
	setNamedElement("RecipientAddresses", 0, emailAddress, "")
	setNamedElement("Subject", 0, emailSubject, "")
	setNamedElement("Attachment", 0, attachmentFilename, "")
	runJavaScript("window.SmtpFields();")
	
	return true
end emailReport

on submitParameters()
	runJavaScript("window.submitParamFrm();")
end submitParameters

-- Wait for a page with the specified title to appear. Includes a 10 second timeout in case the page stalls.
on waitForWindow(windowTitle)
	set timer to 0
	repeat
		if runJavaScript("document.title") is windowTitle then return true
		if timer > 10 then
			tell application applicationName to display alert "Timed out waiting for window Ò" & windowTitle & "Ó."
			return false
		end if
		set timer to timer + 0.5
		delay 0.5
	end repeat
end waitForWindow
