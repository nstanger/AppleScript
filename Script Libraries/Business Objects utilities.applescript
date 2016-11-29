-- Get the year from the user.
on getYear()
	tell application "Safari" to Â
		return text returned of (display dialog "Enter year:" default answer year of (current date))
end getYear

-- Get a teaching period code from the user.
on getTeachingPeriod(validPeriods)
	if validPeriods is missing value then
		set validPeriods to {"ALL", "SS", "S1", "S2", "FY", "N1", "N2"}
	end if
	set saveDelimiters to text item delimiters
	set text item delimiters to ", "
	set periodString to validPeriods as text
	set text item delimiters to saveDelimiters
	
	tell application "Safari"
		set periodCode to text returned of (display dialog "Enter teaching period (" & periodString & "):" default answer "")
		set periodCode to do shell script "echo " & periodCode & " | tr a-z A-Z"
		if validPeriods does not contain periodCode then
			display alert "Invalid period code Ò" & periodCode & "Ó entered" message "The period code must be one of " & periodString & "." as critical
			return
		end if
	end tell
	tell script "List utilities" to set periodIndex to (listPosition(periodCode, validPeriods) - 1)
	
	return {periodCode, periodIndex}
end getTeachingPeriod

-- Get a results data loading mode from the user.
on getMode()
	tell application "Safari" to Â
		return button returned of (display dialog "Choose list mode" buttons {"Initial", "Cutoff", "Final"})
end getMode

-- Execute a chunk of JavaScript in Safari.
on runJavaScript(jsCode)
	tell document 1 of application "Safari" to return (do JavaScript jsCode)
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
on setNamedElement(elementName, subIndex, elementValue, method)
	set jsCode to "document.getElementsByName('" & elementName & "')[" & subIndex & "].value = '" & elementValue & "';"
	set jsCode to jsCode & method & ";"
	runJavaScript(jsCode)
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
		"paramEd[" & fieldNum & "].LoadFromRadioDefault(" & index - 1 & ")")
end selectRadio

-- Editable text field.
on setEditField(fieldNum, fieldValue, fieldLabel)
	setNamedElement("editField" & fieldNum, 0, fieldValue, "paramEd[" & fieldNum & "].Changed('" & fieldLabel & "')")
end setEditField

-- Input parameters from keyboard.
on pickKeyboard(fieldNum)
	checkButton("PickKeyboard_" & fieldNum, 0)
	setNamedElement("PickKeyboard_" & fieldNum, 0, "on", "PickKeyboard(paramEd[" & fieldNum & "])")
end pickKeyboard

-- Input parameters from a file.
on pickFile(fieldNum, filePath)
	checkButton("PickKeyboard_" & fieldNum, 1)
	setNamedElement("PickKeyboard_" & fieldNum, 1, "on", "PickFile(paramEd[" & fieldNum & "])")
	setNamedElement("ReportParameterFile_" & fieldNum & "_", 0, filePath, "")
end pickFile

-- Single-option pick list.
on chooseList(fieldNum, fieldValue)
	setNamedElement("paramEd_defaults" & fieldNum, 0, fieldValue, "")
end chooseList

-- Multiple option pick list.
on chooseMultipleList(fieldNum, fieldValue)
	setNamedElement("paramEd_defaults" & fieldNum, 0, fieldValue, "paramEd[" & fieldNum & "].StoreDefault()")
end chooseMultipleList

-- Load values into a multiple option pick list from a file.
-- Needed because we can't set the value of a file picker input element from JavaScript :(.
on loadMultipleListFromFile(fieldNum, fieldLabel, filePath)
	set itemList to paragraphs of (read filePath)
	repeat with theItem in items 1 thru -2 of itemList -- skip empty line at end
		setEditField(fieldNum, theItem, fieldLabel)
		runJavaScript("paramEd[" & fieldNum & "].StoreDiscrete();")
	end repeat
end loadMultipleListFromFile

-- Wait for a page with the specified title to appear. Includes a 10 second timeout in case the page stalls.
on waitForWindow(windowTitle)
	set timer to 0
	repeat
		if runJavaScript("document.title") is windowTitle then return true
		if timer > 10 then
			tell application "Safari" to display alert "Timed out waiting for window Ò" & windowTitle & "Ó."
			return false
		end if
		set timer to timer + 0.5
		delay 0.5
	end repeat
end waitForWindow