on run
	tell document 1 of application "Safari"
		if name is not "UO Business Objects: Report Scheduling" then
			display alert "This doesnÕt appear to be a Business Objects report parameters form." as critical
			return
		end if
	end tell
	
	tell application "Safari"
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
	
	set validPeriods to {"ALL", "SS", "S1", "S2", "FY", "N1", "N2"}
	set saveDelimiters to text item delimiters
	set text item delimiters to ", "
	set periodString to validPeriods as text
	set text item delimiters to saveDelimiters
	
	tell application "Safari"
		set paperCode to text returned of (display dialog "Enter paper code:" default answer "INFOxxx")
		
		set periodCode to text returned of (display dialog "Enter teaching period (" & periodString & "):" default answer "")
		set periodCode to do shell script "echo " & periodCode & " | tr a-z A-Z"
		if periodCode is not in validPeriods then
			display alert "Invalid period code Ò" & periodCode & "Ó entered" message "The period code must be one of " & periodString & "." as critical
			return
		end if
		
		set theYear to (text returned of (display dialog "Enter year:" default answer year of (current date))) as integer
	end tell
	
	-- Report Type [user specified]
	-- The summary format, which you're instructed NOT to use with Excel, is actually better!
	checkOption("RadioDefault_0", 0, false)
	checkOption("RadioDefault_0", 1, true)
	
	-- Data Type [SMR]
	setNamedElement("paramEd_defaults1", 0, 0, 1, "Default")
	
	-- Academic Year [user specified]
	-- The current year should be selected already. We need to select the year as an offset from the first year in the select list.
	set firstYear to (runJavaScript("document.getElementsByName('paramEd_defaults2')[0].options[0].text")) as integer
	set reportYearIndex to theYear - firstYear
	setNamedElement("paramEd_defaults2", 0, reportYearIndex, 2, "Default")
	
	-- Campus [Dunedin]
	checkOption("RadioDefault_3", 0, true)
	checkOption("RadioDefault_3", 1, false)
	setNamedElement("editField3", 0, "DN", 3, "Discrete")
	
	-- Group Campus? [no]
	checkOption("RadioDefault_4", 0, true)
	checkOption("RadioDefault_4", 1, false)
	
	-- Paper Code [user specified]
	checkOption("RadioDefault_5", 0, true)
	checkOption("RadioDefault_5", 1, false)
	setNamedElement("editField5", 0, paperCode, 5, "Discrete")
	
	-- Teaching Period [user specified]
	checkOption("RadioDefault_6", 0, true)
	checkOption("RadioDefault_6", 1, false)
	setNamedElement("editField6", 0, periodCode, 6, "Discrete")
	
	-- Assessment Number [800]
	checkOption("RadioDefault_7", 0, true)
	checkOption("RadioDefault_7", 1, false)
	setNamedElement("editField7", 0, 800, 7, "Discrete")
	
	-- Marks to Include [actual (A) and agreed (G)]
	setNamedElement("paramEd_defaults8", 0, 1, 8, "Default")
	setNamedElement("paramEd_defaults8", 0, 2, 8, "Default")
	
	-- Include Graphs? [no]
	setNamedElement("paramEd_defaults9", 0, 0, 9, "Default")
	
	-- Include Names? [yes]
	checkOption("RadioDefault_10", 0, false)
	checkOption("RadioDefault_10", 1, true)
	
	-- Names Start With [Surname]
	checkOption("RadioDefault_11", 0, false)
	checkOption("RadioDefault_11", 1, true)
	
	-- Field to Sort By [SPR]
	setNamedElement("paramEd_defaults12", 0, 0, 12, "Default")
	
	-- File Format [depends on report type]
	setNamedElement("ExportFormat", 0, reportTypeIndex, missing value, missing value)
	
	-- Report Destination [email]
	setNamedElement("Destination", 0, 1, missing value, missing value)
	runJavaScript("SubmitDestinations();")
	
	-- We should now be in the Destination Options popup. Loop until it appears.
	repeat
		if runJavaScript("document.title") is "UO Business Objects: Destination Options" then exit repeat
	end repeat
	
	setNamedElement("RecipientAddresses", 0, "nigel.stanger@otago.ac.nz", missing value, missing value)
	setNamedElement("Subject", 0, "Business Objects: " & paperCode & " " & periodCode & " " & theYear & " " & reportType, missing value, missing value)
	setNamedElement("Attachment", 0, paperCode & "_" & periodCode & "_" & theYear & "_" & reportType & ".%EXT%", missing value, missing value)
	runJavaScript("SmtpFields();")
	
	-- Back to the main report parameters form.
	repeat
		if runJavaScript("document.title") is "UO Business Objects: Report Scheduling" then exit repeat
	end repeat
	runJavaScript("submitParamFrm();")
end run

-- Execute a chunk of JavaScript in Safari.
on runJavaScript(jsCode)
	tell document 1 of application "Safari" to return (do JavaScript jsCode)
end runJavaScript

-- Fiddle with checkboxes and radio buttons.
on checkOption(optionName, index, value)
	runJavaScript("document.getElementsByName('" & optionName & "')[" & index & "].checked = " & value & ";")
end checkOption

-- Fiddle with other elements, indexed by name.
on setNamedElement(elementName, elementIndex, elementValue, paramIndex, storeType)
	if storeType is missing value then set storeType to "Default"
	set jsCode to "document.getElementsByName('" & elementName & "')[" & elementIndex & "].value = '" & elementValue & "';"
	if paramIndex is not missing value then
		set jsCode to jsCode & "paramEd[" & paramIndex & "].Store" & storeType & "();"
	end if
	runJavaScript(jsCode)
end setNamedElement
