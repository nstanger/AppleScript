on run
	set {year:reportYear, month:reportMonth, day:reportDay} to (current date)
	
	tell script "Business Objects utilities"
		if not checkBusinessObjects() then return
		set theYear to getYear()
		set {periodCode, periodIndex} to getTeachingPeriod(missing value)
		set theMode to getMode()
		
		-- Academic Year [user specified]
		-- This is a text field on this form.
		setEditField(0, theYear, "Academic Year:")
		
		-- Paper Codes [from file]
		pickKeyboard(1)
		loadMultipleListFromFile(1, "Enter the Paper Code(s). Wildcards can be used e.g. ACCT1*:", Â
			"/Users/nstanger/Documents/Teaching/Results/papers" & periodCode & ".txt")
		
		-- Teaching Period [user specified]
		chooseMultipleList(2, periodIndex)
		
		-- Course Approval Status [Student declared (CA40) = 3]
		chooseMultipleList(3, 3)
		
		-- Attendance Mode [Both = 3]
		selectRadio(4, 3)
		
		-- Location [Dunedin = 3]
		chooseMultipleList(5, 3)
		
		-- File Format [Excel = 2]
		setNamedElement("ExportFormat", 0, 2, "")
		
		-- Report Destination [email]
		setNamedElement("Destination", 0, 1, "")
		runJavaScript("SubmitDestinations();")
		
		if not waitForWindow("UO Business Objects: Destination Options") then return
		
		setNamedElement("RecipientAddresses", 0, "nigel.stanger@otago.ac.nz", "")
		setNamedElement("Subject", 0, "Business Objects: " & theYear & " " & periodCode & " " & theMode & " CA", "")
		setNamedElement("Attachment", 0, periodCode & "_" & reportYear & "-" & reportMonth * 1 & "-" & reportDay & "_courseapproval.%EXT%", "")
		runJavaScript("SmtpFields();")
		
		if not waitForWindow("UO Business Objects: Report Scheduling") then return
		
		--runJavaScript("submitParamFrm();")
	end tell
end run
