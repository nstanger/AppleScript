on run
	set {year:reportYear, month:reportMonth, day:reportDay} to (current date)
	
	tell script "Business Objects utilities"
		if not checkBusinessObjects() then return
		set theYear to getYear()
		set {periodCode, periodIndex} to getTeachingPeriod(missing value)
		if periodCode is missing value then return
		set theMode to getMode()
		
		-- Academic Year [user specified]
		-- This is a text field on this form.
		setEditField(0, theYear, "Academic Year:")
		
		-- Paper Codes [from file]
		pickKeyboard(1)
		loadMultipleListFromFile(1, "Enter the Paper Code(s). Wildcards can be used e.g. ACCT1*:", �
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
		setFileFormat(2)
		
		-- Report Destination [email]
		if not emailReport("nigel.stanger@otago.ac.nz", �
			"Business Objects: " & theYear & " " & periodCode & " " & theMode & " CA", �
			periodCode & "_" & reportYear & "-" & reportMonth * 1 & "-" & reportDay & "_courseapproval.%EXT%") then return
		
		if not waitForWindow("UO Business Objects: Report Scheduling") then return
		
		submitParameters()
	end tell
end run
