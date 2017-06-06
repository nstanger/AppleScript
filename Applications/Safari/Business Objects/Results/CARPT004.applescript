on run
	set {year:reportYear, month:reportMonth, day:reportDay} to (current date)
	set reportMonth to reportMonth * 1
	if reportMonth < 10 then
		set reportMonth to "0" & reportMonth
	else
		set reportMonth to reportMonth as text
	end if
	if reportDay < 10 then
		set reportDay to "0" & reportDay
	else
		set reportDay to reportDay as text
	end if
	
	tell script "Business Objects utilities"
		if not checkBusinessObjects() then return
		set theYear to getYear()
		set {periodCode, periodIndex} to getTeachingPeriod({"All", "FY", "S1", "S2", "SS", "N"})
		if periodCode is missing value then return
		set theMode to getMode()
		
		-- Academic Year [user specified]
		setTextField(0, theYear)
		
		-- Paper Codes [from file]
		pickKeyboard(1)
		loadMultipleListFromFile(1, "/Users/nstanger/Documents/Teaching/Results/papers" & periodCode & ".txt")
		
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
		if not emailReport("nigel.stanger@otago.ac.nz", Â
			"Business Objects: " & theYear & " " & periodCode & " " & theMode & " lists CA", Â
			periodCode & "_" & reportYear & "-" & reportMonth & "-" & reportDay & "_courseapproval.%EXT%") then return
		
		if not waitForWindow("UO Business Objects: Report Scheduling") then return
		
		submitParameters()
	end tell
end run
