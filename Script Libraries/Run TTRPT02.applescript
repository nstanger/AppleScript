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
		set {periodCode, periodIndex} to getTeachingPeriod(missing value)
		if periodCode is missing value then return
		set theMode to getMode()
		
		-- Paper Codes [from file]
		pickKeyboard(0)
		loadMultipleListFromFile(0, POSIX path of (path to documents folder as Unicode text) & "Teaching/Results/papers" & periodCode & ".txt")
		
		-- Teaching Period [user specified]
		chooseMultipleList(1, periodIndex)
		
		-- Course Approval Status [All = 0, Student declared = 3]
		if periodCode is "SS" then
			chooseMultipleList(2, 0)
		else
			chooseMultipleList(2, 3)
		end if
		
		-- Academic Year [user specified]
		-- This is a text field on this form.
		setTextField(3, theYear)
		
		-- Campus [Dunedin = 3]
		chooseMultipleList(4, 3)
		
		-- Display streams? [No = 1]
		selectRadio(5, 1)
		
		-- Display demographic data? [Yes = 0]
		selectRadio(6, 0)
		
		-- Display contact information? [Yes = 0]
		selectRadio(7, 0)
		
		-- Display programme information? [Yes = 0]
		selectRadio(8, 0)
		
		-- Report format [Excel = 0]
		selectRadio(9, 0)
		
		-- Attendance Mode [Both = 0]
		selectRadio(10, 0)
		
		-- File Format [Excel = 2]
		setFileFormat(2)
		
		-- Report Destination [email]
		if not emailReport("nigel.stanger@otago.ac.nz", Â
			"Business Objects: " & theYear & " " & periodCode & " " & theMode & " lists CL", Â
			periodCode & "_" & reportYear & "-" & reportMonth & "-" & reportDay & "_classlist.%EXT%") then return
		
		if not waitForWindow("UO Business Objects: Report Scheduling") then return
		
		submitParameters()
	end tell
end run
