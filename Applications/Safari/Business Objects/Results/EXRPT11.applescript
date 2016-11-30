on run
	tell script "Business Objects utilities"
		if not checkBusinessObjects() then return
		set {reportType, reportTypeIndex} to getResultsReportType()
		set paperCode to getPaperCode()
		set {periodCode, ignore} to getTeachingPeriod(missing value)
		if periodCode is missing value then return
		set theYear to getYear()
		
		-- Report Type [summary format = 1]
		-- The summary format, which you're instructed NOT to use with Excel, is actually better!
		selectRadio(0, 1)
		
		-- Data Type [SMR = 0]
		chooseMultipleList(1, 0)
		
		-- Academic Year [user specified]
		-- The current year should be selected already. We need to select the year as an offset from the first year in the select list.
		set firstYear to (runJavaScript("document.getElementsByName('paramEd_defaults2')[0].options[0].text")) as integer
		set reportYearIndex to theYear - firstYear
		chooseMultipleList(2, reportYearIndex)
		
		-- Campus [Dunedin]
		pickKeyboard(3)
		addTextToList(3, "DN")
		
		-- Group Campus? [no = 0]
		selectRadio(4, 0)
		
		-- Paper Code [user specified]
		pickKeyboard(5)
		addTextToList(5, paperCode)
		
		-- Teaching Period [user specified]
		pickKeyboard(6)
		addTextToList(6, periodCode)
		
		-- Assessment Number [800]
		pickKeyboard(7)
		addTextToList(7, 800)
		
		-- Marks to Include [actual (A) = 1 and agreed (G) = 2]
		chooseMultipleList(8, 1)
		chooseMultipleList(8, 2)
		
		-- Include Graphs? [None = 0]
		chooseList(9, 0)
		
		-- Include Names? [yes = 1]
		selectRadio(10, 1)
		
		-- Names Start With [Surname = 1]
		selectRadio(11, 1)
		
		-- Field to Sort By [SPR = 0]
		chooseList(12, 0)
		
		-- File Format [depends on report type]
		setFileFormat(reportTypeIndex)
		
		-- Report Destination [email]
		if not emailReport("nigel.stanger@otago.ac.nz", Â
			"Business Objects: " & paperCode & " " & periodCode & " " & theYear & " results " & reportType, Â
			paperCode & "_" & periodCode & "_" & theYear & "_results_" & reportType & ".%EXT%") then return
		
		if not waitForWindow("UO Business Objects: Report Scheduling") then return
		
		submitParameters()
	end tell
end run
