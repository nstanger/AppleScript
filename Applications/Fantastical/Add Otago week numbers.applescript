(*

--------------------------------------------
Add Week Numbers 1.0
©2008 William Smith
mecklists@comcast.net

This script may be freely modified for personal or commercial
purposes but may not be republished without prior consent.

If you find this script useful or have ideas for improving it, please let me know.

NJS 2021-11-16: Updated for Fantastical 3.
NJS 2018-05-28: Rewritten to work with Fantastical 2.
NJS 2012-11-22: Otago now apparently works on "week containing 1 Jan", so the code has been reverted back to its original form.
NJS 2010-02-17: Modified to work with iCal.
NJS 2009-03-27: Modified to work with Otago's week numbering scheme (starting first Sunday *after* 1 Jan).
--------------------------------------------

This script adds a new all-day calendar event on every Sunday for the specified year.
Each event is named for the number of the week and the ordinal date for the current year.

Installation: Place this script in your user home folder in this location
~/Library/Scripts/Applications/iCal.

Use: Select "Add Week Numbers" from the Script menu in iCal and enter a year
from 1905 to 2039. All-day events will be added to each Sunday of the year
similar to the following information:

	Week 28 - Day 188

You will also be prompted for the week numbers for the start and end of various teaching
periods (e.g., Semester 1 start/end, mid-semester break, etc.). Additional events will be
added during these periods with additional information on the teaching period week number,
e.g.:

	S1 Week 3
*)


-----------------------------------------------------------
-- Ask for a year - must between 1905 and 2039
-----------------------------------------------------------

set theYear to ""

repeat until theYear ³ 1905 and theYear ² 2039
	set theYear to display dialog "This script will add week numbers to the ÒWeek numbersÓ calendar for the year below." default answer year of (current date) with icon 1 with title "Add week numbers for this year"
	
	set theYear to text returned of theYear
	
	if theYear ² 1905 or theYear ³ 2039 then
		display dialog "Enter a year from 1905 to 2039." with icon 2 with title "Alert!"
	end if
end repeat

--end tell


-----------------------------------------------------------
-- What's the first day of the year?
-- Now, what's the first Sunday of the year?
-- This script will calculate week numbers
-- based on the first Sunday of the year.
--
-- NJS 2013-03-05: Moved this block before getting the teaching period details.
-- NJS 2012-11-22: Switched back to the first Sunday being the Sunday *before* 1 Jan.
-- NJS 2009-03-27: The first Sunday should be the Sunday *following* 1 Jan, so we need to flip the sense of the original code.
-----------------------------------------------------------

set newYearsDay to (date ("1 January" & theYear))
set lastNewYearsDay to date ("1 January " & (theYear - 1))
set daysLastYear to (((date ("31 December " & (theYear - 1))) - lastNewYearsDay) div days) + 1

set day1 to weekday of newYearsDay

if day1 is Sunday then
	set theSunday to 0
else if day1 is Monday then
	set theSunday to 1
else if day1 is Tuesday then
	set theSunday to 2
else if day1 is Wednesday then
	set theSunday to 3
else if day1 is Thursday then
	set theSunday to 4
else if day1 is Friday then
	set theSunday to 5
else if day1 is Saturday then
	set theSunday to 6
end if

set firstSunday to newYearsDay - (theSunday * days)

-- NJS 2012-08-02: Give the user the opportunity to check and correct the calculated date.
set theResult to display dialog "I think week 1 starts on " & (date string of firstSunday) & ". Please correct below if this is incorrect." default answer (date string of firstSunday) with icon 1 with title "Confirm week 1 start date"

set firstSunday to date (text returned of theResult)


-----------------------------------------------------------
-- Get details of teaching periods.
-----------------------------------------------------------

-- List of records specifying each teaching period. Set break property to {} if there is no
-- break for a particular teaching period. The values in here reflect the usual values for
-- Otago. Note that the first semester mid-semester break is variable due to the movement
-- of Easter, but the rest are very unlikely to change.
set thePeriods to {Â
	{id:"SS", name:"Summer School", period:{start:2, finish:7, break:{}}}, Â
	{id:"S1", name:"Semester 1", period:{start:9, finish:22, break:{start:15, finish:15}}}, Â
	{id:"S2", name:"Semester 2", period:{start:29, finish:42, break:{start:36, finish:36}}}, Â
	{id:"Pre-Xmas SS", name:"Pre-Christmas Summer School", period:{start:46, finish:50, break:{}}} Â
		}

set minWeek to 1
set maxWeek to 52

repeat with thisPeriod in thePeriods
	-- Create references to the various properties for ease of access.
	-- Note that we can't use these for setting the value, only reading.
	set theStart to (a reference to start of period of thisPeriod)
	set theEnd to (a reference to finish of period of thisPeriod)
	set theBreak to (a reference to break of period of thisPeriod)
	if (theBreak is not {}) then
		set theBreakStart to (a reference to start of break of period of thisPeriod)
		set theBreakEnd to (a reference to finish of break of period of thisPeriod)
	end if
	
	-- Remember, we can't use the reference for setting the value.
	set start of period of thisPeriod to getProperty(name of thisPeriod, "start", theStart, theEnd, firstSunday)
	set finish of period of thisPeriod to getProperty(name of thisPeriod, "end", theStart, theEnd, firstSunday)
	if (contents of theBreak is not {}) then
		set start of break of period of thisPeriod to getProperty(name of thisPeriod & " Mid-semester Break", "start", theBreakStart, theBreakEnd, firstSunday)
		set finish of break of period of thisPeriod to getProperty(name of thisPeriod & " Mid-semester Break", "end", theBreakStart, theBreakEnd, firstSunday)
	end if
end repeat


-----------------------------------------------------------
-- Go populate the default calendar
-- with week numbers.
-----------------------------------------------------------

set weekNumber to 1
set dayNumber to ((firstSunday - lastNewYearsDay) div days + 1) mod daysLastYear

repeat with i from 0 to 51
	
	set semesterString to ""
	
	repeat with thisPeriod in thePeriods
		if (break of period of thisPeriod is not {}) then
			if ((weekNumber ³ start of period of thisPeriod) and (weekNumber < start of break of period of thisPeriod)) then
				set semesterString to id of thisPeriod & " Week " & weekNumber - (start of period of thisPeriod) + 1
			else if ((weekNumber > finish of break of period of thisPeriod) and (weekNumber ² finish of period of thisPeriod)) then
				set semesterString to id of thisPeriod & " Week " & weekNumber - (start of period of thisPeriod)
			else if ((weekNumber ³ start of break of period of thisPeriod) and (weekNumber ² finish of break of period of thisPeriod)) then
				set semesterString to id of thisPeriod & " Mid-semester Break"
			end if
		else
			if ((weekNumber ³ start of period of thisPeriod) and (weekNumber ² finish of period of thisPeriod)) then
				set semesterString to id of thisPeriod & " Week " & weekNumber - (start of period of thisPeriod) + 1
			end if
		end if
	end repeat
	
	set nextSunday to (firstSunday) + (weeks * i)
	set {nsDay, nsMonth, nsYear} to {day, month, year} of nextSunday
	
	tell application "Fantastical"
		parse sentence "'Week " & weekNumber & " - Day " & dayNumber & "' on " & nsMonth & " " & nsDay & " " & nsYear calendarName "Week numbers" with add immediately
		if (semesterString ­ "") then
			parse sentence "'" & semesterString & "' on " & nsMonth & " " & nsDay & " " & nsYear calendarName "Week numbers" with add immediately
		end if
	end tell
	
	set weekNumber to weekNumber + 1
	set dayNumber to dayNumber + 7
	if (dayNumber > daysLastYear) then
		set dayNumber to dayNumber mod daysLastYear
	end if
	
end repeat


(*
Get the value of a teaching period property from the user. The function does not return
until a valid week number is entered.

Arguments:
	periodName	user-visible name of the teaching period
	periodBound	one of "start" or "end"
	minWeek		the minimum possible week number for this teaching period
	maxWeek		the maximum possible week number for this teaching period
	firstSunday	date of first Sunday of the year

NJS 2013-03-05:
	¥ The default answer for the dialog box now depends on the period bound.
	¥ For convenience (as University documents normally list the dates, not the
	  week numbers), the start date of the proposed week is now included in the
	  dialog box text (and is stripped out again afterwards, if necessary.)
	¥ Added bounds checks to ensure the resulting week number is in the range 1Ð52
	  and that then end week of a period is not earlier than the start week.
	¥ Switched to alerts for errors.
*)
on getProperty(periodName, periodBound, minWeek, maxWeek, firstSunday)
	set weekValid to false
	repeat until weekValid
		if (periodBound = "start") then
			set defaultAnswer to minWeek
		else
			set defaultAnswer to maxWeek
		end if
		
		-- What date does the target week start? This is included in the dialog box text below.
		set weekStart to firstSunday + ((defaultAnswer - 1) * weeks)
		
		set theResult to display dialog "In which week number does " & periodName & " " & periodBound & " (usually weeks " & minWeek & "Ð" & maxWeek & ")?" default answer (defaultAnswer as text) & " (starting " & (day of weekStart as text) & " " & (month of weekStart as text) & " " & (year of weekStart as text) & ")" with icon 1 with title "Set teaching period properties"
		
		-- Extract just the entered week number. There may or may not be a date string to remove.
		set od to AppleScript's text item delimiters
		set AppleScript's text item delimiters to {" "}
		set resultWeek to (text item 1 of (text returned of theResult)) as integer
		set AppleScript's text item delimiters to od
		
		-- Bounds sanity checks.
		if ((resultWeek ³ 1) and (resultWeek ² 52)) then
			
			if ((periodBound = "end") and (resultWeek < minWeek)) then
				display alert "End of " & periodName & " must be no earlier than week " & minWeek & "." as warning
			else
				set weekValid to true
			end if
			
		else
			display alert "The week number must be in the range 1Ð52." as warning
		end if
		
	end repeat
	return resultWeek
end getProperty
