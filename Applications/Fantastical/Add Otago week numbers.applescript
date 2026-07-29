(*

--------------------------------------------
Add Week Numbers 1.0
©2008 William Smith
mecklists@comcast.net

This script may be freely modified for personal or commercial
purposes but may not be republished without prior consent.

If you find this script useful or have ideas for improving it, please let me know.

NJS 2026-07-29: Updated to use ISO 8601 week numbering and for 12 week semesters starting from 2027.
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
-- Calculate the first and last ISO 8601 weeks of the year
-- (see https://en.wikipedia.org/wiki/ISO_week_date).
-- First week contains 4 January, last week contains 28 December.
--
-- Ugh, AppleScript numbers days from 1 = Sunday to 7 = Saturday,
-- but ISO 8601 weeks start on Monday, so we need to rotate
-- everything one day to the right.
--
-- NJS 2026-07-29: Rewrote week calculation from scratch.
-----------------------------------------------------------

set newYearsDay to (date ("1 January" & theYear))
set lastNewYearsDay to date ("1 January " & (theYear - 1))
set daysLastYear to ((date ("31 December " & (theYear - 1))) - lastNewYearsDay + days) div days

set fourthJanuary to (date ("4 January" & theYear))
set fourthJanuaryDay to weekday of fourthJanuary
set twentyEighthDecember to (date ("28 December" & theYear))
set twentyEighthDecemberDay to weekday of twentyEighthDecember

set firstWeekStart to calculateIsoWeekStart(date ("4 January" & theYear))
set lastWeekStart to calculateIsoWeekStart(date ("28 December" & theYear))
set numWeeks to (lastWeekStart - firstWeekStart + weeks) div weeks

-- Give the user the opportunity to check and correct the calculated date.
set theResult to display dialog "I think week 1 starts on " & (date string of firstWeekStart) & ". Please modify below if this is incorrect." default answer (date string of firstWeekStart) with icon 1 with title "Confirm week 1 start date"

set firstWeekStart to date (text returned of theResult)

set theResult to display dialog "I think week " & numWeeks & " starts on " & (date string of lastWeekStart) & ". Please modify below if this is incorrect." default answer (date string of lastWeekStart) with icon 1 with title "Confirm week " & numWeeks & " start date"

set lastWeekStart to date (text returned of theResult)


-----------------------------------------------------------
-- Get details of teaching periods.
-----------------------------------------------------------

-- List of records specifying each teaching period. Set break properties to null if there is no
-- break for a particular teaching period. The values in here reflect the usual values for
-- Otago. Note that the first semester mid-semester break is variable due to the movement
-- of Easter, but the rest are very unlikely to change.
set thePeriods to {Â
	{id:"SS", name:"Summer School", periodStart:2, periodFinish:7, breakStart:null, breakFinish:null}, Â
	{id:"S1", name:"Semester 1", periodStart:9, periodFinish:22, breakStart:15, breakFinish:16}, Â
	{id:"S2", name:"Semester 2", periodStart:28, periodFinish:41, breakStart:35, breakFinish:36}, Â
	{id:"Pre-Xmas SS", name:"Pre-Christmas Summer School", periodStart:46, periodFinish:50, breakStart:null, breakFinish:null} Â
		}

-- Confirm period settings with the user and update as needed.
repeat with thisPeriod in thePeriods
	set periodStart of thisPeriod to getProperty(name of thisPeriod, "start", periodStart of thisPeriod, periodFinish of thisPeriod, firstWeekStart)
	set periodFinish of thisPeriod to getProperty(name of thisPeriod, "end", periodStart of thisPeriod, periodFinish of thisPeriod, firstWeekStart)
	if breakStart of thisPeriod is not null and breakFinish of thisPeriod is not null then
		set breakStart of thisPeriod to getProperty(name of thisPeriod & " Mid-semester Break", "start", breakStart of thisPeriod, breakFinish of thisPeriod, firstWeekStart)
		set breakFinish of thisPeriod to getProperty(name of thisPeriod & " Mid-semester Break", "end", breakStart of thisPeriod, breakFinish of thisPeriod, firstWeekStart)
	end if
end repeat


-----------------------------------------------------------
-- Populate the calendar with week numbers.
-----------------------------------------------------------

set dayNumber to ((firstWeekStart - lastNewYearsDay + days) div days) mod daysLastYear
set thisWeekStart to firstWeekStart

repeat with weekNumber from 1 to numWeeks
	
	set semesterString to ""
	
	repeat with thisPeriod in thePeriods
		set periodId to id of thisPeriod
		set periodStart to (periodStart of thisPeriod)
		set periodFinish to (periodFinish of thisPeriod)
		set breakStart to (breakStart of thisPeriod)
		set breakFinish to (breakFinish of thisPeriod)
		if (breakStart is not null and breakFinish is not null) then
			if ((weekNumber ³ periodStart) and (weekNumber < breakStart)) then
				set semesterString to periodId & " Week " & weekNumber - periodStart + 1
			else if ((weekNumber > breakFinish) and (weekNumber ² periodFinish)) then
				set semesterString to periodId & " Week " & weekNumber - periodStart - breakFinish + breakStart
			else if ((weekNumber ³ breakStart) and (weekNumber ² breakFinish)) then
				set semesterString to periodId & " Mid-semester Break Week " & weekNumber - breakStart + 1
			end if
		else
			if ((weekNumber ³ periodStart) and (weekNumber ² periodFinish)) then
				set semesterString to periodId & " Week " & weekNumber - periodStart + 1
			end if
		end if
	end repeat
	
	set {calDay, calMonth, calYear} to {day, month, year} of thisWeekStart
	
	-- debugging
	-- log "Week " & weekNumber & " - Day " & dayNumber & "' on " & calMonth & " " & calDay & " " & calYear
	-- if (semesterString ­ "") then
	--	log semesterString & "' on " & calMonth & " " & calDay & " " & calYear
	-- end if
	
	tell application "Fantastical"
		parse sentence "'Week " & weekNumber & " - Day " & dayNumber & "' on " & calMonth & " " & calDay & " " & calYear calendarName "Week numbers" with add immediately
		if (semesterString ­ "") then
			parse sentence "'" & semesterString & "' on " & calMonth & " " & calDay & " " & calYear calendarName "Week numbers" with add immediately
		end if
	end tell
	
	set dayNumber to dayNumber + 7
	if (dayNumber > daysLastYear) then
		set dayNumber to dayNumber mod daysLastYear
	end if
	
	set thisWeekStart to thisWeekStart + weeks
	
end repeat


(*
For a given date, return the Sunday before the Monday (ISO 8601 week day 1) for that week.
We want the Sunday because that's a better place to insert the calendar entry. Annoyingly,
AppleScript numbers week days from Sunday = 1 to Saturday = 7, but at least this isn't
changed by the "first day of week" system setting.

Arguments:
	theDate	a valid date
*)
on calculateIsoWeekStart(theDate)
	set theWeekDay to (weekday of theDate) as integer
	-- shift sequence so Monday (2) maps to 1, and Sunday (1) maps to 7
	set isoWeekDay to ((theWeekDay + 5) mod 7) + 1
	return theDate - (isoWeekDay * days)
end calculateIsoWeekStart


(*
Get the value of a teaching period property from the user. The function does not return
until a valid week number is entered.

Arguments:
	periodName		user-visible name of the teaching period
	periodBound		one of "start" or "end"
	minWeek			the minimum possible week number for this teaching period
	maxWeek			the maximum possible week number for this teaching period
	firstWeekStart	date of Sunday of the first week of the year

NJS 2013-03-05:
	¥ The default answer for the dialog box now depends on the period bound.
	¥ For convenience (as University documents normally list the dates, not the
	  week numbers), the start date of the proposed week is now included in the
	  dialog box text (and is stripped out again afterwards, if necessary.)
	¥ Added bounds checks to ensure the resulting week number is in the range 1Ð52
	  and that then end week of a period is not earlier than the start week.
	¥ Switched to alerts for errors.
*)
on getProperty(periodName, periodBound, minWeek, maxWeek, firstWeekStart)
	set weekValid to false
	repeat until weekValid
		if (periodBound = "start") then
			set defaultAnswer to minWeek
		else
			set defaultAnswer to maxWeek
		end if
		
		-- What date does the target week start? This is included in the dialog box text below.
		set weekStart to firstWeekStart + ((defaultAnswer - 1) * weeks)
		
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
