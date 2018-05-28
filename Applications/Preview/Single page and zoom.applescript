tell application "Preview" to set bounds of window 1 to {4, 27, 1147, 962}

--activate application "Preview"
--delay 0.5

tell application "System Events" to tell process "Preview"
	--	set keys to {"+", "+"}
	--	tell menu "View" of menu bar item "View" of menu bar 1
	--		if (value of attribute "AXMenuItemMarkChar" of menu item "Actual Size") is missing value then
	--			set keys to {"0"} & keys
	--		end if
	--		if (value of attribute "AXMenuItemMarkChar" of menu item "Single Page") is missing value then
	--			set keys to {"2"} & keys
	--		end if
	--	end tell
	--	repeat with k in keys
	--		delay 0.25
	--		keystroke k using command down
	--	end repeat
	tell menu "View" of menu bar item "View" of menu bar 1
		if (value of attribute "AXMenuItemMarkChar" of menu item "Single Page") is missing value then
			click menu item "Single Page"
		else
			click menu item "Zoom In"
			delay 0.25
			click menu item "Zoom to Fit"
		end if
	end tell
end tell
