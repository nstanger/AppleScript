tell application "Safari"
	set theURL to URL of current tab of window 1
end tell

-- Open in a new tab in the current window if possible, rather than always opening a new window.
tell application "Google Chrome"
	if (count of (every window where visible is true)) is greater than 0 then
		tell front window
			make new tab
		end tell
	else
		make new window
	end if
	set URL of active tab of front window to theURL
	--	activate
end tell