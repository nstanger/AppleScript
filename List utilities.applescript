on listPosition(thisItem, thisList)
	repeat with i from 1 to the count of thisList
		if item i of thisList is thisItem then return i
	end repeat
	return 0
end listPosition