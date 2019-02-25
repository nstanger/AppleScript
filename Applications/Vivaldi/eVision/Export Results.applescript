-- Execute a chunk of JavaScript in Vivaldi.
on runJavaScript(jsCode)
	tell document 1 of application "Vivaldi" to return (execute javascript jsCode)
end runJavaScript

-- First Export Marks button
runJavaScript("document.getElementsByName('BP113.DUMMY_B.MENSYS.1')[0].click();")
delay 1

-- Second Export Marks button
runJavaScript("document.getElementsByName('BP102.DUMMY_B.MENSYS.1')[0].click();")
delay 4

-- View File link
-- Programmatically click a link.
-- https://stackoverflow.com/questions/902713/how-do-i-programmatically-click-a-link-with-javascript
runJavaScript("var links = document.getElementsByTagName('a'); var link = null; for (var i = 0; i < links.length; i++) { if (links[i].text == 'View File') { link = links[i]; } } var cancelled = false; if (document.createEvent) { var event = document.createEvent('MouseEvents'); event.initMouseEvent('click', true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null); cancelled = !link.dispatchEvent(event); } else if (link.fireEvent) { cancelled = !link.fireEvent('onclick'); } if (!cancelled) { window.location = link.href; }")
delay 3

-- First Back button
runJavaScript("document.getElementsByName('BP101.DUMMY_B.MENSYS.1')[0].click();")
delay 1

-- Second Back button
runJavaScript("document.getElementsByName('BP101.DUMMY_B.MENSYS.1')[0].click();")
delay 1

