Set WshShell=WScript.CreateObject("WScript.Shell")
WshShell.AppActivate"ºÙºÙ"
for i=1 to 20
WScript.Sleep 10
WshShell.SendKeys"^v"
WshShell.SendKeys"%s"
Next