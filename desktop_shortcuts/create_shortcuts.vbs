set ws  = WScript.CreateObject("WScript.Shell")
set arg = Wscript.Arguments
linkFile = arg(0)
set link = ws.CreateShortcut(linkFile)
link.TargetPath = arg(1)
link.Save