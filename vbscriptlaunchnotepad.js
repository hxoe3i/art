Set WshShell = WScript.CreateObject("WScript.Shell")
Dim exeName Dim statusCode
exeName = "%windir%\notepad"
statusCode = WshShell.Run (exeName, 1, true)
MsgBox("End of Program")
