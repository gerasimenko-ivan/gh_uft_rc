' dummy executor
' should return QTP ExitCode
' 20 seconds for execution

Dim oShell : Set oShell = CreateObject("WScript.Shell")
execState = oShell.Run("C:\_qtp\resources\dummytimer.bat", , True)


Wscript.Echo "kill timer"
oShell.Run "taskkill /fi ""WINDOWTITLE eq TestExecuteTimer*"" /f /t", , True
Wscript.Echo "timer is killed"

Wscript.Echo "execState " & execState

WScript.Quit execState