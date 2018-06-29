'*******************************************************************
' »м€: CleanStandbyMemory.vbs
' язык: VBScript
' ќписание: „истим пам€ть в разделе ќжидание
'*******************************************************************
On Error Resume Next

Dim WshShell
' Создаем объект WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

' Переключаемся на англ. раскладку. В настройках винды я создал
' горячие клавиши для переключения.  просто нажимаем их (только для qtp2)
WshShell.SendKeys "^9"
WScript.Sleep 1000
WshShell.Run "C:\RAMMap\RAMMap.exe"
WScript.Sleep 5000
WshShell.AppActivate "RamMap - Sysinternals: www.sysinternals.com"
WshShell.SendKeys "%{E}{DOWN 3}"
WScript.Sleep 3000
WshShell.SendKeys "{ENTER}"
WScript.Sleep 3000
WshShell.SendKeys "%{F4}"


'Блок требует доработки,а пока убивает запущенное приложение
If isRammapRun() Then 
		WshShell.Run("taskkill /IM RAMMap64.exe")
		'WshShell.AppActivate("RamMap - Sysinternals: www.sysinternals.com")
		'WshShell.SendKeys "{VK_LWIN}"
	Else 
		'WScript.Echo "не запущено"
End if

If Err.Number <> 0 Then
	WScript.Echo "Была ошибка", Err.Description, Err.Source
	Err.Clear
End if

WshShell.close

WScript.Quit

'*******************************************************************
Function isRammapRun()
 	Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
	Set colProcessList = objWMIService.ExecQuery _
	("SELECT * FROM Win32_Process WHERE Name = 'RAMMap64.exe' ")
	If colProcessList.Count = 0 Then 
		isRammapRun = False
		Else 
			isRammapRun = True
	End if	
End Function


 ' For Each objProcess in colProcessList
	' WScript.Echo objProcess.Caption, objProcess.CommandLine, objProcess.ProcessId
    ' objProcess.Terminate
 ' Next