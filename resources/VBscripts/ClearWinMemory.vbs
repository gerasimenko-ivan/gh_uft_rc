'*******************************************************************
' ���: CleanStandbyMemory.vbs
' ����: VBScript
' ��������: ������ ������ � ������� ��������
'*******************************************************************
On Error Resume Next

Dim WshShell
' ������� ������ WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

' ������������� �� ����. ���������. � ���������� ����� � ������
' ������� ������� ��� ������������.  ������ �������� �� (������ ��� qtp2)
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


'���� ������� ���������,� ���� ������� ���������� ����������
If isRammapRun() Then 
		WshShell.Run("taskkill /IM RAMMap64.exe")
		'WshShell.AppActivate("RamMap - Sysinternals: www.sysinternals.com")
		'WshShell.SendKeys "{VK_LWIN}"
	Else 
		'WScript.Echo "�� ��������"
End if

If Err.Number <> 0 Then
	WScript.Echo "���� ������", Err.Description, Err.Source
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