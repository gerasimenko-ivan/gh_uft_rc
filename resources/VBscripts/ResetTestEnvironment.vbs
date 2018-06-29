

' Self error handling
On Error Resume Next

'=========================================================================================================================================
' variables
'=========================================================================================================================================
Dim oShell : Set oShell = CreateObject("WScript.Shell")

Dim exitCode
exitCode = 0


'=========================================================================================================================================
' Logging
'=========================================================================================================================================
' Log dirs. Create (if not exist) directories for logs
logDir = "C:\_qtp\runresults\logs"
CreateDir(logDir)

caoLogDir = logDir & "\CAO-Test"
CreateDir(caoLogDir)

zaoLogDir = logDir & "\ZAO-Test"
CreateDir(zaoLogDir)

uzaoLogDir = logDir & "\UZAO-Test"
CreateDir(uzaoLogDir)

caoTcodLogDir = logDir & "\CAO-TCOD"
CreateDir(caoTcodLogDir)

'=========================================================================================================================================
' Log files. Create (if not exist) files for logs - monthly \YYYY.MM
YYYYMM = Mid(DateNowString, 1, 7)

caoLogFilePath = caoLogDir & "\" & YYYYMM & "_Log.txt"
CreateLogFile caoLogFilePath, False

zaoLogFilePath = zaoLogDir & "\" & YYYYMM & "_Log.txt"
CreateLogFile zaoLogFilePath, False

uzaoLogFilePath = uzaoLogDir & "\" & YYYYMM & "_Log.txt"
CreateLogFile uzaoLogFilePath, False

caoTcodLogFilePath = caoTcodLogDir & "\" & YYYYMM & "_Log.txt"
CreateLogFile caoTcodLogFilePath, False

' Log Errors
caoLogErrorsFilePath = caoLogDir & "\" & YYYYMM & "_LogErrors.txt"
CreateLogFile caoLogErrorsFilePath, False

zaoLogErrorsFilePath = zaoLogDir & "\" & YYYYMM & "_LogErrors.txt"
CreateLogFile zaoLogErrorsFilePath, False

uzaoLogErrorsFilePath = uzaoLogDir & "\" & YYYYMM & "_LogErrors.txt"
CreateLogFile uzaoLogErrorsFilePath, False

caoTcodLogErrorsFilePath = caoTcodLogDir & "\" & YYYYMM & "_LogErrors.txt"
CreateLogFile caoTcodLogErrorsFilePath, False

' Log Warnings
caoLogWarningsFilePath = caoLogDir & "\" & YYYYMM & "_LogWarnings.txt"
CreateLogFile caoLogWarningsFilePath, False

zaoLogWarningsFilePath = zaoLogDir & "\" & YYYYMM & "_LogWarnings.txt"
CreateLogFile zaoLogWarningsFilePath, False

uzaoLogWarningsFilePath = uzaoLogDir & "\" & YYYYMM & "_LogWarnings.txt"
CreateLogFile uzaoLogWarningsFilePath, False

caoTcodLogWarningsFilePath = caoTcodLogDir & "\" & YYYYMM & "_LogWarnings.txt"
CreateLogFile caoTcodLogWarningsFilePath, False

' Add new horizontalLine & information about this script execution (time)
horizontalLine = "------------------------------------------------------------------------------------------------------------------------"
LogToFile caoLogFilePath, horizontalLine
LogToFile zaoLogFilePath, horizontalLine
LogToFile uzaoLogFilePath, horizontalLine
LogToFile caoTcodLogFilePath, horizontalLine

LogToFile caoLogErrorsFilePath, horizontalLine
LogToFile zaoLogErrorsFilePath, horizontalLine
LogToFile uzaoLogErrorsFilePath, horizontalLine
LogToFile caoTcodLogErrorsFilePath, horizontalLine

LogToFile caoLogWarningsFilePath, horizontalLine
LogToFile zaoLogWarningsFilePath, horizontalLine
LogToFile uzaoLogWarningsFilePath, horizontalLine
LogToFile caoTcodLogWarningsFilePath, horizontalLine

LogToFile caoLogFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile zaoLogFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile uzaoLogFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile caoTcodLogFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString

LogToFile caoLogErrorsFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile zaoLogErrorsFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile uzaoLogErrorsFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile caoTcodLogErrorsFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString

LogToFile caoLogWarningsFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile zaoLogWarningsFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile uzaoLogWarningsFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString
LogToFile caoTcodLogWarningsFilePath, "Reset test environment: " & DateNowString & " - " & TimeNowString

'=========================================================================================================================================
' Test Results. Create (if not exist) folders for the todays testresults	"C:\_qtp\runresults\testresults\CAO\YYYY.MM.DD"
resultsDir = "C:\_qtp\runresults\testresults"
CreateDir(resultsDir)

caoDistrictResultsDir = resultsDir & "\CAO-Test" 'DateNowString
CreateDir(caoDistrictResultsDir)

zaoDistrictResultsDir = resultsDir & "\ZAO-Test"
CreateDir(zaoDistrictResultsDir)

uzaoDistrictResultsDir = resultsDir & "\UZAO-Test"
CreateDir(uzaoDistrictResultsDir)

caoTcodDistrictResultsDir = resultsDir & "\CAO-TCOD"
CreateDir(caoTcodDistrictResultsDir)

todayCAOTestResultsDir = caoDistrictResultsDir & "\" & DateNowString
CreateDir(todayCAOTestResultsDir)

todayZAOTestResultsDir = zaoDistrictResultsDir & "\" & DateNowString
CreateDir(todayZAOTestResultsDir)

todayUZAOTestResultsDir = uzaoDistrictResultsDir & "\" & DateNowString
CreateDir(todayUZAOTestResultsDir)

todayCAOTcodResultsDir = caoTcodDistrictResultsDir & "\" & DateNowString
CreateDir(todayCAOTcodResultsDir)

'=========================================================================================================================================
' Reset test environment.
'=========================================================================================================================================

Wscript.Echo "Kill Word, Excel & Pdf processes"

Call KillAllWord(True)
Call KillAllExcel(True)
Call KillAllPdf(True)

Wscript.Echo "Kill QTP/UFT & ASU processes"
oShell.Run "taskkill /f /im QTPro.exe", , True
oShell.Run "taskkill /f /im UFT.exe", , True
'oShell.Run "taskkill /f /FI ""USERNAME ne система"" /im javaw.exe", , True

'	Checking QTP is launched. Quit QTP.
'	Dim qtApp
'	Set qtApp = CreateObject("QuickTest.Application")
'
'	If  qtApp.launched = True then 
'		Wscript.Echo "QTP is running. Quit QTP."
'		qtApp.Quit
'	Else
'		Wscript.Echo "QTP is not running."
'	End If 
'
'	WScript.Sleep(5000)
'	If  qtApp.launched = True then 
'		Wscript.Echo "CHECKED: QTP is still running."
'	Else
'		Wscript.Echo "CHECKED: QTP is not running."
'	End If

Wscript.Echo ""


'	Terminate QTP & ASU processes NOT owned by system (cause uDeploy is running as 'javaw.exe' system process)
'Dim Process, Service, strOwner, Response
'Set Service = GetObject ("winmgmts:")

'For Each Process in Service.InstancesOf( "win32_process" )
'	If IsObject(Process) Then
'		If UCase( Process.name ) <> UCase("System Idle Process") And UCase( Process.name ) <> UCase("System") Then
'			'Wscript.Echo "" & Process.name
'			Response = Process.GetOwner(strOwner)
'			If Err.Number <> 0 Then
'				messageErrNumberDescription = "Error number: " & Err.Number & "; Err description: " & Err.Description
'				messageErrSource = "	Source: " & Err.Source
'				Err.Clear
'				LogToFileWithEcho caoLogErrorsFilePath, messageErrNumberDescription
'				LogToFileWithEcho caoLogErrorsFilePath, messageErrSource
'				LogToFileWithEcho caoLogErrorsFilePath, "Process.name: " & Process.name
'				WScript.Quit(Err.Number)
'			End If
'			'Wscript.Echo "strOwner: " & UCase(strOwner) & " : " & UCase(strUser) & " : " & Process.name & " : " & Not(strOwner = strUser)
'			'REMOVED FROM IF: Or UCase( Process.name ) = UCase( "QTPro.exe" )
'			If (UCase( Process.name ) = UCase( "javaw.exe" ) ) _
'					And Not(UCase(strOwner) = UCase("СИСТЕМА")) Then
'				'IsProcessRunningNotEqUsername = True
'				Wscript.Echo "Terminate process '" & Process.name & "' owned by '" & strOwner & "'"
'				Process.Terminate()
'				'Set Service = Nothing
'				'Exit Function
'			End If
'		End If
'	End If
'Next

If Err.Number <> 0 Then
	messageErrNumberDescription = "Error number: " & Err.Number & "; Err description: " & Err.Description
	messageErrSource = "	Source: " & Err.Source
	Err.Clear
	LogToFileWithEcho caoLogErrorsFilePath, messageErrNumberDescription
	LogToFileWithEcho caoLogErrorsFilePath, messageErrSource
End If
WScript.Quit(Err.Number)


'=========================================================================================================================================
' Functions part
'=========================================================================================================================================
Function IsProcessRunningNotEqUsername( strProcess, strUser )
    Dim Process, Service, strOwner, Response
	
	Set Service = GetObject ("winmgmts:")
    IsProcessRunningNotEqUsername = False
    
    For Each Process in Service.InstancesOf( "win32_process" )
		If IsObject(Process) Then
			If UCase( Process.name ) <> UCase("System Idle Process") Then
				'Wscript.Echo "" & Process.name
				Response = Process.GetOwner(strOwner)
				If Err.Number <> 0 Then
					LogToFileWithEcho caoLogErrorsFilePath, "Error number: " & Err.Number & "; Err description: " & Err.Description
					Wscript.Echo "	" & Process.name
					WScript.Quit
				End If
				'Wscript.Echo "strOwner: " & UCase(strOwner) & " : " & UCase(strUser) & " : " & Process.name & " : " & Not(strOwner = strUser)
				If UCase( Process.name ) = UCase( strProcess ) And Not(UCase(strOwner) = UCase(strUser)) Then
					IsProcessRunningNotEqUsername = True
					Wscript.Echo "Terminate process '" & Process.name & "' owned by '" & strOwner & "'"
					Process.Terminate()
					'Set Service = Nothing
					'Exit Function
				End If
			End If
		End If
    Next
	Set Service = Nothing
End Function

'=========================================================================================================================================
'Возвращает строку с выравниванием по правому краю добавлением символов слева
'	
'Параметры: str - выравниваемая строка, pad - символ для заполнения пустых мест в результирующей строке, length - ширина строки с выравниванием
'=========================================================================================================================================
Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

'=========================================================================================================================================
'	Возвращает строковое значение текущей даты
'	Формат возвращаемого значения: YYYY.MM.DD
'
'	Author: Gerasimenko I.S.
'=========================================================================================================================================
Function DateNowString ()
	DateNowString = Year(Now) & "." & LPad(Month(Now), "0", 2) & "." & LPad(Day(Now), "0", 2)
End Function
'=========================================================================================================================================
'Возвращает строковое значение текущего времени 
'Формат возвращаемого значения: hh.mm.ss
'	
'Параметры: -
'=========================================================================================================================================
Function TimeNowString ()
	TimeNowString = LPad(Hour(Now), "0", 2) & "." & LPad(Minute(Now), "0", 2) & "." & LPad(Second(Now), "0", 2)
End Function

'=========================================================================================================================================
'	Создаёт директорию по указанному адресу, если директория не существует
'	
'Параметры: dirPath- путь к создаваемой директории
'
'	Author: Gerasimenko I.S.
'=========================================================================================================================================
Function CreateDir (dirPath)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If Not oFSO.FolderExists(dirPath) Then
		Set objFolder = oFSO.CreateFolder(dirPath)
	End If
End Function

'=========================================================================================================================================
'	Создаёт файл по указанному адресу
'	
'Параметры: 
'		- filePath - путь к создаваемому файлу, e.g. "C:\testfile.txt"
'		- text - текст, добавляемый в файл
'		- overwrite - перезаписывать существующий файл
'
'	Author: Gerasimenko I.S.
'=========================================================================================================================================
Sub CreateLogFile (filePath, overwrite)
	Dim fso, MyFile
	Set fso = CreateObject("Scripting.FileSystemObject")	
	
	If Not(fso.FileExists(filePath) And Not overwrite) Then
		Set MyFile = fso.CreateTextFile(filePath, True, 0)
		MyFile.Close
	End If
End Sub

'=========================================================================================================================================
'	Добавляет новую строку с заданным текстом в файл (если файла нет, то он НЕ создаётся)
'	
'Параметры: 
'		- filePath - путь к файлу для записи, e.g. "C:\testfile.txt"
'		- text - текст, добавляемый в файл
'
'	Author: Gerasimenko I.S.
'=========================================================================================================================================
Sub LogToFile (filePath, text)
	Dim fso, MyFile
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.OpenTextFile(filePath, 8, False, 0)
	
	MyFile.WriteLine(text)
	MyFile.Close
End Sub

Sub LogToFileWithEcho (filePath, text)
	LogToFile filePath, text
	Wscript.Echo text
End Sub

'=========================================================================================================================================
'	kill Word, Excel, Pdf processes
'=========================================================================================================================================
Sub KillAllWord(doReport)
	Dim oShell : Set oShell = CreateObject("WScript.Shell")
	oShell.Run "taskkill /im winword.exe /f /t", 0, True
End Sub

Sub KillAllExcel(doReport)
	Dim oShell : Set oShell = CreateObject("WScript.Shell")
	oShell.Run "taskkill /im EXCEL.EXE /f /t", 0, True
End Sub

Sub KillAllPdf(doReport)
	Dim oShell : Set oShell = CreateObject("WScript.Shell")
	oShell.Run "taskkill /im AcroRd32.exe /f /t", 0, True
End Sub