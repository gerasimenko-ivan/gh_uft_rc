'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'RunThisTest
'by Michael Innes
'November 2012
'by Ivan Gerasimenko
'December 2015
'
'	Params:
'		- arg0 - testPath		- path to the test location
'		- arg1 - testTimeLimit	- time limit for test execution
'
' REFACTIOR: create deep logging (start end time, log and logerrors files)
' REFACTIOR: move to functions all actions that could be described as single parametrized block
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' execute a command like this: cscript.exe "C:\RunThisTest.vbs" "L:\Test Path\The Test Itself\"

' Self error handling
On Error Resume Next

'Variables
Dim objArgs
Dim oShell : Set oShell = CreateObject("WScript.Shell")
Dim exitCode
exitCode = -1

'Getting the test path
Set objArgs = wscript.Arguments
testPath = objArgs(0)

'=========================================================================================================================================
'	Test results folder

'Creating (if not exist) results folder AS "C:\_qtp\runresults\testresults\CAO\YYYY.MM.DD"
testDayResultPath = "C:\_qtp\runresults\testresults\" & objArgs(1) & "\" & DateNowString 
CreateDir(testDayResultPath)

'Setting results path AS "C:\_qtp\runresults\testresults\CAO\YYYY.MM.DD\hh.mm.ss_testName"
testResultPath = testDayResultPath & "\" & TimeNowString & "_" & GetLastNameInPath(testPath)

'=========================================================================================================================================
'	Logging preparation

'Creating (if not exist) log files AS "C:\_qtp\runresults\logs\CAO\YYYY.MM_Log.txt"
YYYYMM = Mid(DateNowString, 1, 7)

logFilePath = "C:\_qtp\runresults\logs\" & objArgs(1) & "\" & YYYYMM & "_Log.txt"
CreateLogFile logFilePath, False

logErrorsFilePath = "C:\_qtp\runresults\logs\" & objArgs(1) & "\" & YYYYMM & "_LogErrors.txt"
CreateLogFile logErrorsFilePath, False

logWarningsFilePath = "C:\_qtp\runresults\logs\" & objArgs(1) & "\" & YYYYMM & "_LogWarnings.txt"
CreateLogFile logWarningsFilePath, False

'	startTime for logging
startTime = DateNowString & " | " & TimeNowString

'=========================================================================================================================================
'	Test execution
Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
DoesFolderExist = objFSO.FolderExists(testPath)
	exitCode = CheckAndLogError(Err, "(line 62)DoesFolderExist = objFSO.FolderExists(testPath) [testPath = " & testPath & "]", logErrorsFilePath)
Set objFSO = Nothing

If DoesFolderExist Then
		Wscript.Echo "Test (folder) exists execution block starts."
    Dim qtApp 'Declare the Application object variable
    Dim qtTest 'Declare a Test object variable
    Set qtApp = CreateObject("QuickTest.Application") 'Create the Application object
		exitCode = CheckAndLogError(Err, "(line 70)    Set qtApp = CreateObject(""QuickTest.Application"")", logErrorsFilePath)
    qtApp.Launch 'Start QuickTest
	If (Err.Number <> 0) Then 
		exitCode = CheckAndLogError(Err, "(line 72)	qtApp.Launch", logErrorsFilePath)
		WScript.Sleep(1000)
		Set qtApp = CreateObject("QuickTest.Application")
		WScript.Sleep(1000)
			exitCode = CheckAndLogError(Err, "(line inside if.1)	Set qtApp = CreateObject(""QuickTest.Application"")", logErrorsFilePath)
		qtApp.Launch 'Start QuickTest
			exitCode = CheckAndLogError(Err, "(line inside if.2)	qtApp.Launch", logErrorsFilePath)
	End If
    qtApp.Visible = True 'Make the QuickTest application visible
		exitCode = CheckAndLogError(Err, "(line 74)    qtApp.Visible = True", logErrorsFilePath)
    qtApp.Open testPath, False 'Open the test in read-only mode
		exitCode = CheckAndLogError(Err, "(line 76)    qtApp.Open testPath, False", logErrorsFilePath)
    Set qtTest = qtApp.Test
		exitCode = CheckAndLogError(Err, "(line 78)    Set qtTest = qtApp.Test", logErrorsFilePath)
	
	' Add parameters to the test
	qtTest.Environment("District") = objArgs(1)

    Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object
	qtResultsOpt.ResultsLocation = testResultPath ' Specify the location to save the test results.
		exitCode = CheckAndLogError(Err, "(line 83)	qtResultsOpt.ResultsLocation = testResultPath", logErrorsFilePath)
    qtTest.Run qtResultsOpt, True 'Run the test and wait until end of the test run
		exitCode = CheckAndLogError(Err, "(line 85)    qtTest.Run qtResultsOpt, True", logErrorsFilePath)
		
	'Logging
	LogRow logFilePath, testPath, startTime, TimeNowString, qtTest.LastRunResults.Status
	'Logging errors & warnings
	If Not (qtTest.LastRunResults.Status = "Passed") Then
		LogRow logErrorsFilePath, testPath, startTime, TimeNowString, qtTest.LastRunResults.Status
		If qtTest.LastRunResults.Status = "Warning" Then
			LogRow logWarningsFilePath, testPath, startTime, TimeNowString, qtTest.LastRunResults.Status
		End If
	End If
	
	'Write to console test execution status
	Wscript.Echo "Run status: " & qtTest.LastRunResults.Status
	Wscript.Echo "Run status (returned as ExitCode): " & TestResultCode(qtTest.LastRunResults.Status)
	exitCode = TestResultCode(qtTest.LastRunResults.Status)
	
    'qtTest.Run 'Run the test
    qtTest.Close 'Close the test
	'we do not need to close QTP after each test
    'qtApp.Quit
		
Else
	'Couldn't find the test folder. That's bad. Guess we'll have to report on how we couldn't find the test.
    'Insert reporting mechanism here.
	Wscript.Echo "Run status: Couldn't find the test"
	Wscript.Echo "Run status (returned as ExitCode): " & TestResultCode("N/A")
	
	'Logging
	LogRow logFilePath, testPath, startTime, TimeNowString, "Error: Could not find the test"
	LogRow logErrorsFilePath, testPath, startTime, TimeNowString, "Error: Could not find the test"
		
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.CreateFolder testResultPath & "__TEST_NOT_FOUND"
	Set objFile = objFSO.CreateTextFile(testResultPath & "__TEST_NOT_FOUND\log.txt")
	objFile.WriteLine "Execution completed."
	objFile.WriteLine "Error: test path '" & testPath & "' not found."
	objFile.Close
	Set objFSO = Nothing
	
	exitCode = 4
End If

'=========================================================================================================================================
' Error handling
If Err.Number <> 0 Then
	errorMessage = "Error number " & Err.Number & ", " & Err.Description & ", has occurred"
	' Logging
	LogRow logFilePath, testPath, startTime, TimeNowString, errorMessage
	LogRow logErrorsFilePath, testPath, startTime, TimeNowString, errorMessage
	LogToFile logErrorsFilePath, "	Source: " & Err.Source
	
	WScript.Echo logErrorsFilePath
	WScript.Echo "Error occurred."
	WScript.Echo errorMessage
	
    Err.Clear
	exitCode = 1
End If


WScript.Quit(exitCode)



'=========================================================================================================================================
' Definition of the functions
'=========================================================================================================================================
'Возвращает строку с выравниванием по правому краю добавлением символов слева
'	
'Параметры: str - выравниваемая строка, pad - символ для заполнения пустых мест в результирующей строке, length - ширина строки с выравниванием
'=========================================================================================================================================
Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function
'=========================================================================================================================================
'Возвращает строку с выравниванием по левому краю добавлением символов справа
'	
'Параметры: str - выравниваемая строка, pad - символ для заполнения пустых мест в результирующей строке, length - ширина строки с выравниванием
'=========================================================================================================================================
Function RPad (str, pad, length)
    RPad = str & String(length - Len(str), pad)
End Function
'=========================================================================================================================================

'=========================================================================================================================================
'Возвращает строковое значение текущей даты и времени 
'Формат возвращаемого значения: YYYY.MM.DD__hh.mm.ss
'	
'Параметры: -
'=========================================================================================================================================
Function DatetimeNowString ()
	dateString = Year(Now) & "." & LPad(Month(Now), "0", 2) & "." & LPad(Day(Now), "0", 2)
	timeString = LPad(Hour(Now), "0", 2) & "." & LPad(Minute(Now), "0", 2) & "." & LPad(Second(Now), "0", 2)

	DatetimeNowString = dateString & "__" & timeString
End Function
'=========================================================================================================================================

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
'Возвращает последнее имя в пути до знака "\"
'
'	
'Параметры: path - строковое значение пути
'=========================================================================================================================================
Function GetLastNameInPath (path)
	GetLastNameInPath = Right(path, Len(path) - InStrRev(path, "\"))
End Function
'=========================================================================================================================================

'=========================================================================================================================================
'Возвращает код завершения теста QTP по строковому значению результата
'
'	
'Параметры: testResultString - значение результата выполнения теста QTP
'=========================================================================================================================================
Function TestResultCode (testResultString)
	Select Case testResultString
		Case "N/A"
			testResultCode = -1
		Case "N / A"
			testResultCode = -1
		Case "Passed"
			testResultCode = 0
		Case "Failed"
			testResultCode = 1
		Case "Done"
			testResultCode = 2
		Case "Warning"
			testResultCode = 0	'to see tests that actually failed
		Case "No test folder"
			testResultCode = 4
		Case "Stopped"
			testResultCode = 5
		Case Else	'Possible statuses: Paused, Not Completed...?
			testResultCode = 666
	End Select
	TestResultCode = testResultCode
End Function
'=========================================================================================================================================
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
'	Logging functions
'=========================================================================================================================================
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

'=========================================================================================================================================
'	Добавляет новую строку с заданным текстом в файл формате
'[Test path] | [startTime] -> [endTime] | [results] 
'	
'Параметры: 
'		- filePath - путь к файлу для записи, e.g. "C:\testfile.txt"
'		- testPath, startTime, endTime, results - текст, добавляемый в файл
'
'	Author: Gerasimenko I.S.
'=========================================================================================================================================
Sub LogRow(filePath, testPath, startTime, endTime, results)
	testResultSymbol = ""
	Select Case results
		Case "N/A"
			testResultSymbol = "N"
		Case "N / A"
			testResultSymbol = "N"
		Case "Passed"
			testResultSymbol = "V"
		Case "Failed"
			testResultSymbol = "x"
		Case "Done"
			testResultSymbol = "d"
		Case "Warning"
			testResultSymbol = "!"	'to see tests that actually failed
		Case "No test folder"
			testResultSymbol = "f"
		Case "Stopped"
			testResultSymbol = "□"
		Case Else	'Possible statuses: Paused, Not Completed...?
			testResultSymbol = "?"
	End Select
	
	If Len(testPath) < 71 Then
		LogToFile filePath, RPad(testPath, " ", 70) & " | " & startTime & " -> " & endTime & " | " & testResultSymbol & " | " & results
	Else
		line1 = Mid(testPath, 1, 70)
		line2 = Mid(testPath, 71, 70)
		LogToFile filePath, RPad(line1, " ", 70) & " | " & startTime & " -> " & endTime & " | " & testResultSymbol & " | " &results
		LogToFile filePath, RPad(line2, " ", 70) & " |            |                      |   |"
		If Len(testPath) > 140 Then
			line3 = Mid(testPath, 141, 70)
			LogToFile filePath, RPad(line3, " ", 70) & " |            |                      |   |"
		End If
	End If
	' extra line break for readability
	If InStr(filePath, "LogErrors") > 0 OR InStr(filePath, "LogWarnings") > 0 Then
		LogToFile filePath, ""
	End If
End Sub


Function CheckAndLogError(errObject, text, filePath)
	If errObject.Number <> 0 Then
		errorMessage = "Error number " & errObject.Number & ", " & errObject.Description & ", has occurred"
		' Logging
		LogToFile filePath, "Run time error (" & TimeNowString & "):"
		LogToFile filePath, "	" & errorMessage
		LogToFile filePath, "	Source: " & errObject.Source
		LogToFile filePath, "	Text: " & text
		'LogToFile filePath, "	Current Action name: '" & Environment.Value("ActionName") & "'"
		
		WScript.Echo "Run time error (" & TimeNowString & "):"
		WScript.Echo "	" & errorMessage
		WScript.Echo "	Source: " & errObject.Source
		WScript.Echo "	Current Action name: '" & Environment.Value("ActionName") & "'"
		WScript.Echo "	Text: " & text
		
		errObject.Clear
		
		CheckAndLogError = -1
	Else
		WScript.Echo "	Successfully executed: " & text & "; Err.Number: " & errObject.Number
		CheckAndLogError = 0
	End If
End Function