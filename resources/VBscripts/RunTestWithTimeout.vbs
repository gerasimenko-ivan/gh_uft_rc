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

'Variables
Dim objArgs
Dim oShell : Set oShell = CreateObject("WScript.Shell")
Dim exitCode
exitCode = -1

'Getting the test path
Set objArgs = wscript.Arguments
testPath = objArgs(0)

'Setting results folder
testResultPath = "C:\Test Logs and Results\" & DatetimeNowString & "__" & GetLastNameInPath(testPath)

'Determining that the test does exist
Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
DoesFolderExist = objFSO.FolderExists(testPath)
Set objFSO = Nothing

If DoesFolderExist Then
    Dim qtApp 'Declare the Application object variable
    Dim qtTest 'Declare a Test object variable
    Set qtApp = CreateObject("QuickTest.Application") 'Create the Application object
    qtApp.Launch 'Start QuickTest
    qtApp.Visible = True 'Make the QuickTest application visible
    qtApp.Open testPath, False 'Open the test in read-only mode
    Set qtTest = qtApp.Test

    Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object
    qtResultsOpt.ResultsLocation = testResultPath ' Specify the location to save the test results.
    qtTest.Run qtResultsOpt,True 'Run the test and wait until end of the test run
	
	'Write to console test execution status
	Wscript.Echo "Run status: " & TestResultCode(qtTest.LastRunResults.Status)
	exitCode = TestResultCode(qtTest.LastRunResults.Status)
	
    'qtTest.Run 'Run the test
    qtTest.Close 'Close the test
	'we do not need to close QTP after each test
    'qtApp.Quit
	
	'kill timer which is used to close test that run too long
	Wscript.Echo "kill timer"
	oShell.Run "taskkill /fi ""WINDOWTITLE eq timeouttask*"" /f /t", , True
	Wscript.Echo "timer is killed"
		
Else
	'Couldn't find the test folder. That's bad. Guess we'll have to report on how we couldn't find the test.
    'Insert reporting mechanism here.
	Wscript.Echo "Run status: " & TestResultCode("N/A")
	
	Wscript.Echo "kill timer"
	oShell.Run "taskkill /fi ""WINDOWTITLE eq timeouttask*"" /f /t", , True
	Wscript.Echo "timer is killed"
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.CreateFolder testResultPath & "__TEST_NOT_FOUND"
	Set objFile = objFSO.CreateTextFile(testResultPath & "__TEST_NOT_FOUND\log.txt")
	objFile.WriteLine "Execution completed."
	objFile.WriteLine "Error: test path '" & testPath & "' not found."
	objFile.Close
	Set objFSO = Nothing
	
	exitCode = 4
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
			testResultCode = 3
		Case "No test folder"
			testResultCode = 4
		Case Else	'Possible statuses: Paused, Not Completed...?
			testResultCode = 666
	End Select
	TestResultCode = testResultCode
End Function
'=========================================================================================================================================