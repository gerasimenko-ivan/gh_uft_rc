
Dim App
Set App = CreateObject("QuickTest.Application")

App.Launch
App.Visible = True

Dim QTP_Tests(16)
For i = 1 to UBound (QTP_Tests)
	QTP_Tests(i) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.005 EPD"
Next



testRunFolder = DatetimeNowString & "_run"
resultsFolder = "C:\_qtp\runresults\" & testRunFolder
CreateDir(resultsFolder)

Set res_obj = CreateObject("QuickTest.RunResultsOptions")

logFilePath = resultsFolder & "\Log.txt"
CreateLogFile logFilePath, "Smoke tests set log. Started execution at " & DatetimeNowString

logErrorFilePath = resultsFolder & "\LogError.txt"
CreateLogFile logErrorFilePath, "Error reporting log:"

Wscript.Echo "Saving execution results to: '" & resultsFolder & "'"
Wscript.Echo "Starting tests..."
Wscript.Echo "Run status      - Test"
Wscript.Echo ""

For i = 1 to UBound (QTP_Tests)
	LogToFile logFilePath, ""
	LogToFile logFilePath, QTP_Tests(i) ' Full path used cause "Test destination changes, but log remains the same!"
	
	App.Open QTP_Tests(i), True
	Set QTP_Test = App.Test
	res_obj.ResultsLocation = resultsFolder & "\" & LPad(i,"0",2) & "." & GetLastNameInPath(QTP_Tests(i)) ' Set the results location
	
	testStartTimeMessage = "    start at: " & DatetimeNowString
	LogToFile logFilePath, testStartTimeMessage
	
	QTP_Test.Run res_obj, True
	
	testEndAtTimeMessage = "    end at:   " & DatetimeNowString
	LogToFile logFilePath, testEndAtTimeMessage
	
	testRunResultMessage = "    result:   " & QTP_Test.LastRunResults.Status
	LogToFile logFilePath, testRunResultMessage
	Wscript.Echo LPad(QTP_Test.LastRunResults.Status, " ", 15) & " - " & QTP_Tests(i)
	
	If Not (QTP_Test.LastRunResults.Status = "Passed") Then		' Error logging block
		LogToFile logErrorFilePath, ""
		LogToFile logErrorFilePath, QTP_Tests(i)
		LogToFile logErrorFilePath, testStartTimeMessage
		LogToFile logErrorFilePath, testEndAtTimeMessage
		LogToFile logErrorFilePath, testRunResultMessage
	End If
	
	QTP_Test.Close
	Wscript.Sleep 1000
Next

'App.Quit
Set res_obj = nothing
Set QTP_Test = nothing
Set App = nothing

	   
Wscript.Echo ""
Wscript.Echo "Test set execution completed."
	   
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
'	Возвращает строковое значение текущей даты и времени 
'	Формат возвращаемого значения: YYYY.MM.DD__hh.mm.ss
'
'	Author: Gerasimenko I.S.
'=========================================================================================================================================
Function DatetimeNowString ()
	dateString = Year(Now) & "." & LPad(Month(Now), "0", 2) & "." & LPad(Day(Now), "0", 2)
	timeString = LPad(Hour(Now), "0", 2) & "." & LPad(Minute(Now), "0", 2) & "." & LPad(Second(Now), "0", 2)

	DatetimeNowString = dateString & "_" & timeString
End Function
'=========================================================================================================================================
'=========================================================================================================================================
'Возвращает последнее имя в пути до знака "\"
'
'Параметры: path - строковое значение пути
'=========================================================================================================================================
Function GetLastNameInPath (path)
	GetLastNameInPath = Right(path, Len(path) - InStrRev(path, "\"))
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
'=========================================================================================================================================
'	Создаёт файл по указанному адресу, заносит в него строку из параметра (если файл есть, то он удаляется и создаётся новый)
'	
'Параметры: 
'		- filePath - путь к создаваемому файлу, e.g. "C:\testfile.txt"
'		- text - текст, добавляемый в файл
'
'	Author: Gerasimenko I.S.
'=========================================================================================================================================
Sub CreateLogFile (filePath, text)
	Dim fso, MyFile
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.CreateTextFile(filePath, True, 0)
	
	MyFile.WriteLine(text)
	MyFile.Close
End Sub

'=========================================================================================================================================
'=========================================================================================================================================
'	Добавляет новую строку с заданным текстом в файл (если файла нет, то он создаётся)
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
	Set MyFile = fso.OpenTextFile(filePath, 8, True, 0)
	
	MyFile.WriteLine(text)
	MyFile.Close
End Sub