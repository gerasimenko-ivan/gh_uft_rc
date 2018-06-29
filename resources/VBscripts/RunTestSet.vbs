
' How to run this:
'
'	cscript.exe "C:\_qtp\resources\VBScripts\RunTestSet.vbs" "ZAO"

Dim App
Set App = CreateObject("QuickTest.Application")

Dim objArgs
Set objArgs = wscript.Arguments

App.Launch
App.Visible = True


Dim QTP_Tests(80)

j = 0

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\_Ne opredelen\S.001 Otkritie glavnih modulei"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.001 Sozdaniye novogo LS s zaved jilca+udaleniye"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.002 TO arest FLS"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.003 TO audit FLS"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.004 TO doli adresa lyudi sushh LS sozdanie zhiltsa"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.005 TO parametry"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.006 TO svobodnaya ploschad"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.007 TO udobstva dobav redakt udalenie atributa"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.008 TO udobstva dobav redakt udalenie"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.009 TO vrem reg 6 mes"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.010 Operacii s kartochkoi"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.011 TO doli adresa lyudi sushh LS proverka prozhivaniya na2x dolyakh"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.001 Perepaschet za vrem ubitiye"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.002 TO group oper"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.003 TO skidki"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.004 Documents"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.005 EPD"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.006 ODPU"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.007 Pechat_TO_1"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.008 Pechat_TO_2"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.009 Pechat_TO_3"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.001 arch doc dobavleniye redaktirovaniye udaleniye"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.002 perepropiska blokirovka+razblokirovka"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.003 Vipiska+vosstanovleniye propiski vremennaya vipiska"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.004 Vzrosli zitel registrac in grazhd"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.005 Vzrosli zitel rodstvenniye otnosheniya"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie Lica\S.006 Vzrosli zitel. Vidat novii pasport"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie Lica\S.007 Vzrosli zitel. Izmenenie FIO"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.001 Pechat adresni listok prebitiya"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.002 Rebenok"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.003 RegCard_ORStyle"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.004 AdultCitizen_ORStyle"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.005 Vipiska_ORStyle propiska+vipiska+vosstanovlenie. Vzrosli"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.006 Vipiska_ORStyle Vipiska+vosstanovlenie. Vzrosli i rebenok vipisani"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.007 Vipiska_ORStyle Vzrosli vipisan Rebenok prikreplen k druomu jilcu"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.008 Vipiska_ORStyle Vzrosli vipisan Rebenok propisan samostoyat"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.009 ORStyle Prichina smert Vzrosli i rebenok vipisani"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.010 ORStyle Prichina smert Vzrosli vipisan Rebenok prikreplen k druomu jilcu"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.011 ORStyle Prichina smert Vzrosli vipisan Rebenok propisan samostoyat"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\S.001 Raschiren poisk"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\S.002 Doma i operacii s nimi.Sozdanie"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\S.003 Redaktirovanie atributov"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\S.004 Full cycle"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\M.001 Doma i operacii s nimi.Redaktirovanie"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\M.002 BTI report"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\S.001 Pechat Dogovor+Dop soglashenie"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\S.002 Sozdanie UL"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\S.003 Sozdanie dogovora"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\M.001 Pechat okno Oboroti"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\M.002 Izmenenie uchastnikov dogovora"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\M.003 Dobavlenie uslugi"

For i = j + 1 to UBound (QTP_Tests)
	QTP_Tests(i) = QTP_Tests(i-j)
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
	
	' Add parameters to the test
	QTP_Test.Environment("District") = objArgs(0)
	
	res_obj.ResultsLocation = resultsFolder & "\" & LPad(i,"0",2) & "." & GetLastNameInPath(QTP_Tests(i)) ' Set the results location
	
	testStartTimeMessage = "    start at: " & DatetimeNowString
	LogToFile logFilePath, testStartTimeMessage
	
	QTP_Test.Run res_obj, True
	
	testEndAtTimeMessage = "    end at:   " & DatetimeNowString
	LogToFile logFilePath, testEndAtTimeMessage
	
	testRunResultMessage = "    result:   " & QTP_Test.LastRunResults.Status
	LogToFile logFilePath, testRunResultMessage
	Wscript.Echo "#" & LPad(i, " ", 3) & " " & LPad(QTP_Test.LastRunResults.Status, " ", 11) & " - " & QTP_Tests(i)
	
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

App.Quit
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
	If length < Len(str) Then
		LPad = str
	Else
		LPad = String(length - Len(str), pad) & str
	End If
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

' Next dummy function is used just to store  full test set, and it could be copied (full or a part of it) to the above part of script
Function Dummy()
Dim QTP_Tests(200)

For i = 1 to UBound (QTP_Tests)
	QTP_Tests(i) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.004 Documents"
Next

' j=1
' j+1
' !!! j = 0
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\_Ne opredelen\Monitoring proizvoditelnosti"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\_Ne opredelen\S.001 Otkritie glavnih modulei"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.001 Sozdaniye novogo LS s zaved jilca+udaleniye"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.002 TO arest FLS"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.003 TO audit FLS"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.004 TO doli adresa lyudi sushh LS sozdanie zhiltsa"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.005 TO parametry"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.006 TO svobodnaya ploschad"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.007 TO udobstva dobav redakt udalenie atributa"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.008 TO udobstva dobav redakt udalenie"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.009 TO vrem reg 6 mes"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.010 Operacii s kartochkoi"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\S.011 TO doli adresa lyudi sushh LS proverka prozhivaniya na2x dolyakh"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.001 Perepaschet za vrem ubitiye"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.002 TO group oper"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.003 TO skidki"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.004 Documents"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.005 EPD"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.006 ODPU"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.007 Pechat_TO_1"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.008 Pechat_TO_2"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Finansovie Licevie Scheta\M.009 Pechat_TO_3"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.001 arch doc dobavleniye redaktirovaniye udaleniye"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.002 perepropiska blokirovka+razblokirovka"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.003 Vipiska+vosstanovleniye propiski vremennaya vipiska"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.004 Vzrosli zitel registrac in grazhd"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\S.005 Vzrosli zitel rodstvenniye otnosheniya"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie Lica\S.006 Vzrosli zitel. Vidat novii pasport"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie Lica\S.007 Vzrosli zitel. Izmenenie FIO"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.001 Pechat adresni listok prebitiya"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.002 Rebenok"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.003 RegCard_ORStyle"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.004 AdultCitizen_ORStyle"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.005 Vipiska_ORStyle propiska+vipiska+vosstanovlenie. Vzrosli"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.006 Vipiska_ORStyle Vipiska+vosstanovlenie. Vzrosli i rebenok vipisani"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.007 Vipiska_ORStyle Vzrosli vipisan Rebenok prikreplen k druomu jilcu"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.008 Vipiska_ORStyle Vzrosli vipisan Rebenok propisan samostoyat"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.009 ORStyle Prichina smert Vzrosli i rebenok vipisani"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.010 ORStyle Prichina smert Vzrosli vipisan Rebenok prikreplen k druomu jilcu"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Fizicheskie lica\M.011 ORStyle Prichina smert Vzrosli vipisan Rebenok propisan samostoyat"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\S.001 Raschiren poisk"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\S.002 Doma i operacii s nimi.Sozdanie"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\S.003 Redaktirovanie atributov"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\S.004 Full cycle"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\M.001 Doma i operacii s nimi.Redaktirovanie"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Jiloi fond\M.002 BTI report"

QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\S.001 Pechat Dogovor+Dop soglashenie"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\S.002 Sozdanie UL"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\S.003 Sozdanie dogovora"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\M.001 Pechat okno Oboroti"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\M.002 Izmenenie uchastnikov dogovora"
QTP_Tests(IncByRef(j)) = "C:\_qtp\tests\Uridicheskie Lica\M.003 Dobavlenie uslugi"

End Function

Function IncByRef(ByRef Value)
	Value = Value + 1
	IncByRef = Value
End Function