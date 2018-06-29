'	usage example: cscript.exe "C:\_qtp\resources\VBscripts\update-jnlp-file.vbs" CAO-Test

'	ARG: "CAO-Test"

Dim objArgs
Set objArgs = wscript.Arguments

Set oShell = WScript.CreateObject("WSCript.shell")

'-------------------------------------------------------------------------------------------------------------------
' 	Input params:
districtAndEnvironment = UCase(objArgs(0))

' 	get district & environment
dashPosition = InStr(districtAndEnvironment, "-")

district = UCase(Left(districtAndEnvironment, dashPosition - 1))
environment = UCase(Right(districtAndEnvironment, Len(districtAndEnvironment) - dashPosition))

Wscript.Echo "district = " & district
Wscript.Echo "environment = " & environment



serverUrl = ""
jarsDownloadUrl = ""
jnlpFileName = ""
ruAsuTitle = ""
'.JAR download url:
Select Case environment
	Case "PROD"
		'	host:7003/cbs-web/client/
		serverUrl = LCase(district) & ".wl.eirc.mos.ru"
		jarsDownloadUrl = "http://" & LCase(district) & ".host:7003/cbs-web/client/"
		jnlpFileName = "cbs-client-" & LCase(district) & "-wls.jnlp"
		ruAsuTitle = "Клиент " & ruDistrict & " (Веблоджик)"
	Case "TCOD"
		'	host:7003/cbs-web/client/
		serverUrl = LCase(district) & "0.wl.test.eirc.mos.ru"
		jarsDownloadUrl = "http://" & LCase(district) & "0.host:7003/cbs-web/client/"
		jnlpFileName = "cbs-client-" & LCase(district) & "-db0-wls.jnlp"
		ruAsuTitle = "Клиент " & ruDistrict & " (Тест-ТЦОД) Веблоджик"
	Case "TEST"
		'	cao-t.wl.test.eirc.mos.ru:7003/cbs-web/client/
		serverUrl = LCase(district) & "-t.wl.test.eirc.mos.ru"
		jarsDownloadUrl = "http://" & LCase(district) & "-t.host:7003/cbs-web/client/"
		jnlpFileName = "cbs-client-" & LCase(district) & "-test-em-wls.jnlp"
		ruAsuTitle = "Клиент " & ruDistrict & "-Т (Тест ЭМ) Веблоджик"
	Case Else
    	Wscript.Echo "Unexpected value of 2nd argument (environment): " & environment
		Wscript.Echo "	expected values: PROD, TEST, TCOD"
		WScript.Quit(1)
		ExitTest
End Select



Wscript.Echo ""
Wscript.Echo "----------------------------------------------------------"
Wscript.Echo "Downloading JNLP-file..."
oShell.run "cscript.exe C:\asu-eirc-monitoring\asu-eirc-tests\download-file.vbs " & jarsDownloadUrl & jnlpFileName & " C:\_qtp\resources\jnlp\ " & districtAndEnvironment & ".jnlp", 1, True

Wscript.Echo "    JNLP-file downloaded."