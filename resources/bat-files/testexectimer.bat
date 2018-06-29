
timeout /t %~1
taskkill /fi "WINDOWTITLE eq C:\Windows\system32\cmd.exe*" /f /t

exit