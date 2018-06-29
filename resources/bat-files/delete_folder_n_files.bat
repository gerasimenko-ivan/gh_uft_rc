rem set folder="C:\Users\gis\AppData\Local\Temp"
rem cd /d %folder%
cd /d %1
for /F "delims=" %%i in ('dir /b') do (rmdir "%%i" /s/q || del "%%i" /s/q)