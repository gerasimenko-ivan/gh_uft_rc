CLS
ECHO OFF
CHCP 1251
SCHTASKS /Create  /SC DAILY /TN "CopY_Postgres" /TR "C:\Dump\backup_postgres_localhost.bat" /ST 03:00:00
IF NOT %ERRORLEVEL%==0 MSG * "Ошибка при создании задачи резервного копирования."