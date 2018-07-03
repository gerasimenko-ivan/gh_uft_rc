REM СОЗДАНИЕ РЕЗЕРВНОЙ КОПИИ БАЗЫ ДАННЫХ POSTGRESQL
CLS
ECHO OFF
CHCP 1251

REM Установка переменных окружения
SET PGBIN=C:\Program Files\PostgreSQL\9.3\bin\
SET PGDATABASE=postgres
SET PGHOST=localhost
SET PGPORT=5432
SET PGUSER=postgres
SET PGPASSWORD=1

REM Смена диска и переход в папку из которой запущен bat-файл
%~d0
CD %~dp0

REM Формирование имени файла резервной копии и файла-отчета
SET DATETIME=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2% %TIME:~0,2%-%TIME:~3,2%-%TIME:~6,2%
SET DUMPFILE=%PGDATABASE% %DATETIME%.backup
SET LOGFILE=%PGDATABASE% %DATETIME%.txt
SET DUMPPATH="Backup_localhost\%DUMPFILE%"
SET LOGPATH="Backup_localhost\%LOGFILE%"

REM Создание резервной копии
IF NOT EXIST Backup_localhost MD Backup_localhost
CALL "%PGBIN%\pg_dump.exe" --format=custom --verbose --file=%DUMPPATH% 2>%LOGPATH%

REM Анализ кода завершения
IF NOT %ERRORLEVEL%==0 GOTO Error
GOTO Successfull

REM В случае ошибки удаляется поврежденная резервная копия и делается соответствующая запись в журнале
:Error
DEL %DUMPPATH%
MSG * "Ошибка при создании резервной копии базы данных. Смотрите backup_localhost.txt."
ECHO %DATETIME% Ошибки при создании резервной копии базы данных %DUMPFILE%. Смотрите отчет %LOGFILE%. >> backup_localhost.txt
GOTO End

REM В случае удачного резервного копирования просто делается запись в журнал
:Successfull
ECHO %DATETIME% Успешное создание резервной копии %DUMPFILE% >> backup_localhost.txt
GOTO End

:End