REM �������� ��������� ����� ���� ������ POSTGRESQL
CLS
ECHO OFF
CHCP 1251

REM ��������� ���������� ���������
SET PGBIN=C:\Program Files\PostgreSQL\9.3\bin\
SET PGDATABASE=postgres
SET PGHOST=localhost
SET PGPORT=5432
SET PGUSER=postgres
SET PGPASSWORD=1

REM ����� ����� � ������� � ����� �� ������� ������� bat-����
%~d0
CD %~dp0

REM ������������ ����� ����� ��������� ����� � �����-������
SET DATETIME=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2% %TIME:~0,2%-%TIME:~3,2%-%TIME:~6,2%
SET DUMPFILE=%PGDATABASE% %DATETIME%.backup
SET LOGFILE=%PGDATABASE% %DATETIME%.txt
SET DUMPPATH="Backup_localhost\%DUMPFILE%"
SET LOGPATH="Backup_localhost\%LOGFILE%"

REM �������� ��������� �����
IF NOT EXIST Backup_localhost MD Backup_localhost
CALL "%PGBIN%\pg_dump.exe" --format=custom --verbose --file=%DUMPPATH% 2>%LOGPATH%

REM ������ ���� ����������
IF NOT %ERRORLEVEL%==0 GOTO Error
GOTO Successfull

REM � ������ ������ ��������� ������������ ��������� ����� � �������� ��������������� ������ � �������
:Error
DEL %DUMPPATH%
MSG * "������ ��� �������� ��������� ����� ���� ������. �������� backup_localhost.txt."
ECHO %DATETIME% ������ ��� �������� ��������� ����� ���� ������ %DUMPFILE%. �������� ����� %LOGFILE%. >> backup_localhost.txt
GOTO End

REM � ������ �������� ���������� ����������� ������ �������� ������ � ������
:Successfull
ECHO %DATETIME% �������� �������� ��������� ����� %DUMPFILE% >> backup_localhost.txt
GOTO End

:End