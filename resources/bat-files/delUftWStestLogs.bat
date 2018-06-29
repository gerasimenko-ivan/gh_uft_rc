@echo off
del C:\_qtp\runresults\screenshots /Q
del C:\WS_tests\results\history /Q

rmdir C:\_qtp\runresults\testresults\ZAO-Test /S /Q
mkdir C:\_qtp\runresults\testresults\ZAO-Test

rmdir C:\_qtp\runresults\testresults\CAO-Test /S /Q
mkdir C:\_qtp\runresults\testresults\CAO-Test

rmdir C:\_qtp\runresults\testresults\UZAO-Test /S /Q
mkdir C:\_qtp\runresults\testresults\UZAO-Test

rmdir C:\_qtp\runresults\testresults\CAO-TCOD /S /Q
mkdir C:\_qtp\runresults\testresults\CAO-TCOD

rmdir C:\_qtp\runresults\testresults\NAO-Test /S /Q
mkdir C:\_qtp\runresults\testresults\NAO-Test

rmdir C:\_qtp\runresults\testresults\SZAO-Test /S /Q
mkdir C:\_qtp\runresults\testresults\SZAO-Test