
forfiles -p "C:\_qtp\runresults\screenshots" -s -m *.* /D -5 /C "cmd /c del /Q /S @path"

forfiles -p "C:\_qtp\runresults\testresults\CAO-Test" -s -m *.* /D -5 /C "cmd /c IF @isdir == TRUE rd /Q /S @path"
forfiles -p "C:\_qtp\runresults\testresults\ZAO-Test" -s -m *.* /D -5 /C "cmd /c IF @isdir == TRUE rd /Q /S @path"
forfiles -p "C:\_qtp\runresults\testresults\SZAO-Test" -s -m *.* /D -5 /C "cmd /c IF @isdir == TRUE rd /Q /S @path"
forfiles -p "C:\_qtp\runresults\testresults\UZAO-Test" -s -m *.* /D -5 /C "cmd /c IF @isdir == TRUE rd /Q /S @path"
forfiles -p "C:\_qtp\runresults\testresults\NAO-Test" -s -m *.* /D -5 /C "cmd /c IF @isdir == TRUE rd /Q /S @path"
forfiles -p "C:\_qtp\runresults\testresults\CAO-TCOD" -s -m *.* /D -5 /C "cmd /c IF @isdir == TRUE rd /Q /S @path"