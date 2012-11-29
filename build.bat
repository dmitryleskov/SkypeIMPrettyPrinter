set PANDOC_HOME=C:\EXE\pandoc-1.8.2
set VERSION=0.1

%PANDOC_HOME%\bin\pandoc.exe -f markdown -t html README.md -o README.html
set ARCHIVE=SkypeIMPrettyPrinter-%VERSION%.zip
if exist %ARCHIVE% del %ARCHIVE%
zip %ARCHIVE% SkypeIMPrettyPrinter.bas README.html LICENSE
del README.html
