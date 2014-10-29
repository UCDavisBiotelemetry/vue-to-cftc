@echo off
ECHO Error logging wrapper for another DOS .bat file
ECHO Written by Matt Pagel

SET "EX=ConvertVUE2forCFTC"
SET EC="%~dp0%EX%.bat"
FOR /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
SET "YY=%dt:~2,2%" & set "YYYY=%dt:~0,4%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%"
SET "HH=%dt:~8,2%" & set "Min=%dt:~10,2%" & set "Sec=%dt:~12,2%"

SET "EL=%~dp0%EX%-error-%YYYY%%MM%%DD%_%HH%%Min%%Sec%.log"

CALL %EC% %* 2>%EL%
