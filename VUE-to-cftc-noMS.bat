@echo off
ECHO Error logging wrapper for another DOS .bat file
ECHO Written by Matt Pagel

SET "EX=ConvertVUE2forCFTC"
SET EC="%~dp0%EX%.bat"
FOR /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
SET "YY=%dt:~2,2%" & set "YYYY=%dt:~0,4%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%"
SET "HH=%dt:~8,2%" & set "Min=%dt:~10,2%" & set "Sec=%dt:~12,2%"

SET EL2="%~dp0%EX%-error2-%YYYY%%MM%%DD%_%HH%%Min%%Sec%.log"
SETLOCAL enabledelayedexpansion enableextensions

CALL !EC! /noMS %* 2>!EL2!
REM echo "!EL2! error log"
FOR /F %%A IN ("!EL2!") DO ( 
	set size=%%~zA
rem	echo !size!
)

IF !size! EQU 0 (
  DEL "!EL2!" > NUL
)

REM CALL %EC% %* 2>%EL%
