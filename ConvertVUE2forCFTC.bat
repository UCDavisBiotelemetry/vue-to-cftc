@echo off
SET EX=perl.exe "%~dp0ConvertVUE2forCFTC.pl"

REM if this program tells you it can't find perl, either add your perl installation directory to your DOS %PATH% or explicitly reference as below
REM SET EX=D:\programs\strawberry\perl\bin\perl.exe "%~dp0ConvertVUE2forCFTC.pl"
REM Set path above to appropriate Perl install directory

ECHO Batch file for mass conversion of VUE2 CSVs
ECHO Written by Matt Pagel, UC Davis, October 2014

REM timestamp output file
FOR /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
SET "YY=%dt:~2,2%" & set "YYYY=%dt:~0,4%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%"
SET "HH=%dt:~8,2%" & set "Min=%dt:~10,2%" & set "Sec=%dt:~12,2%"
SET FN="%~dp0%~n0-output-%YYYY%%MM%%DD%_%HH%%Min%%Sec%.csv"
SET EL="%~dp0error.log"
copy /y NUL %FN% > NUL
copy /y NUL %EL% > NUL
SETLOCAL enabledelayedexpansion enableextensions
SET argCount=0
SET failures=0
SET cmdXtra=""
REM Check for any command-line parameters other than file names
FOR %%x in (%*) do (
   IF "%%x"=="/noMS" SET cmdXtra=/noMS
   IF "%%x"=="/yesMS" SET cmdXtra=/yesMS
)
echo !cmdXtra!
FOR %%x in (%*) do (
   echo %%x
   IF "%%x" NEQ "/noMS" (
      IF "%%x" NEQ "/yesMS" (
         SET /A argCount+=1
         SET PF="%%~fx"
         echo.
         echo PROCESSING !argCount!: !EX! !cmdXtra! -from- !PF! -to- !FN!
REM Feel free to REM the two echo lines immediately above if you don't want to receive notification of each file processing
REM Supression may actually be handy if you are encountering a lot of file processing failures in order to see all those errors
         CMD /C !EX! !cmdXtra! < !PF! >> !FN! 2>!EL!
         FOR /F %%A IN ("!EL!") DO set size=%%~zA
         IF !size! NEQ 0 (
            echo Error message in !PF!: 1>&2
            type "!EL!" 1>&2
            echo. 1>&2
            echo Common problem #1 Incorrect Perl path or missing Date::Calc library from Perl 1>&2
            echo #2 XLS rather than CSV format or CSV with tabs as the delimiter between cols 1>&2
REM Feel free to REM the 3 echo statements immediately above if you don't want the same exact text each time
            set /A failures+=1
   )  )  )
   DEL !EL! > NUL
)
REM If not interested in having the window stick around after program execution, REM the next two lines
ECHO Number of processed files: %argCount% (%failures% failures)
PAUSE
REM if you ever intend to tie this batch file into a different script (e.g. a GUI overlay or a more complex error logger), send the appropriate exit code
exit /b %failures%
