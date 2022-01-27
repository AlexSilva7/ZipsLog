@echo off

set logPath=C:\intelitrader\log\InteliOrder
set folderZipDestination=C:\intelitrader\log\InteliOrder
set logs=%logPath%\*.log

rem format date from: https://superuser.com/questions/1086136/batch-script-cmd-resulting-ind-dd-mm-yyyy-weekday-format

for /f "usebackq skip=1 tokens=1-3" %%g in ('wmic Path Win32_LocalTime Get Day^,Month^,Year ^| findstr /r /v "^S"') do (
    set _day = 00%%g
    set _month=00%%h
    set year=00%%i
)

set _month=%_month:~-2%
set _day=%_day:~-2%

for /f %%k in ('powershell ^(get-date^).DayOfWeek') do (
    set _dow=%%k
)

set strDate=%_day%-%_month%-%_year%
set logZipDestination=%folderZipDestination%\%strDate%.7zip

forfiles /s /p %logPath% /m *.log /D -3 /C "cmd /c 7z a -sdel %logZipDestination% @path"
