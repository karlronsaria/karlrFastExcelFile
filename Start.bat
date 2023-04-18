@echo off

set "thisDir=%~dp0"
set "cmd=%thisDir%script\MyScript.ps1"

if "%~1" EQU "" goto noParams
set "params=%~1"
goto addParams

:noParams
set "params=%cd%"

:addParams
set "cmd=%cmd% -Path %params%"
powershell -Command %cmd%
