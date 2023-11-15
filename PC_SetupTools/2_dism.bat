@echo off

setlocal enabledelayedexpansion
cd /d %~dp0

rem 管理者権限で実行
openfiles > nul

if "%1"=="" (
set arg=
) else (
set arg= -ArgumentList "%1"
)
if errorlevel 1 (
   PowerShell.exe -Command Start-Process \"%~f0\"%arg% -Verb runas
   exit
)




Dism /online /enable-feature /featurename:NetFX3 /All /Source:D:\sources\sxs






pause

