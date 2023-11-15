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

rem 資格情報マネージャーにNASのフォルダにアクセスする権限を与える
cmdkey /add:172.20.6.12 /user:VHX-8000 /pass:VHX

rem ローカルフォルダをNASフォルダと同期
robocopy /copy:DT d:\CommonData \\172.20.6.12\Keyence_VHX-8000\CommonData /mir /r:3 /w:5



pause

