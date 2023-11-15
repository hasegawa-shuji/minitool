@echo off


REM 年月日をセット
setlocal

set year=%date:~0,4%
set month=%date:~5,2%
set day=%date:~8,2%
set date2=%year%%month%%day%


REM 構成をテキストファイルに保存
ipconfig /all >\\filesv01\ipconfig$\%COMPUTERNAME%.txt

REM ファイル名を変更
rename \\filesv01\ipconfig$\%COMPUTERNAME%.txt %COMPUTERNAME%_%date2%.txt

endlocal

pause

