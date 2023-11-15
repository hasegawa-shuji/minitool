@echo off

chcp 65001

xcopy \\filesv01\lotty$\00_Programs\50_Now C:\USERS\%USERNAME%\Desktop\LOTTY_フォルダ /C /E /I /Y

chcp 932

pause