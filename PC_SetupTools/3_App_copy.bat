@echo off

CHCP 932

Robocopy \\filesv01\電算s$\10_共通\explaner C:\Users\%USERNAME%.ANOEX_DC\Desktop tnsnames.ora /r:3 /w:5



Xcopy \\filesv01\電算s$\10_共通\Oracle_client-11g C:\Users\%USERNAME%.ANOEX_DC\Desktop\Oracle_client-11g /S/I



Xcopy \\filesv01\電算s$\10_共通\Flash_Installer C:\Users\%USERNAME%.ANOEX_DC\Desktop\Flash_Installer /S/I




pause