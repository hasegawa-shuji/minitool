@echo off

CHCP 932

Robocopy \\filesv01\�d�Zs$\10_����\explaner C:\Users\%USERNAME%.ANOEX_DC\Desktop tnsnames.ora /r:3 /w:5



Xcopy \\filesv01\�d�Zs$\10_����\Oracle_client-11g C:\Users\%USERNAME%.ANOEX_DC\Desktop\Oracle_client-11g /S/I



Xcopy \\filesv01\�d�Zs$\10_����\Flash_Installer C:\Users\%USERNAME%.ANOEX_DC\Desktop\Flash_Installer /S/I




pause