@echo off
editbin Winpark.exe /subsystem:windows,5.1 > EDITBIN.txt
Winpark.exe | taskkill /F /IM cmd.exe
