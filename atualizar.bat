@echo off
ping -n 2 -w 1000 0.0.0.1 > nul 
tskill prFichasG /A
copy \\192.168.0.100\d\Temp\prFichasG.exe .\ /y
start prFichasG.exe