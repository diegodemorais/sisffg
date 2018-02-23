@echo off
ping -n 2 -w 1000 0.0.0.1 > nul
copy \\192.168.0.100\d\"SisFFG (desenv)"\prFichasG.exe d:\Sistemas\"SisFFG (desenv)"\ /Y 
copy \\192.168.0.100\d\"SisFFG (desenv)"\prFichasG.exe \\192.168.0.100\d\temp\ /Y