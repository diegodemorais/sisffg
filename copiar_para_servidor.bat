@echo off
ping -n 2 -w 1000 0.0.0.1 > nul 
xcopy D:\Sistemas\"SisFFG (desenv)"\* \\192.168.0.100\d\"SisFFG (desenv)"\ /Y/E/K