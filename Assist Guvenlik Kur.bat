@echo off
rem ===================================================================
rem  ASSIST GUVENLIK KUR (15.10.2026)
rem  Bu bilgisayarda imza zorunlulugunu acar: bundan sonra Assist yalniz
rem  bakim sorumlusunun imzaladigi kodla calisir. Her kullanici BIR KEZ
rem  calistirir. Geri almak icin komut satirindan:  "Assist Guvenlik Kur.bat" -Kaldir
rem ===================================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sunucu\imza_kur.ps1" %*
echo.
pause
