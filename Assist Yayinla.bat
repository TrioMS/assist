@echo off
rem ===================================================================
rem  ASSIST YAYINLA (yalniz bakim sorumlusu)
rem  Kodda ya da sayfada degisiklik yaptiktan sonra cift tiklayin:
rem  tum betikler ve sayfa imzalanir, Assist kullananlar yeni surume gecer.
rem  Imzalanmazsa imza zorunlu bilgisayarlarda eski surum calismaya devam eder.
rem ===================================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sunucu\imzala.ps1" %*
echo.
pause
