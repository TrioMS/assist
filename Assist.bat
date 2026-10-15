@echo off
rem ===================================================================
rem  ASSIST (eski adi: Is Takip)
rem  Sunucu: Sunucu\istakip_sunucu.ps1   Arayuz: Portal\is_takip.html   Port 8096
rem  (Hukuk portali 8095, HSE 8090.)
rem  Butun acilis isi Sunucu\baslat.ps1'de: PENCERESIZ calisir (gizli.ps1), bu bat
rem  onu baslatip hemen kapanir - siyah pencere yalniz ~1 sn gorunur (15.10.2026).
rem  Sira: sunucu > tepsi simgesi > uygulama penceresi (acilis ekrani) > Outlook /
rem  MiCollab / Teams > MiCollab izleyici. Hatalar mesaj kutusunda gosterilir.
rem  Kayit: %LOCALAPPDATA%\Assist\_baslangic_<BILGISAYAR>.log
rem ===================================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sunucu\gizli.ps1" -Betik "%~dp0Sunucu\baslat.ps1" -Port 8096
