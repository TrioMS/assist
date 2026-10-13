@echo off
setlocal
set PORT=8096

rem ===================================================================
rem  ASSIST (eski adi: Is Takip)
rem  Sunucu: Sunucu\istakip_sunucu.ps1   Arayuz: Portal\is_takip.html
rem  Hukuk portali 8095, HSE 8090 - is takip AYRI PORTTA (8096) durur.
rem  Sunucu GIZLI calisir (pencere yok); saatin yanindaki simge
rem  (Sunucu\tepsi.ps1) ile acilir / kapatilir. Cikti kayda yazilir:
rem  Sunucu\_baslangic_<BILGISAYAR>.log
rem ===================================================================
set "KOK=%~dp0"

rem 1) Port dinleniyor mu?
powershell -NoProfile -Command "$c=New-Object System.Net.Sockets.TcpClient;try{$c.Connect('localhost',%PORT%);$c.Close();exit 0}catch{exit 1}" >nul 2>&1
if %ERRORLEVEL% NEQ 0 goto startserver

rem 2) Dinliyor - bu kopyanin sunucusu mu, kodu guncel mi?
rem    Cikis: 0 = guncel | 4 = eski kapatildi, yeniden baslat | 3 = baska kopya | 2 = tanimsiz
rem    Cevap yoksa ama portu bizim (takilmis) sunucumuz tutuyorsa: kapatilip yeniden baslatilir.
powershell -NoProfile -Command "$f=Join-Path $env:KOK 'Sunucu\istakip_sunucu.ps1'; $exp=(Resolve-Path -LiteralPath $env:KOK).Path.TrimEnd('\'); try{$j=(Invoke-WebRequest -Uri 'http://localhost:%PORT%/__saglik' -UseBasicParsing -TimeoutSec 8).Content | ConvertFrom-Json}catch{ $t=@(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'powershell.exe' -and $_.ProcessId -ne $PID -and $_.CommandLine -match '-File' -and $_.CommandLine -match 'sunucu_calistir\.ps1|istakip_sunucu\.ps1' }); if($t.Count){ Write-Host '  Sunucu cevap vermiyor - yeniden baslatiliyor...'; $t | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; Start-Sleep -Seconds 2; exit 4 }; exit 2}; if($j.urun -ne 'istakip'){exit 2}; if($j.portalKok.TrimEnd('\') -ine $exp){ Write-Host ('  Calisan sunucunun koku: ' + $j.portalKok); exit 3 }; $z=[string](Get-Item -LiteralPath $f).LastWriteTimeUtc.Ticks; if([string]$j.betikZaman -eq $z){exit 0}; Write-Host '  Sunucu dosyasi guncellenmis - yeniden baslatiliyor...'; Stop-Process -Id $j.pid -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 2; exit 4"
set RC=%ERRORLEVEL%
if "%RC%"=="0" goto tepsi
if "%RC%"=="4" goto startserver
echo.
if "%RC%"=="3" echo  UYARI: %PORT% portunda BASKA bir kopyanin Assist sunucusu calisiyor.
if "%RC%"=="2" echo  UYARI: %PORT% portunda tanimlanamayan bir uygulama var.
echo  Cozum: o pencereyi kapatip bu dosyayi yeniden calistirin.
echo.
pause
exit /b 1

:startserver
echo Assist baslatiliyor...
rem PENCERESIZ (gizli.ps1): Windows Terminal varsayilan konsolken -WindowStyle Hidden
rem dikkate alinmiyor, bos pencere kaliyordu. Cikti ve hatalar kayda yazilir.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sunucu\gizli.ps1" -Betik "%~dp0Sunucu\sunucu_calistir.ps1"

rem Port cevap verene kadar bekle (en cok 30 sn).
powershell -NoProfile -Command "$s=[Diagnostics.Stopwatch]::StartNew(); while($s.Elapsed.TotalSeconds -lt 30){ try{ $c=New-Object Net.Sockets.TcpClient; $c.Connect('localhost',%PORT%); $c.Close(); exit 0 }catch{ Start-Sleep -Milliseconds 400 } }; exit 1"
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo  SUNUCU ACILAMADI. Hata kaydi: %~dp0Sunucu\_baslangic_%COMPUTERNAME%.log
  echo.
  start "" notepad "%~dp0Sunucu\_baslangic_%COMPUTERNAME%.log"
  pause
  exit /b 1
)

:tepsi
rem Saatin yanindaki simge (tek kopya: zaten aciksa yenisi acilmaz).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sunucu\gizli.ps1" -Betik "%~dp0Sunucu\tepsi.ps1" -Port %PORT%
rem Ayarlar > Acilista baslat: Outlook / MiCollab (varsayilan acik), Teams (varsayilan kapali).
rem Kapali olanlar baslatilir - sunucu zaten aciksa da.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sunucu\uygulamalar.ps1"
rem MiCollab pencere izleyicisi (Emergency penceresini kapatir, listeyi saatin yanina gizler; tek kopya).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sunucu\gizli.ps1" -Betik "%~dp0Sunucu\micollab_izle.ps1"

:openbrowser
set "ADRES=http://localhost:%PORT%"
powershell -NoProfile -Command "try{ (Invoke-WebRequest -Uri 'http://localhost:%PORT%/__saglik' -UseBasicParsing -TimeoutSec 5) | Out-Null; exit 0 }catch{ exit 1 }"
if %ERRORLEVEL% NEQ 0 set "ADRES=http://127.0.0.1:%PORT%"
start "" "%ADRES%/?v=%RANDOM%%RANDOM%"
