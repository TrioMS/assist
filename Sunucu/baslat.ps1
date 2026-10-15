# =====================================================================
#  ASSIST - BASLATICI (penceresiz)
#  ---------------------------------------------------------------
#  Assist.bat bu betigi gizli.ps1 ile (konsol penceresi olmadan) baslatip
#  hemen kapanir: siyah pencere yalniz ~1 sn gorunur. Sira:
#    1) sunucu (calismiyorsa ya da kodu eskiyse yeniden)  2) tepsi simgesi
#    3) uygulama penceresi (-Ilk: acilis ekrani Outlook/MiCollab hazir olana kadar doner)
#    4) Outlook / MiCollab / Teams (Ayarlar > Acilista baslat)  5) MiCollab izleyici
#  Hatalar mesaj kutusunda gosterilir.
#  NOT (15.10.2026): konsolu "headless" acan kisayol denendi; kurumsal guvenlik
#  yazilimi o kisayolu SILDI (zararli yazilim kalibi). Bu yol mevcut gizli.ps1'i kullanir.
#  Dosya ASCII'dir.
# =====================================================================
param([int]$Port = 8096)
$S = $PSScriptRoot
$kok = Split-Path -Parent $S
$log = Join-Path $env:LOCALAPPDATA ("Assist\_baslangic_" + $env:COMPUTERNAME + ".log")

function Mesaj([string]$m) {
  Add-Type -AssemblyName System.Windows.Forms
  [void][System.Windows.Forms.MessageBox]::Show($m, "Assist", "OK", "Warning")
}
function Port-Acik {
  $c = New-Object Net.Sockets.TcpClient
  try { $c.Connect("localhost", $Port); $c.Close(); return $true } catch { return $false }
}
function Gizli([string]$betik, [int]$p = 0) { & (Join-Path $S "gizli.ps1") -Betik (Join-Path $S $betik) -Port $p }

# 1) Sunucu: port dinleniyorsa bu kopyanin sunucusu mu, kodu guncel mi?
$baslat = -not (Port-Acik)
if (-not $baslat) {
  $j = $null
  try { $j = (Invoke-WebRequest -Uri "http://localhost:$Port/__saglik" -UseBasicParsing -TimeoutSec 8).Content | ConvertFrom-Json } catch {}
  if (-not $j) {
    # Cevap yok: portu bizim (takilmis) sunucumuz tutuyorsa kapatilip yeniden baslatilir
    $t = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq "powershell.exe" -and $_.ProcessId -ne $PID -and $_.CommandLine -match "-File" -and $_.CommandLine -match "sunucu_calistir\.ps1|istakip_sunucu\.ps1" })
    if (-not $t.Count) { Mesaj "$Port portunda tanimlanamayan bir uygulama var. Cozum: o uygulamayi kapatip Assist'i yeniden baslatin."; return }
    $t | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2; $baslat = $true
  }
  elseif ($j.urun -ne "istakip") { Mesaj "$Port portunda tanimlanamayan bir uygulama var. Cozum: o uygulamayi kapatip Assist'i yeniden baslatin."; return }
  elseif ($j.portalKok.TrimEnd('\') -ine $kok.TrimEnd('\')) { Mesaj ("$Port portunda BASKA bir kopyanin Assist sunucusu calisiyor:`n" + $j.portalKok + "`nO pencereyi kapatip Assist'i yeniden baslatin."); return }
  elseif ([string]$j.betikZaman -ne [string](Get-Item -LiteralPath (Join-Path $S "istakip_sunucu.ps1")).LastWriteTimeUtc.Ticks) {
    Stop-Process -Id $j.pid -Force -ErrorAction SilentlyContinue   # sunucu dosyasi guncellenmis
    Start-Sleep -Seconds 2; $baslat = $true
  }
}
if ($baslat) {
  Gizli "sunucu_calistir.ps1"
  $sw = [Diagnostics.Stopwatch]::StartNew()
  while (-not (Port-Acik) -and $sw.Elapsed.TotalSeconds -lt 30) { Start-Sleep -Milliseconds 400 }
  if (-not (Port-Acik)) {
    if (Test-Path -LiteralPath $log) { Start-Process notepad.exe -ArgumentList ('"' + $log + '"') }
    Mesaj ("Assist sunucusu acilamadi.`nHata kaydi: " + $log)
    return
  }
}

# 2) Saatin yanindaki simge (tek kopya)
Gizli "tepsi.ps1" $Port
# 3) Uygulama penceresi; yeni aciliyorsa acilis ekrani Outlook/MiCollab hazir olana kadar doner
try { & (Join-Path $S "pencere.ps1") -Port $Port -Ilk } catch {}
# 4) Outlook / MiCollab / Teams (kapali olanlar; Teams beklenmez)
try { & (Join-Path $S "uygulamalar.ps1") } catch {}
# 5) MiCollab pencere izleyicisi (tek kopya)
Gizli "micollab_izle.ps1"
