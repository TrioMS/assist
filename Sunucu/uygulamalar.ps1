# =====================================================================
#  ASSIST (IS TAKIP) - ACILISTA YANINDAKI UYGULAMALARI BASLAT
#  ---------------------------------------------------------------
#  Bat her calistiginda cagrilir; sunucu da Ayarlar'da bir uygulama acilinca
#  ayni betigi -Yalniz <ad> ile cagirir. Hangi uygulamalarin baslayacagi
#  makineye ozel ayar dosyasindadir (istakip_ayar.<PC>.json > baslat):
#    outlook  (varsayilan ACIK)  - Outlook kapaliysa simge durumunda acilir
#    micollab (varsayilan ACIK)  - MiCollab kapaliysa acilir; Emergency penceresini
#                                  ve liste ekranini micollab_izle.ps1 kapatir
#    teams    (varsayilan KAPALI)
#  Calisan uygulamaya dokunulmaz.
# =====================================================================
param([string]$Yalniz = "")
$VARSAYILAN = @{ outlook = $true; micollab = $true; teams = $false }

$secim = @{} + $VARSAYILAN
try {
  $ayar = Get-Content -LiteralPath (Join-Path $PSScriptRoot ("istakip_ayar." + $env:COMPUTERNAME + ".json")) -Raw -Encoding UTF8 | ConvertFrom-Json
  if ($ayar.baslat) { foreach ($k in @($VARSAYILAN.Keys)) { if ($null -ne $ayar.baslat.$k) { $secim[$k] = [bool]$ayar.baslat.$k } } }
} catch {}
function Istendi([string]$ad) { if ($Yalniz) { return $Yalniz -eq $ad } else { return $secim[$ad] } }

# Outlook: sunucu Outlook'a COM ile baglaninca Outlook PENCERESIZ calisiyor olabilir;
# o durumda da pencere acilir (/recycle: varsa mevcut Outlook'u kullanir).
if (Istendi "outlook") {
  try {
    $ol = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    if (-not $ol.Count -or -not ($ol | Where-Object { $_.MainWindowHandle -ne 0 })) {
      $exe = (Get-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction Stop).'(default)'
      if ($exe -and (Test-Path -LiteralPath $exe)) { Start-Process -FilePath $exe -ArgumentList "/recycle" -WindowStyle Minimized }
    }
  } catch {}
}

# MiCollab: klasor, bilgisayarda kayitli tel: isleyicisinden bulunur
if (Istendi "micollab") {
  try {
    if (-not (Get-Process -Name "UCA" -ErrorAction SilentlyContinue)) {
      $klasor = "C:\Program Files (x86)\Mitel\Unified Communicator Advanced 7.3\"
      foreach ($k in @("HKCU:\Software\Classes\tel\shell\open\command", "HKLM:\SOFTWARE\Classes\tel\shell\open\command")) {
        try { $c = [string](Get-ItemProperty -LiteralPath $k -ErrorAction Stop).'(default)'
              if ($c -match 'Mitel' -and $c -match '"?([^"]+\\)ucadialer\.exe') { $klasor = $Matches[1]; break } } catch {}
      }
      $exe = Join-Path $klasor "UCA.exe"
      if (Test-Path -LiteralPath $exe) { Start-Process -FilePath $exe -WorkingDirectory $klasor -WindowStyle Minimized }
    }
  } catch {}
}

# Teams: yeni Teams (MSTeams paketi, surec ms-teams) ya da klasik Teams (Teams.exe)
if (Istendi "teams") {
  try {
    if (-not (Get-Process -Name "ms-teams", "Teams" -ErrorAction SilentlyContinue)) {
      $paket = Get-AppxPackage -Name "MSTeams" -ErrorAction SilentlyContinue | Select-Object -First 1
      $klasik = Join-Path $env:LOCALAPPDATA "Microsoft\Teams\Update.exe"
      if ($paket) { Start-Process ("shell:AppsFolder\" + $paket.PackageFamilyName + "!MSTeams") }
      elseif (Test-Path -LiteralPath $klasik) { Start-Process -FilePath $klasik -ArgumentList '--processStart "Teams.exe"' }
    }
  } catch {}
}
