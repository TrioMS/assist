# =====================================================================
#  IS TAKIP - SUNUCUYU KAYITLA CALISTIR
#  ---------------------------------------------------------------
#  Sunucu penceresiz calistigi icin ciktisi ve hatalari ekrana degil
#  kayit dosyasina yazilir: %LOCALAPPDATA%\Assist\_baslangic_<BILGISAYAR>.log
#  (tepsi menusu: "Sunucu kaydini goster"). Kayit UTF-8 yazilir.
# =====================================================================
$yerel = Join-Path $env:LOCALAPPDATA "Assist"; if (-not (Test-Path -LiteralPath $yerel)) { New-Item -ItemType Directory -Path $yerel -Force | Out-Null }   # yazilan dosyalar BILGISAYARDA (paylasilan klasor salt okunur olabilir)
$log = Join-Path $yerel ("_baslangic_" + $env:COMPUTERNAME + ".log")
if ($env:ISTAKIP_LOG) { $log = $env:ISTAKIP_LOG }   # test: gercek kaydin uzerine yazilmasin
"" | Out-File -LiteralPath $log -Encoding utf8
try {
  & (Join-Path $PSScriptRoot "istakip_sunucu.ps1") *>&1 | ForEach-Object { Add-Content -LiteralPath $log -Value ([string]$_) -Encoding UTF8 }
} catch {
  Add-Content -LiteralPath $log -Value ("HATA: " + $_.Exception.Message) -Encoding UTF8
  Add-Content -LiteralPath $log -Value ([string]$_.ScriptStackTrace) -Encoding UTF8
}
