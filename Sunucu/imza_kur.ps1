# =====================================================================
#  ASSIST - IMZA ZORUNLULUGUNU BU BILGISAYARDA AC / KAPAT
#  ---------------------------------------------------------------
#  Guvenlik incelemesi bulgu 1 (15.10.2026). Her kullanici BIR KEZ calistirir
#  ("Assist Guvenlik Kur.bat"). Yaptiklari:
#    1) Sunucu\assist_imza.cer sertifikasini gosterir; parmak izi bakim
#       sorumlusunun bildirdigiyle ayni mi diye SORAR
#    2) Sertifikayi yalniz bu kullanicinin "Guvenilen yayimcilar" ve
#       "Kok" deposuna ekler (yonetici gerekmez; Windows bir onay sorar)
#    3) Tum betik imzalarini ve sayfa katalogunu dogrular - biri bile
#       gecersizse HICBIR SEY DEGISTIRMEZ
#    4) %LOCALAPPDATA%\Assist\guvenlik.json (sabitlenen parmak izi) ve
#       %LOCALAPPDATA%\Assist\Assist.bat (AllSigned yerel baslatici) yazar
#    5) Masaustu / Baslangic kisayollarini yerel baslaticiya yonlendirir
#    6) Assist'i yeniden baslatir
#  Bundan sonra imzasi bozulmus betik Windows tarafindan calistirilmaz,
#  sunucu imzasiz kodu yeniden yuklemez, dogrulanamayan sayfayi sunmaz.
#  Geri almak:  "Assist Guvenlik Kur.bat" -Kaldir
#  Dosya ASCII'dir.
# =====================================================================
param([switch]$Kaldir, [switch]$Sessiz)
$ErrorActionPreference = "Stop"
$S = $PSScriptRoot
$kok = Split-Path -Parent $S
. (Join-Path $S "guvenlik.ps1")
Add-Type -AssemblyName System.Windows.Forms
$BASLIK = "Assist guvenlik"

function Soru([string]$m) {
  if ($Sessiz) { return $true }
  return ([System.Windows.Forms.MessageBox]::Show($m, $BASLIK, "YesNo", "Question") -eq [System.Windows.Forms.DialogResult]::Yes)
}
function Bilgi([string]$m, [string]$ikon = "Information") {
  Write-Host $m
  if (-not $Sessiz) { [void][System.Windows.Forms.MessageBox]::Show($m, $BASLIK, "OK", $ikon) }
}
$ws = New-Object -ComObject WScript.Shell
# Masaustu ve Baslangic'taki Assist kisayollari (hedefi bir "Assist.bat" olanlar) verilen hedefe yonlendirilir
function Kisayollar-Yonlendir([string]$hedef) {
  foreach ($kl in @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Startup"))) {
    foreach ($l in @(Get-ChildItem -LiteralPath $kl -Filter "*.lnk" -ErrorAction SilentlyContinue)) {
      $k = $ws.CreateShortcut($l.FullName)
      if ([string]$k.TargetPath -match '\\Assist\.bat$') {
        $k.TargetPath = $hedef; $k.WorkingDirectory = $kok; $k.Save()
        Write-Host ("  kisayol: " + $l.Name + " -> " + $hedef)
      }
    }
  }
}
# Bu kullanicinin Assist surecleri (sunucu, tepsi, izleyici) kapatilip baslatici yeniden calistirilir
function Assist-Yeniden-Baslat([string]$baslatici) {
  $ben = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  foreach ($p in @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'")) {
    if ($p.ProcessId -eq $PID -or [string]$p.CommandLine -notmatch 'sunucu_calistir\.ps1|istakip_sunucu\.ps1|tepsi\.ps1|micollab_izle\.ps1') { continue }
    try { $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner; if (($o.Domain + "\" + $o.User) -ne $ben) { continue } } catch { continue }
    Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
  }
  Start-Sleep -Seconds 2
  Start-Process -FilePath $baslatici -WorkingDirectory $kok -WindowStyle Minimized
}

# ---------------------------------------------------------------- KALDIR
if ($Kaldir) {
  if (-not (Soru "Assist imza zorunlulugu bu bilgisayarda KAPATILSIN mi?`n(Betikler yeniden imza denetimi olmadan calisir.)")) { exit 1 }
  foreach ($d in @($script:GuvenlikDos, $script:YerelBat)) { if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Force } }
  $bat = Join-Path $kok "Assist.bat"
  Kisayollar-Yonlendir $bat
  Assist-Yeniden-Baslat $bat
  Bilgi "Imza zorunlulugu kapatildi. Assist yeniden baslatiliyor."
  exit 0
}

# ---------------------------------------------------------------- KUR
$cer = Join-Path $S "assist_imza.cer"
if (-not (Test-Path -LiteralPath $cer)) { Bilgi "Imza sertifikasi (Sunucu\assist_imza.cer) bulunamadi. Once bakim sorumlusu Sunucu\imzala.ps1'i calistirmali. Hicbir sey degistirilmedi." "Warning"; exit 1 }
$c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $cer
$pi = $c.Thumbprint.ToUpperInvariant()
if ($c.HasPrivateKey) { Bilgi "assist_imza.cer ozel anahtar iceriyor - bu beklenmiyor. Hicbir sey degistirilmedi; yoneticiye bildirin." "Error"; exit 1 }
$m = "Assist kodunu imzalayan sertifika:`n`n  " + $c.Subject + "`n  Parmak izi: " + $pi + "`n  Gecerlilik: " + $c.NotAfter.ToString("dd.MM.yyyy") +
     "`n`nBu parmak izi, bakim sorumlusunun size e-postayla bildirdigiyle AYNI mi?`nAyni degilse HAYIR deyin ve yoneticiye haber verin."
if (-not (Soru $m)) { Bilgi "Iptal edildi. Hicbir sey degistirilmedi."; exit 1 }

# Sertifika guveni: yalniz bu kullanici (CurrentUser). Kurumsal kokten gelen sertifikada Root adimi zaten gecerlidir.
try {
  foreach ($depo in @("Root", "TrustedPublisher")) {
    $st = New-Object System.Security.Cryptography.X509Certificates.X509Store -ArgumentList $depo, "CurrentUser"
    $st.Open("ReadWrite")
    try { if (-not $st.Certificates.Find("FindByThumbprint", $pi, $false).Count) { $st.Add($c) } } finally { $st.Close() }
  }
} catch { Bilgi ("Sertifika eklenemedi (onay reddedilmis olabilir): " + $_.Exception.Message + "`nHicbir sey degistirilmedi.") "Warning"; exit 1 }

# Imzalar: TUM betikler + sayfa katalogu gecerli olmali
$bozuk = @(Get-ChildItem -LiteralPath $S -Filter "*.ps1" | Where-Object { -not (Assist-Imza-Gecerli $_.FullName $pi) } | ForEach-Object { $_.Name })
if (-not (Assist-Katalog-Gecerli $kok $pi)) { $bozuk += "Portal (assist.cat katalogu)" }
if ($bozuk.Count) { Bilgi ("Su dosyalarin imzasi dogrulanamadi; imza zorunlulugu ACILMADI:`n`n" + ($bozuk -join "`n") + "`n`nBakim sorumlusu imzala.ps1'i yeniden calistirmali.") "Warning"; exit 1 }

# Yerel ayar + baslatici (paylasilan klasorde DEGIL: oradaki dosyalar degistirilse de etkilenmez)
if (-not (Test-Path -LiteralPath $script:AssistYerel)) { New-Item -ItemType Directory -Path $script:AssistYerel -Force | Out-Null }
$ayarMetin = (ConvertTo-Json @{ imzaZorunlu = $true; parmakIzi = $pi; kurulum = (Get-Date).ToString("s") })
[System.IO.File]::WriteAllText($script:GuvenlikDos, $ayarMetin, (New-Object System.Text.ASCIIEncoding))
# Bat dosyasi cmd'nin OEM kod sayfasiyla yazilir (yoldaki Turkce harfler icin, orn. "dosyalari")
$oem = [System.Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
if ($oem.GetString($oem.GetBytes($S)) -ne $S) {
  Remove-Item -LiteralPath $script:GuvenlikDos -Force
  Bilgi ("Klasor yolu bu bilgisayarin komut satiri kod sayfasiyla yazilamiyor:`n" + $S + "`nHicbir sey degistirilmedi; yoneticiye bildirin.") "Warning"; exit 1
}
$bat = "@echo off`r`n" +
       "rem ASSIST YEREL BASLATICI - imza_kur.ps1 olusturdu (" + (Get-Date).ToString("dd.MM.yyyy HH:mm") + "). Elle degistirmeyin.`r`n" +
       "rem Betikler AllSigned ile calisir: imzasi bozulmus betigi Windows calistirmaz.`r`n" +
       "powershell -NoProfile -ExecutionPolicy AllSigned -File """ + (Join-Path $S "gizli.ps1") + """ -Betik """ + (Join-Path $S "baslat.ps1") + """ -Port 8096`r`n"
[System.IO.File]::WriteAllBytes($script:YerelBat, $oem.GetBytes($bat))

Kisayollar-Yonlendir $script:YerelBat
Assist-Yeniden-Baslat $script:YerelBat
Bilgi ("Imza zorunlulugu ACILDI (parmak izi " + $pi.Substring(0, 8) + "...).`n`nAssist yeniden baslatiliyor. Bundan sonra Assist'i masaustundeki simgeden acin; klasordeki Assist.bat'i dogrudan calistirmayin.")

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCACr0mIEJ1xaH2h
# +O2ewDmstrxq2G52Mgzl3pCbIduL4qCCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
# gkepWCR7MrEPMA0GCSqGSIb3DQEBCwUAMCMxITAfBgNVBAMMGFNURkEgQXNzaXN0
# IEtvZCBJbXphbGFtYTAeFw0yNjEwMTUwOTIwMzRaFw0zMTEwMTUwOTMwMzNaMCMx
# ITAfBgNVBAMMGFNURkEgQXNzaXN0IEtvZCBJbXphbGFtYTCCAaIwDQYJKoZIhvcN
# AQEBBQADggGPADCCAYoCggGBALol4yzEvwnBFpLaqcZWkvILmaw2+1sgO9N1S+re
# HoUijjLl4Wq8NWQ87+9VP0ylDfCMiVZb8UBF/hvUKj9PC3hZNOTnND2+gTftgWVI
# vytdYehJgxtaPqV0i+UqtpwMrWNT6I9t7Ic2H4Dnq8CsJZoormENBah8iV6vghK9
# jyI7zOyK1Mi9EZuf4kFnit9kZGVXUqZ5DyPUzvS3nktniHC2uFJ/mM+pFJWlEGBB
# Th3qPbOnlHVv9PxTx+n101y7u7aMYNnWmXUqhv8kqst7t/P44bk0t+byPXNYHkJB
# 3i1U7qZOEt8uO4rv+GnM59YiPlyeFQTYk32RkI46EEfs0WqwbOvp4qrQgreq6Afs
# ynK7BPrDooj+IzfamKVl6lD/YkfabsgGfsa7FOz4A7egsErC1R4PEiCOXKvRNvR0
# mSkW2E1xAohxIiEWrydYf8aJVS5h4wRh4UmLfohIzsDl5LvyyNBaaCbVt4ZAGu0p
# chHeC8Dj8axLww1L0S2pYyfLNQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFFs9XAgqsDeD6t6sZCu2pNf361G+
# MA0GCSqGSIb3DQEBCwUAA4IBgQBCHRE3cTj056EKV1BLbvv0ScvX6EX05r2ZuH83
# 2d9kbPvUfa868rtwS4spEIn/3KC9aNp63BpAnojZvMQq6cC0lSQVrl77UzW5pWFg
# 0RWCZUmjLVKyj3aJlJdeIF5S7EW4McJ9WoJI23EHl90ZKR4VWYOejGG7uaT2f9jM
# APkI+4mypRnQVDr9u9jUB5mrO0vbMF7ecVLL/DK3ComtDmQU7lpGRhSXetJwdH19
# lJa1jDycnK1vqW3c28zvj6fZbP5ZYlfCJMMOOFxzXR24gq/vjc5UtOCIPNcPS6sL
# aA2X3+6BNuZyxliyBunma4WQLHj5UM77RljLYGIS6m95AANd6w6It4hZXfyH5f1q
# 7cN2PV0tcDcJGsq4D1eAhPIA+f+3huv43rk1RLGcvfYpwgefQO1Bbzx1KM5fLC5T
# kTOOGxPXTIdlM2Ql2AwgJ386mASmf6jBSvZ0NsgMHJteZFHZW8TVP0cw3rcr5grJ
# hsc4FGXpjYyqKMG1oc5W+fvLYRgwggWNMIIEdaADAgECAhAOmxiO+dAt5+/bUOII
# QBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0Rp
# Z2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0zMTEx
# MDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL/m
# kHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/zG6Q4
# FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZanMy
# lNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7sWxq8
# 68nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL2pNe
# 3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfbBHMq
# bpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3JFxG
# j2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3cAORF
# JYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqxYxhE
# lRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0vias
# tkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aLT8LW
# RV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAPBgNV
# HRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAfBgNV
# HSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYweQYI
# KwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0cDov
# L2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAR
# BgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4X6Cj
# dBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc47/T/
# gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0mcud
# T6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk1T3o
# sdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLBsln1
# VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGpn1eq
# XijiuZQwgga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEB
# CwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNV
# BAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQg
# Um9vdCBHNDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQyMzU5NTlaMGkxCzAJBgNV
# BAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNl
# cnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBD
# QTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0eDHTCphBcr48RsAc
# rHXbo0ZodLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq+RuwOnPhof6pvF4uGjwj
# qNjfEvUi6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqX
# mRinvuNgxVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOECS1UkxBvMgEdgkFiDNYi
# OTx4OtiFcMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdglTcaarps0wjUjsZvkgFk
# riK9tUKJm/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZa/zbCclF83bRVFLeGkuA
# hHiGPMvSGmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeSLsJygoLP
# p66bkDX1ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8
# VtYGqLt9MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL6s36czwz
# sucuoKs7Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rM
# kpLiQ6bGRinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFUeEY0qVjP
# KOWug/G6X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDig
# NqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz
# 8OjuhUxjaaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83afsl3YTj+IQhQE7jU/kXj
# jytJgnn0hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9rEVKChHyf
# pzee5kH0F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8
# dTXqvcnTmpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu4y81hjajV/gxdEkMx1NK
# U4uHQcKfZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3ty9qIijan
# rUR3anzEwlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4
# ZaoBdqbhOhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll38F0cuJG7uEBYTptMSbh
# dhGQDpOXgpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWs
# WCiKi9EOLLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDpMPx0LckTetiSuEtQvLsN
# z3Qbp7wGWqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8WhbaM2tszWkPZPubdcMIIG
# 7TCCBNWgAwIBAgIQCE/cM09+RU7bww+P+ZIYNTANBgkqhkiG9w0BAQsFADBpMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERp
# Z2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIw
# MjUgQ0ExMB4XDTI2MDgwNTAwMDAwMFoXDTM3MTEwNDIzNTk1OVowYzELMAkGA1UE
# BhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2Vy
# dCBTSEEyNTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjYgMTCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALZ7pvLJ/s1K+NSbTGWz/TjGMPh8
# CQ6RucZCLv5anHzWJjF/NWJrFIhy24fcpKXlgRiky4WAawDfU3YP0BMxt9l3Dm5o
# CG5Z69AqEN1kgHg2epx+l+lZBcmJCcN0ASURML5uFIS80sZsDwO3BSkUxDjLJhBI
# +qiZP3aixAC/qEGLjsBNlLol9VZ7pfGEXiMlneJIC5/YKuizVzNFKZZEeoy/0B8Z
# m+nzKBgSWG52lCO1w+nCg6XpCtklTJXeIg283hw7TmmsZXR+SMbjbrEOvZ3fP2Vx
# IgeR28Y90ZStd3F9VuA5RVynb/whITPAo9b75Zr4Ta6Mj3URm26QZYMn/FnbuTeg
# coRcFEZ9FOqM5T6MTdtr/n74lIT/ug0eeOzmZ6QTFg33otX+bFRsIolvykE1jive
# 4PuESaT8zzVeFWDAMDtozNgLctkGD1ZjkEyZtJrLl5ya0m5doH/ScpaZCZVl6pNU
# OCybMc/kxC6EAmSJY24L0yYKD1Nkddsnb/ItVKi/2nXpQNMu1PT5prW83vV8d67W
# owuUs0HdY4H8AMLGvdL/WHEj3ZnqMqAQQP9u3Ai9t+5eQ02GDwy0ODjdzi0xlp70
# W+ow63/0++YDEX1M0iwgUHwbrJvfpklkZQvw3+kv3vUPItdwroczk9icflf55W1z
# OEKAcJVAIXpcMCU9AgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQW
# BBQUyWOKMC7USvtulPPm40B+9ezN4jAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW
# 9i/USezLTjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgw
# gZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1
# Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAy
# NUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqG
# SIb3DQEBCwUAA4ICAQCNxTphHp1SCt+ZrAmAfn0oQLFr0mLywSLaDXQIENoyKqxr
# FbJblzCVP/pkXmwXOdrOpWygLzlT12os5ipDCy35RBCg2UMeApEtrfGhz45F4Wt4
# WGdNdIbRWt3YTYJmpR+b7lr4d7Uwn+H600u4D7RnOGf8Wj4UNgAdZkfHhHv1mx9E
# Vh71SJelcEN/oORSjXzdjfw1iZH9d8Nh/thn6hH23d+VsPAr6GAYyzSA02nXD1nY
# LI7Ijmiv+xLCiYC41DSFYL3GhTiy0PxpawPtGRyaBVGzq+UiTfM8pD7KVyF5aQyW
# P4KhVGUUTnmm/RlYJoW3TiXA/+t0YcT2oRVBm3JETjajHug2AL+v5jhtKVnd3D0r
# bHXEu27o+Q8p4sEWPMqKDB+qbceb6T/6WcwTwXmQ9lOCLLYcsQeSWmvKqzpAec9e
# tE14jOQAzLKWdE3w/TCaKtLRaRT7LCkRYVnhA2D73FLje1O5b3HR5eHs0NzU/+xX
# 7NbEdcofy0W3Wdwd1XOqtlpg/JgwtKfZM5dqO94lbUveOiJBI+xZEbGRsMNbXmMR
# EUTgu+Oca7Y73MPWcslIx2VhkSKSXjDbD6rgg39H5Mh7QfieAIjWagkJNt68Yfim
# 6cjEzVSiLSeZfdkr5dtFPTW6jATlWJdYeeDRGCyatf8R1hSjzSvdN8yWQPT9gzGC
# BZMwggWPAgEBMDcwIzEhMB8GA1UEAwwYU1RGQSBBc3Npc3QgS29kIEltemFsYW1h
# AhA2m8l6HI/0gkepWCR7MrEPMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcC
# AQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYB
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIL2HzW2vzwCv
# zMvzYv/BL/RH5mdo4KrCoUUaroB32iFSMA0GCSqGSIb3DQEBAQUABIIBgJvt4rsB
# Y4rSfxMC00xLXbtDyDC3RM61obq7e9ICS33//4wleI9njeqiP5hnmNYgshRBzzVL
# /BsJnO7bjX72uhc8reaFuSO+L1Ac/C0SC2kljJc5gWgtkTl37fvgJqt7rX75NTV6
# AfunUfudCvtuiHMULNsXjwX/nq30ykgtD7ifp4s+mmijGhB8/4muR+woINrXwtd6
# MSnz9epnOjbNpOhq+/VVxLgBYvnx5ZBSYSV6cP4tLjZkgwXpc380Um7tohihYkNg
# FoX71bKFkTlP/w74qzx7jIshaMgBNGlD98giHotooO4KqymHlwiKPdymZppHeQi/
# Pb1kL1FvRisNJrZx08m81CrKjQR6dxAcAD1VgJAw8BhaPyJoabOPeLkRlsgpnbFT
# Ll3EZby4O9oC4l4O+6TXsRgwN6eVr5bZxsn5iz4hP5G6c3XF3ekHMCOekciA7Z+U
# GA9IZmn04HphgsXR/8A2IB8QZpTUKdlCBDWgCny7rIG9ksB11B3H3pkbYqGCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxMDUxMzlaMC8GCSqGSIb3DQEJBDEi
# BCAL2bqcRnoccRsAb7H+1JfgjAjT6rsEHFwdQf3JsuLRUTANBgkqhkiG9w0BAQEF
# AASCAgCDJ0KbAs1tGjfd3ufBUXjBrufHnxJ8G+0SEYGPaM2ab9OZGf38JDBcyZP3
# hfjA0oN3+TuzDTfVAyP+iErWV6HHHCSjGQb9TApPMKgEqH2iJJo4J7WuhPmDbpTr
# GKndWeYeAsKD+YAS0FDqXZZ69agZ3wgkv0O/ZFIzN/ZtQMYpSMvc+zfsmNsIBbbV
# Ryrb1aqHvIcO3jo+r0RKnXjoo80ZuQjIAVKa1+/R9Sa501lx23Yvsdu6fv32nAtC
# DDQajYdLYnqKORSxkYJ26KJLJO3TzzmWlTr1nrvzELh80mckXU1H1z6gs07vCkp4
# fGrZlKBbhR5/OdbD2XTUUZfbBuA6SsDZDmMxAfYgAuwRMfb9QCkxPfjfpYw+aqoP
# 1nCMssJhmdT5P2cims7qDWmj0dIyUe50gMrEBwvKeHeM2ZrY6+2E6bJt6JQN0nPP
# n3sGOJcrL+RfsuvZZbNOqSUoKOZQ0bPgX/AJE59gbj2UUOlAGM04QgpJQx4ZB2sF
# MsC+CDllOd2cIi7Hc41cxu7fshhXZYScvsAAOa8lEG9wdrskvnlYH9BdyA40jXra
# JDfY+6kdyZoz8UFvz3Jz2O57EQ5441PH+KDDEAOpZgmg9sG/wbh3H9p5Kv9gimp1
# sLE+0rcr85R1Ukm+zZf/mpZ36kTooTHNOZK71wRDC3bnRgEFpA==
# SIG # End signature block
