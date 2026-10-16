# =====================================================================
#  ASSIST - GUVENLIK YARDIMCILARI (15.10.2026, guvenlik incelemesi)
#  ---------------------------------------------------------------
#  Sunucu, baslat.ps1, gizli.ps1, pencere.ps1 ve tepsi.ps1 bu dosyayi
#  ". guvenlik.ps1" ile yukler.
#
#  1) OTURUM ANAHTARI (bulgu 2): Yerel API artik sabit bir baslikla degil,
#     rastgele 256 bitlik bir anahtarla korunur. Anahtar yalniz bu Windows
#     kullanicisinin okuyabildigi %LOCALAPPDATA%\Assist\oturum.anahtar
#     dosyasindadir; ayni bilgisayardaki baska kullanici/surec bilemez.
#     Sayfaya pencere.ps1 adresle (?k=) verir, sayfa adresten silip
#     sessionStorage'da tutar.
#  2) SUNUCU KANITI (bulgu 2): /__saglik?nonce=... cevabinda sunucu,
#     anahtarla HMAC-SHA256 "kanit" uretir. Anahtari bilmeyen bir program
#     8096 portunu once kapsa bile kendini Assist sunucusu gibi gosteremez.
#  3) IMZA ZORUNLULUGU (bulgu 1): %LOCALAPPDATA%\Assist\guvenlik.json
#     varsa (imza_kur.ps1 yazar) betikler -ExecutionPolicy AllSigned ile
#     calisir, sunucu imzasi gecersiz kodu yeniden yuklemez, imzali
#     katalogla dogrulanamayan sayfayi sunmaz. Dosya yoksa eski davranis.
#  Dosya ASCII'dir.
# =====================================================================
$script:AssistYerel = Join-Path $env:LOCALAPPDATA "Assist"
$script:AnahtarDos  = Join-Path $script:AssistYerel "oturum.anahtar"
$script:GuvenlikDos = Join-Path $script:AssistYerel "guvenlik.json"
$script:YerelBat    = Join-Path $script:AssistYerel "Assist.bat"

function Assist-Rastgele([int]$n) {
  $b = New-Object byte[] $n
  $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
  try { $rng.GetBytes($b) } finally { $rng.Dispose() }
  return ,$b
}

# --- Oturum anahtari --------------------------------------------------
function Assist-Anahtar-Oku {
  try {
    if (Test-Path -LiteralPath $script:AnahtarDos) {
      $a = ([System.IO.File]::ReadAllText($script:AnahtarDos)).Trim()
      if ($a -match '^[A-Za-z0-9_-]{43}$') { return $a }
    }
  } catch {}
  return ""
}
# Yalniz sunucu cagirir. Var olan anahtar 7 gunden yeniyse korunur (acik pencere bozulmasin);
# yoksa ya da eskiyse yenisi uretilir.
function Assist-Anahtar-Hazirla {
  $a = Assist-Anahtar-Oku
  if ($a) {
    try { if ((Get-Item -LiteralPath $script:AnahtarDos).LastWriteTime -gt (Get-Date).AddDays(-7)) { return $a } } catch {}
  }
  if (-not (Test-Path -LiteralPath $script:AssistYerel)) { New-Item -ItemType Directory -Path $script:AssistYerel -Force | Out-Null }
  $a = [Convert]::ToBase64String((Assist-Rastgele 32)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
  [System.IO.File]::WriteAllText($script:AnahtarDos, $a, (New-Object System.Text.ASCIIEncoding))
  return $a
}
# Sabit zamanli karsilastirma (zamanlama farkindan anahtar tahmin edilemesin)
function Assist-Esit([string]$x, [string]$y) {
  if (-not $x -or -not $y -or $x.Length -ne $y.Length) { return $false }
  $d = 0
  for ($i = 0; $i -lt $x.Length; $i++) { $d = $d -bor ([int][char]$x[$i] -bxor [int][char]$y[$i]) }
  return ($d -eq 0)
}
function Assist-Hmac([string]$anahtar, [string]$mesaj) {
  $h = New-Object System.Security.Cryptography.HMACSHA256 (,[System.Text.Encoding]::UTF8.GetBytes($anahtar))
  try { return ([BitConverter]::ToString($h.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($mesaj))) -replace '-', '').ToLowerInvariant() }
  finally { $h.Dispose() }
}
function Assist-Kanit-Metni($j, [string]$nonce) {
  return ($nonce + "|" + [string]$j.urun + "|" + [string]$j.pid + "|" + [string]$j.betikZaman + "|" + [string]$j.portalKok)
}
# Porttaki sunucunun BU kullanicinin Assist sunucusu oldugunu dogrular.
# durum: "tamam" (kanit dogru) | "yabanci" (cevap var, kanit yok/yanlis) | "yok" (cevap yok)
function Assist-Sunucu-Dogrula([int]$Port, [int]$Sure = 8000) {
  $nonce = ([BitConverter]::ToString((Assist-Rastgele 16)) -replace '-', '').ToLowerInvariant()
  $j = $null
  try {
    $r = [System.Net.WebRequest]::Create("http://localhost:$Port/__saglik?nonce=$nonce")
    $r.Timeout = $Sure
    $c = $r.GetResponse()
    try { $t = (New-Object System.IO.StreamReader($c.GetResponseStream())).ReadToEnd() } finally { $c.Close() }
    $j = $t | ConvertFrom-Json
  } catch { return [pscustomobject]@{ durum = "yok"; saglik = $null } }
  $a = Assist-Anahtar-Oku
  if ($j -and $a -and [string]$j.urun -eq "istakip" -and (Assist-Esit ([string]$j.kanit) (Assist-Hmac $a (Assist-Kanit-Metni $j $nonce)))) {
    return [pscustomobject]@{ durum = "tamam"; saglik = $j }
  }
  return [pscustomobject]@{ durum = "yabanci"; saglik = $j }
}

# --- Imza zorunlulugu -------------------------------------------------
function Assist-Guvenlik-Ayar {
  try {
    if (Test-Path -LiteralPath $script:GuvenlikDos) {
      $g = [System.IO.File]::ReadAllText($script:GuvenlikDos) | ConvertFrom-Json
      $pi = ([string]$g.parmakIzi).ToUpperInvariant()
      if ($g.imzaZorunlu -and $pi -match '^[0-9A-F]{40}$') { return [pscustomobject]@{ imzaZorunlu = $true; parmakIzi = $pi } }
    }
  } catch {}
  return [pscustomobject]@{ imzaZorunlu = $false; parmakIzi = "" }
}
function Assist-Politika { if ((Assist-Guvenlik-Ayar).imzaZorunlu) { return "AllSigned" } else { return "Bypass" } }
# Imza gecerli VE imzalayan sabitlenen sertifika mi?
function Assist-Imza-Gecerli([string]$yol, [string]$parmakIzi) {
  if (-not $parmakIzi -or -not (Test-Path -LiteralPath $yol)) { return $false }
  try { $s = Get-AuthenticodeSignature -LiteralPath $yol } catch { return $false }
  if (-not $s -or -not $s.SignerCertificate) { return $false }
  return ([string]$s.Status -eq "Valid" -and $s.SignerCertificate.Thumbprint -eq $parmakIzi)
}
# Portal\ klasoru (sayfa) imzali katalogla (Sunucu\assist.cat) birebir ayni mi?
function Assist-Katalog-Gecerli([string]$portalKok, [string]$parmakIzi) {
  $cat = Join-Path $portalKok "Sunucu\assist.cat"
  if (-not (Assist-Imza-Gecerli $cat $parmakIzi)) { return $false }
  try { return ([string](Test-FileCatalog -CatalogFilePath $cat -Path (Join-Path $portalKok "Portal")) -eq "Valid") } catch { return $false }
}
# Kisayollarin hedefi: imza zorunluysa yerel baslatici (paylasilan klasordeki bat degistirilse de etkilenmez)
function Assist-Baslatici([string]$portalKok) {
  if ((Assist-Guvenlik-Ayar).imzaZorunlu -and (Test-Path -LiteralPath $script:YerelBat)) { return $script:YerelBat }
  return (Join-Path $portalKok "Assist.bat")
}

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBkmVeZzBqhm7Jt
# DAJTdrIsex/JuQJ8+OXOBDNcj2izHKCCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIKF1QHu70VnQ
# 1e936f74pyBXvEalz5+3tE75jF2tWDyZMA0GCSqGSIb3DQEBAQUABIIBgCWrd9W5
# SDVzjslOweIPAAyE8SeEf8NDEA8dAP9pdwf1/Hb6fAG7WhYc2WsWwD16zWzVmWdr
# tSBan2K+sJmXJFxEOEh6dF52B9EAHrPQUwxUZ9zlIWT375D9i4SqE7yWIvCc/Osz
# 5V7E53uBlq3SfWzVnxFZpPioWOSvbTCAmIKWVn3wGvanqBUr8c49Qp4/a/sXSUR2
# dfsG28uyHLt6b7rRfhXOxMmUgMc2mPOaL4Wssylc2qsA7LX0pn5pW3T2zlhidMG0
# wqYfJQUEjMd60UFrdsdJyz8dLW1QkJHFYxKxnX+KBz+30HfWr/4vuU6x7RMdHCqi
# lp0wh7FGecnmGfmb9zLWloAlAlbN4YbCBI94OHpZIirA+TarQaHaRB4Fp25do7h/
# Wzz4vh1PB390t/jhKWLbJDD84YhcU95Xz9HX67VpZUVYZCLEYVpN8+1Y/W9nhw+6
# kmYJr6bRki5mS8zgfVjwWK6Oy3C/NeAnoybBZ84GC4HAuQ3KphESXiMQAaGCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTYxMTI4NDJaMC8GCSqGSIb3DQEJBDEi
# BCBqUkZ9JVT+dIv26oPO4zNeWDNZMzX3Y1E/HRyJ5B5G4TANBgkqhkiG9w0BAQEF
# AASCAgCItT3yOcygby6YA9nSsuTYzKJlkURV1Wbul98zVnmhXAaJVeTDxKLxVAF1
# o/miHey5eRdDF18eItg7XfKcc0HTwGy1nrtsoEXIOQPrJCYMbc2MjQNok+zKtYUT
# gJFZO2gcZr991/MBReK1VbRNX6ObZhxiNVAmMNvfuwN1gshrGZZ0FnSpS88FRdbw
# SSa5maXpCX9Cl+y4NLM5a9JbKKi7Z7o4nDhl6BAgAGtcR6kne6QWHYxbkPNKNlra
# t1ZUSFeV15sMaV5M9tsMNxbIlQLRMOuYGp+9F4oIV0Iv3r3Qau2V1GgAHgGpBGL6
# L8O2+fC0oqfekhybFDBraDu75TEix4Cx24cDnEJ1R+m35UGfCQMYxjN7qq3WDQBD
# /ixC+0IMNAds2FMbzZkFY+8cvHo9lh/H+xl+li1yenCT+bGDs/geDEZwwb1h34Ll
# fhoyFtA0mCxZvQjGSxBnOH1+I1Cb4Xc4McKbzVJUx25r8eSq07x/r//nj0eJqK22
# 6KfyS0mae7p3iVfsT9AHs2ZcuMYB42u8uptnnFDsRMJK5cDnEjOP1HmpfhnEEoy3
# 2nnAuoJbXXdUtF8d+NXIhIE5ybjoV+AJTUC2pgWQxJOgPEqrHmUS5k3nslUE+/R/
# MoPtcnV1SBf7051/vM6XT/Rxvs2IJi2DcbUUMiR2I3bkc/Bn7g==
# SIG # End signature block
