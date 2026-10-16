# =====================================================================
#  ASSIST - CEVRIMDISI CEVIRI MOTORU (Mozilla Bergamot) KURULUM / GUNCELLEME
#  ---------------------------------------------------------------
#  16.10.2026. Posta cevirisi bilgisayarda yapilir; metin disari gitmez.
#  Bu betik YALNIZ bakim sorumlusu tarafindan, bilerek calistirilir:
#    1) Mozilla'nin Firefox'a dagittigi resmi katalogdan (Remote Settings)
#       en yeni KARARLI motoru (bergamot-translator, WebAssembly) ve secilen
#       dil ciftlerinin modellerini bulur
#    2) firefox-settings-attachments.cdn.mozilla.net'ten indirir, her dosyanin
#       SHA-256 ozetini ve boyutunu katalogla karsilastirir (tutmazsa DURUR)
#    3) Motorun JavaScript baglanti dosyasini Firefox'un ayni surume ait acik
#       kaynak kodundan alir; icindeki surum imzasi motorla eslesmezse DURUR
#    4) Portal\lib\bergamot\ altina koyar, surum.json yazar
#  Sonra Sunucu\imzala.ps1 calistirilmali (dosyalar imzali kataloga girer).
#  Kullanim:  powershell -ExecutionPolicy Bypass -File Sunucu\bergamot_indir.ps1 [-Ciftler entr,tren] [-Kontrol]
#    -Kontrol: hicbir sey indirmez; kurulu surumlerle katalogu karsilastirir.
#  Dosya ASCII'dir.
# =====================================================================
param([string[]]$Ciftler = @("entr", "tren"), [switch]$Kontrol)
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$S = $PSScriptRoot; $hedef = Join-Path (Split-Path -Parent $S) "Portal\lib\bergamot"
$KAT = "https://firefox.settings.services.mozilla.com/v1/buckets/main/collections/"
$CDN = "https://firefox-settings-attachments.cdn.mozilla.net/"
function Surum([string]$v) { try { return [version]($v -replace '[^\d.]', '') } catch { return [version]"0.0" } }

$w = @((Invoke-RestMethod ($KAT + "translations-wasm/records") -TimeoutSec 60).data) | Where-Object { $_.name -eq "bergamot-translator" -and -not $_.filter_expression }
$motor = $w | Sort-Object { Surum $_.version } -Descending | Select-Object -First 1
$m = @((Invoke-RestMethod ($KAT + "translations-models/records") -TimeoutSec 60).data) | Where-Object { -not $_.filter_expression }
$plan = [ordered]@{}
foreach ($c in $Ciftler) {
  if ($c -notmatch '^[a-z]{4}$') { throw "Gecersiz cift: $c (ornek: entr)" }
  $g = @($m | Where-Object { $_.fromLang -eq $c.Substring(0, 2) -and $_.toLang -eq $c.Substring(2, 2) })
  if (-not $g.Count) { throw "Katalogda kararli paket yok: $c" }
  $son = ($g | Sort-Object { Surum $_.version } -Descending)[0].version
  $plan[$c] = @{ surum = $son; dosyalar = @($g | Where-Object { $_.version -eq $son }) }
}

$kurulu = $null; $sj = Join-Path $hedef "surum.json"
if (Test-Path -LiteralPath $sj) { $kurulu = Get-Content -LiteralPath $sj -Raw | ConvertFrom-Json }
Write-Host ("Motor   : katalog " + $motor.version + " (" + $motor.release + ")  kurulu " + $(if ($kurulu) { $kurulu.motor.surum } else { "-" }))
foreach ($c in $plan.Keys) { Write-Host ("Cift " + $c + ": katalog " + $plan[$c].surum + "  kurulu " + $(if ($kurulu -and $kurulu.ciftler.$c) { $kurulu.ciftler.$c.surum } else { "-" })) }
if ($Kontrol) { return }

$gec = Join-Path $env:TEMP ("assist_bergamot_" + [guid]::NewGuid().ToString("N")); New-Item -ItemType Directory $gec | Out-Null
try {
  function Indir($kayit, [string]$ad) {
    $yol = Join-Path $gec $ad
    Invoke-WebRequest ($CDN + $kayit.attachment.location) -OutFile $yol -UseBasicParsing -TimeoutSec 600
    $h = (Get-FileHash -LiteralPath $yol -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($h -ne ([string]$kayit.attachment.hash).ToLowerInvariant() -or (Get-Item -LiteralPath $yol).Length -ne [long]$kayit.attachment.size) { throw ("OZET TUTMADI, durduruldu: " + $ad) }
    Write-Host ("  dogrulandi: " + $ad + "  " + [math]::Round((Get-Item -LiteralPath $yol).Length / 1MB, 1) + " MB")
    return $h
  }
  $sonuc = [ordered]@{ tarih = (Get-Date).ToString("yyyy-MM-ddTHH:mm"); kaynak = "Mozilla Remote Settings (Firefox Translations)"; lisans = "MPL-2.0"; ciftler = [ordered]@{} }
  $wh = Indir $motor "bergamot-translator.wasm"
  # Baglanti JS'i: motorun derlendigi Firefox surumunun kaynak kodundan; surum imzasi motorla eslesmeli
  $fx = ([string]$motor.fx_release -split '\.')[0]
  $jsYol = Join-Path $gec "bergamot-translator.js"
  Invoke-WebRequest ("https://raw.githubusercontent.com/mozilla-firefox/firefox/FIREFOX_" + $fx + "_0_RELEASE/toolkit/components/translations/bergamot-translator/bergamot-translator.js") -OutFile $jsYol -UseBasicParsing -TimeoutSec 120
  $beklenen = [string]$motor.release + "+" + ([string]$motor.revision).Substring(0, 9)
  if (-not (Get-Content -LiteralPath $jsYol -Raw).Contains('BERGAMOT_VERSION_FULL = "' + $beklenen)) { throw ("Baglanti JS'i motorla eslesmiyor (" + $beklenen + "), durduruldu.") }
  Write-Host ("  dogrulandi: bergamot-translator.js  (surum " + $beklenen + ")")
  $sonuc.motor = [ordered]@{ surum = [string]$motor.version; release = [string]$motor.release; revision = [string]$motor.revision; sha256 = $wh; js_sha256 = (Get-FileHash -LiteralPath $jsYol -Algorithm SHA256).Hash.ToLowerInvariant() }
  foreach ($c in $plan.Keys) {
    New-Item -ItemType Directory -Force (Join-Path $gec $c) | Out-Null
    $d = [ordered]@{}
    foreach ($k in $plan[$c].dosyalar) { $d[[string]$k.fileType] = [ordered]@{ ad = [string]$k.name; sha256 = (Indir $k ($c + "\" + $k.name)) } }
    $sonuc.ciftler[$c] = [ordered]@{ surum = $plan[$c].surum; dosyalar = $d }
  }
  # Kurulum: eski klasor yerine yenisi (yarim kalan indirme kurulu surumu bozmaz)
  if (Test-Path -LiteralPath $hedef) { Remove-Item -LiteralPath $hedef -Recurse -Force }
  New-Item -ItemType Directory -Force (Split-Path -Parent $hedef) | Out-Null
  Move-Item -LiteralPath $gec -Destination $hedef
  [System.IO.File]::WriteAllText($sj, ($sonuc | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding $false))
  Write-Host ""
  Write-Host "KURULDU: Portal\lib\bergamot. Simdi Sunucu\imzala.ps1 calistirin." -ForegroundColor Green
} finally { if (Test-Path -LiteralPath $gec) { Remove-Item -LiteralPath $gec -Recurse -Force -ErrorAction SilentlyContinue } }

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC59nVMaqiqUbFA
# ZWQLBC9zCtIvT+KXTFwAdylkIj62ZaCCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIM/R4wVM+j7d
# mZa2ITv8kww8WfF5Qq/zzn7XTtNAivCAMA0GCSqGSIb3DQEBAQUABIIBgA4Cp8FM
# czJ1QknaCH9Rj0EY+PRqWhwHIXBL4WcdSY6SPtekpS+VtmtypLcCJDTvR2QqG1Gc
# 0ch+luXdM0m6HNk8CK8c4PzWLLllknSCH87Sn/YuM5VPbfoaTMDctiAIz6QzB/pJ
# Mlap2UrtpXsmfwaxNwY62tPjFvvdctbRxlLoj+2qfN2Qgo2mnS0c5Gg+yjKuk7G5
# UvOSKm8vn6r/OTXRHsw7B0svGxzOC2aUVU2wbSKnBx7WIODkLuyKLNSesLoOkacp
# id5GmBBmEC2Gnh8Fo0aj2lbgOhoGjDBKNpFylxy4kQbtp/iSm+0PkrPK8oOMveLT
# 7k7fFyxCUA8NWV/6QuewY5/Hu6h1eyoflglGaIN99774xWk4yqtXfImpy6FoNTKB
# kN5L8koutnOW66jlm5mR1WUragQDF1ULfLSUzYyYh0d2Beg9pCnOrqNNmDmsNCsn
# sTvHnpV0mGEGntBPCNerIWZy0wwEsMsP+db+iPG0XZ3ae9XUOOVuCYNmjqGCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTYxMTI4NDFaMC8GCSqGSIb3DQEJBDEi
# BCB8J/YT0fFZ31zDnWYAw8vmVsasMpecpwSKXks8/nRJbjANBgkqhkiG9w0BAQEF
# AASCAgArqB4V/eoNBAkvfVHEiedrA9xkeGNIs06id+SH6HmqhRHiQ8MXD+YQSdE/
# Bf+KHkMJqPQgrqY93kMDXM61xtYp5z28ydkmR4vDcmxqfC9HjIO20dSH0TREwrwz
# 98YLcQmnRUa7K0UKNXsnwjX6MxaNn90u8ydaUhXPF9NpPsuTQl1++90bNyEU7oCp
# LhNtoi+1C2RbKPchEy3wqAw7y8zeJ1gqPaBiB42x5clUn+AKriz9TQ8m22ODTbLh
# nG3mbiMrj6OY0HkeSqTatas1qJwyXUsJap6xXynHFiGsuiu5d1kEDtlbv8M4WAvB
# +debFL2Z6QNCzAqxKNtVMabKT5vIugKLTTgW6O6Xi8UPwDVwppGDKfvZj0zWB0ZL
# QrRt0vbHqMWW5DXGCha0Vh3eGcn5w6ZWhjeybbYhSSqu00bUHeNicV70EmNtFhiV
# w/Z8pLi7Sb7Mkg0oV+3zvXitpP3iAn5epmZk7RYZX8qwItuRYX/IM2UG1h/0EbXO
# gFLP0wD+KkiUlncc2sycDts2CWv/oJkfdI6xEzO3gsUPCv0iaT9Lvan4KMzzkV0i
# GMPIjLscUyAaG7VPIEmB5QRKwQTSwQa28qunKjSffGyYv9KrYKM9pnWmeAGRc5yw
# rCVbucKunIofPjD0BfWndk6NkfCD1HgxUvFAmYBcyQioTalngQ==
# SIG # End signature block
