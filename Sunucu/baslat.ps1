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
#  GUVENLIK (15.10.2026): porttaki sunucu artik anahtarli kanitla (HMAC) dogrulanir;
#  imza zorunlu kipte imzasi gecersiz sunucu dosyasi baslatilmaz (Sunucu\guvenlik.ps1).
#  Dosya ASCII'dir.
# =====================================================================
param([int]$Port = 8096)
$S = $PSScriptRoot
. (Join-Path $S "guvenlik.ps1")
$GUV = Assist-Guvenlik-Ayar
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
  # Kanit dogrulanamazsa (cevap yok, eski surum ya da taklit) sunucu bizim degil sayilir.
  $d = Assist-Sunucu-Dogrula $Port 8000
  $j = if ($d.durum -eq "tamam") { $d.saglik } else { $null }
  if (-not $j) {
    # Dogrulanamadi: portu bizim (takilmis / eski surum) sunucumuz tutuyorsa kapatilip yeniden baslatilir
    $t = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq "powershell.exe" -and $_.ProcessId -ne $PID -and $_.CommandLine -match "-File" -and $_.CommandLine -match "sunucu_calistir\.ps1|istakip_sunucu\.ps1" })
    if (-not $t.Count) { Mesaj "$Port portunda tanimlanamayan bir uygulama var. Cozum: o uygulamayi kapatip Assist'i yeniden baslatin."; return }
    $t | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2; $baslat = $true
  }
  elseif ($j.urun -ne "istakip") { Mesaj "$Port portunda tanimlanamayan bir uygulama var. Cozum: o uygulamayi kapatip Assist'i yeniden baslatin."; return }
  elseif ($j.portalKok.TrimEnd('\') -ine $kok.TrimEnd('\')) { Mesaj ("$Port portunda BASKA bir kopyanin Assist sunucusu calisiyor:`n" + $j.portalKok + "`nO pencereyi kapatip Assist'i yeniden baslatin."); return }
  elseif ([string]$j.betikZaman -ne [string](Get-Item -LiteralPath (Join-Path $S "istakip_sunucu.ps1")).LastWriteTimeUtc.Ticks) {
    if ($GUV.imzaZorunlu -and -not (Assist-Imza-Gecerli (Join-Path $S "istakip_sunucu.ps1") $GUV.parmakIzi)) {
      # Yeni sunucu dosyasi imzasiz/bozuk: calisan (dogrulanmis) surume dokunulmaz
      Mesaj "Assist sunucu dosyasi degismis ama imzasi gecersiz. Calisan surum kullanilmaya devam ediyor; yoneticiye bildirin."
    } else {
      Stop-Process -Id $j.pid -Force -ErrorAction SilentlyContinue   # sunucu dosyasi guncellenmis
      Start-Sleep -Seconds 2; $baslat = $true
    }
  }
}
if ($baslat) {
  if ($GUV.imzaZorunlu) {
    $bozuk = @("istakip_sunucu.ps1", "sunucu_calistir.ps1", "guvenlik.ps1", "gizli.ps1") | Where-Object { -not (Assist-Imza-Gecerli (Join-Path $S $_) $GUV.parmakIzi) }
    if ($bozuk) { Mesaj ("Assist baslatilmadi: su dosyalarin imzasi gecersiz:`n" + ($bozuk -join ", ") + "`nYoneticiye bildirin.") ; return }
  }
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

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxqEjHFdgB4rIz
# fzfq4vHDgEAj737yrzEAxR7WkTBez6CCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIEml8KOwxCck
# 9ERskCx9VVcJUCCeeZUx64JVfZfs2TIRMA0GCSqGSIb3DQEBAQUABIIBgKzlUekE
# iPcn+QsOnlxfFbOYkftvrNJl5Uum8GH4j8lHgNVmLSddkYSSsZN2DoojhP/Z2yKE
# mwavDyaB77VBqRBB7QlL9vtcY4uR9049j+voyIvfTkN370qXjLIhe7Uibj/BeXBk
# oOpoB9aG4cHRMGmszfuN0MrosCnSRwiL1k+EsC+iSdNgTuFQdCwSnSU3oE8e/vxA
# xFmU35kUYoLEodZw5oHk3oRMq80FQR7MwtRkBSddrNsN46mTOWILN7phIExg9Xhs
# SfAMnKMh9G5qaRcxBoElUYLUjVmliA3Qe3CxS1ND/0+rj9HM0NJvVt8ULn+nETMD
# wpgrOAu8ozdbAby/ZFXZ1gmc7WdVu7Lpg8pIo90vdYwNFTcm2xoM5a3mxGfMQT1o
# epXXhpP6MKeostb0JQMFRI9Ts6SWdvSiNrWxsuoxD6cGdQdWVnK22oYKq3sedrhw
# iiy5HPe/iXRQjrU/57huIiyHmed4uBGgb1sx7Nm+hfhgIi2PUgHv7EGshaGCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTYxMTI4NDFaMC8GCSqGSIb3DQEJBDEi
# BCBsSHHB2KCM7oGTT2r14T6ijtVtOR5FK19sf1liSmC94zANBgkqhkiG9w0BAQEF
# AASCAgCg2ys7LwvChUBde/Qw3SbXOc8hg4mjm6WWgmBUQgRMhNK+MNX/KMyjTep6
# hJFmxL/0GfOKxdHavz+n12W3vGgYgPJ95fr5Nmcx4jm0hWARMEIESV84JVUXMPxD
# duC2V7yva6OpXGCDlZ3sjQ/sBycpGIqPKdxCP7gdQZJu1udOL00FjDi4o8KMyT91
# vnbN7SuqqTFlUq8j+/hHX+Kn0pukglLeeNkRZbrUHULjmPxI5b86B9+/0Ipnfj4z
# MWrFMOpTHilcJb25AT41V/5S+J2wTmwQnHPviwzqWdgL1mmMykanZDUHLgNUEbzd
# SDlTnIUrwA/Bav52c1viyeRawGZKqV5Z4vF8aOXenx6vxf67wehQYiI/u3CbMFlE
# 9xTkNCpMZWnuQ15WKQt5wamqy0oVCTegzaQrFu7XJUMqcZWLWC1zHgcQRKyDalDB
# LCe3v25z8rMDGGZYSIIfClnX4K9jtbsigsP5EFe339AHPT7fVIQxyRFEGLOwpnl1
# M8DkIYgtmwzgWPSBUpe+nm9CB98fvfMW69UcT9oiHde5a8HzweLJ/jqHbQOdlmND
# 7FTmEyVBiH90wDZ+TNDavE8NXIBsH//rgcjTK4MBt0H0Y5nx5sMcUz9HrUT48tOp
# dUuZaOFRp7Y+YYaN0+PR/W6DiRPkgP1vCD6fd7h0rakKID5g+w==
# SIG # End signature block
