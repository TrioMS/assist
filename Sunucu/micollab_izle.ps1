# =====================================================================
#  IS TAKIP - MICOLLAB PENCERE IZLEYICISI (penceresiz, tek kopya)
#  ---------------------------------------------------------------
#  1) MiCollab yeni acilinca cikan "Emergency Call" penceresinde "Close"a basar.
#  2) MiCollab liste penceresini ("MiCollab Client") kapatir: MiCollab'in kapat
#     dugmesi programi kapatmaz, saatin yanina gizler (gorev cubugunda da kalmaz).
#  Yalniz Mitel sureclerinin pencerelerine dokunulur.
#  AYRI SURECTIR: tepsi simgesiyle ayni is parcaciginda UI Automation calisinca
#  simge tiklamalara cevap vermiyordu (13.10.2026).
# =====================================================================
$yeni = $false
$mutex = New-Object System.Threading.Mutex($true, "Local\IsTakipMiCollabIzle", [ref]$yeni)
if (-not $yeni) { return }

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type -Namespace IsTakip -Name Pencere -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(System.IntPtr h, int k); [DllImport("user32.dll")] public static extern bool IsWindowVisible(System.IntPtr h);'
$yerel = Join-Path $env:LOCALAPPDATA "Assist"; if (-not (Test-Path -LiteralPath $yerel)) { New-Item -ItemType Directory -Path $yerel -Force | Out-Null }   # yazilan dosyalar BILGISAYARDA (paylasilan klasor salt okunur olabilir)
$log = Join-Path $yerel ("_tepsi_" + $env:COMPUTERNAME + ".log")
function Kayit([string]$m) { try { Add-Content -LiteralPath $log -Value ((Get-Date).ToString("dd.MM.yyyy HH:mm:ss") + " " + $m) } catch {} }

$MITEL = @("UCA", "53xxPCApp", "UCASoftphoneManager", "MCAC", "Mitel.PIM.ServiceHost")
$AE = [System.Windows.Automation.AutomationElement]
$DUGME = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
$YAZI = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text)

function Mitel-Pencereleri {
  $L = @()
  foreach ($p in $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)) {
    $sur = ""; try { $sur = (Get-Process -Id $p.Current.ProcessId -ErrorAction Stop).ProcessName } catch {}
    if ($MITEL -contains $sur) { $L += , @($p, $sur) }
  }
  return , $L
}

function Acil-Pencere-Kapat {
  $kapandi = $false
  try {
    foreach ($ps in (Mitel-Pencereleri)) {
      $p = $ps[0]
      $acil = $p.Current.Name -match "Emergency"
      if (-not $acil) { foreach ($y in $p.FindAll([System.Windows.Automation.TreeScope]::Descendants, $YAZI)) { if ($y.Current.Name -match "Emergency") { $acil = $true; break } } }
      if (-not $acil) { continue }
      $basildi = $false
      foreach ($d in $p.FindAll([System.Windows.Automation.TreeScope]::Descendants, $DUGME)) {
        if ($d.Current.Name -match "^\s*(Close|Kapat)\s*$") { $d.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); $basildi = $true; break }
      }
      if (-not $basildi) { try { $p.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch {} }
      Kayit ("MiCollab Emergency penceresi kapatildi (" + $(if ($basildi) { "Close" } else { "pencere" }) + ")")
      $kapandi = $true
    }
  } catch {}
  return $kapandi
}

function Liste-Gizle {
  try {
    foreach ($ps in (Mitel-Pencereleri)) {
      $p = $ps[0]
      if ($ps[1] -ne "UCA" -or $p.Current.Name -notmatch "^MiCollab") { continue }   # yalniz ana liste penceresi
      $h = [IntPtr]$p.Current.NativeWindowHandle
      if ($h -eq [IntPtr]::Zero -or -not [IsTakip.Pencere]::IsWindowVisible($h)) { continue }
      try { $p.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() }
      catch { [void][IsTakip.Pencere]::ShowWindowAsync($h, 6) }   # 6 = SW_MINIMIZE
    }
  } catch {}
}

# MiCollab acildiktan sonraki ilk 5 dk 2 sn'de bir (acil penceresi), ilk 90 sn ve
# acil penceresi kapandiktan sonraki 30 sn liste penceresi gizlenir. Diger zamanlar
# dakikada bir bakilir. Kullanici MiCollab'i sonradan kendisi acarsa dokunulmaz.
# Outlook'u saatin yanina indir: yalniz Assist'in baslattigi Outlook (uygulamalar.ps1 isaret birakir;
# kullanici Outlook'u kendisi acarsa isaret olmaz, dokunulmaz). Isaret 3 dk gecerli.
# Outlook'un "Kucultuldugunde gizle" ayari acik oldugu icin kucultulen pencere gorev cubugundan kalkar.
Add-Type -Namespace IsTakip -Name OutlookPencere -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(System.IntPtr h, int k); [DllImport("user32.dll")] public static extern bool IsWindowVisible(System.IntPtr h); [DllImport("user32.dll")] public static extern bool IsIconic(System.IntPtr h);'
$olIsaret = Join-Path $env:TEMP "assist_outlook_kucult.txt"
function Outlook-Kucult {
  if (-not (Test-Path -LiteralPath $olIsaret)) { return }
  $yas = ((Get-Date) - (Get-Item -LiteralPath $olIsaret).LastWriteTime).TotalSeconds
  if ($yas -gt 180) { Remove-Item -LiteralPath $olIsaret -Force -ErrorAction SilentlyContinue; return }
  foreach ($p in @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
    $h = $p.MainWindowHandle
    if ($h -eq [IntPtr]::Zero -or -not [IsTakip.OutlookPencere]::IsWindowVisible($h) -or [IsTakip.OutlookPencere]::IsIconic($h)) { continue }
    Start-Sleep -Seconds 2   # Outlook acilisini bitirsin (erken kucultulunce kendini geri aciyor)
    [void][IsTakip.OutlookPencere]::ShowWindowAsync($h, 6)   # 6 = SW_MINIMIZE
    Remove-Item -LiteralPath $olIsaret -Force -ErrorAction SilentlyContinue
    Kayit "Outlook saatin yanina indirildi"
    return
  }
}
$sayac = 0; $gizleBitis = [DateTime]::MinValue
while ($true) {
  Start-Sleep -Seconds 2
  $sayac++
  try { Outlook-Kucult } catch {}
  $yas = 1e9
  try { $yas = ((Get-Date) - (Get-Process -Name UCA -ErrorAction Stop | Select-Object -First 1).StartTime).TotalSeconds } catch {}
  if ($yas -lt 300 -or ($sayac % 30 -eq 0)) { if (Acil-Pencere-Kapat) { $gizleBitis = (Get-Date).AddSeconds(30) } }
  if ($yas -lt 90 -or (Get-Date) -lt $gizleBitis) { Liste-Gizle }
}

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBB+sVDJRFnBBDH
# 25Io+pWP28+ezH6GeU+iR8vLwZEiraCCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINbTWfNhDFC4
# /jDnwzJ6M0UzkzuPiLSqhflio5CQIpfsMA0GCSqGSIb3DQEBAQUABIIBgF9ohxby
# G2j+l68zsFx7h3xSc7EHt6inXuN9UENHYGorThmsw/JFcDnnnjcXZw8Rf7uX1Fwb
# t0VAGEYZ6Dk8UuXY7MzN/6OlI9Ifs9jlpXgL4EjhMJ9qTK4xKO1UNvmIGmruIKie
# B4MuPwxNaoPKVSbqsqUZqUQ7WuRz2/ltDDSIeKihvd8AAErm2hTijk/i/hef7Asu
# qXAEoMkzrt9o/jcnB2Os7CGofMjNOuxma7J+Q/4O8sup4qUeX59ljrqz3YOkdxEH
# Kf3KuAbdpNI9SwEdtJEFpLLRvgz96bK/udMuUgd/meYS8gCVIMbzRkkXKgQNbQw/
# OWBqJTsO9ZUHLKDeShTGTZ9krsq1GqahoyezqTCfob4cY+tcYQmffCePkPUQPcq0
# 3juuwEbnn7icdv/bVD5nFqKf1jFxq0gpC1CLd23XeV0Nt+YhR+ndmN4hFvyv7ij/
# 1NEYTbp7ikHnMZzRYjpnNVQpURECpQAsQtSsNHSOvZuXwVQpAljLaNqGKaGCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxMDUxMzlaMC8GCSqGSIb3DQEJBDEi
# BCDsGCDRSdAd4gsP069JSetdGCep63n3271x4gScuPsQ8zANBgkqhkiG9w0BAQEF
# AASCAgCQCMsd4QrV1ay4baptgnErWntDZLbMhp9endskaYet7bYzfNRfA8tkWb/Y
# GX/X5INIBUoT/2w7mTPeshhGrt34cdJgSspp3LphYEjBTKsAyPQmDD9TDtXrWIdI
# YsNzY/kgmJ/Kvovgg3AcjwLx57fsei0srKq5vwjQsBYc/GJm+uAFGSFYTkdb24m0
# j420PDW/EEQnfpkFv7IlTc5cPiZ9I1lWmlvCDO63wqdOoBY+MOx+TSZHu/A9/yDm
# Uh8DRW9TZn73aElMe5rKP2kwP536FZFrkUSp0zfNR94XHn3SgkpQSMXi94IqN/6u
# Km84VUZsJ2Q6QZPQsLgW3KKdFPpI9CZrZOcbUysYVQr50LDVuLYYVFYLfPo8sPcN
# 9bNCyaZQt9/MSKsa3GYqwRj7Ow2BlmiHsLC9rYKb0bQx2xSxI7PNOBh9o2OXLvPz
# YCbAC6BDNBBXkLa6lbio6Wq2XOZlnqAiCDd25VCsW88StV5W4/X+gSZHN637Qhyu
# mijNzKOcGlSTSTqtb3c3vW7xZckXgI+m3C4HTz4XR01MEyu1c9WXwscGIUioNNVH
# fWfqGoXH866KiMa3zcAI+ZKC7xR2i4DTnXrTSxftYa/iT3Mw0ZNj/NB7mKuWjG+y
# 9pXYtAU1yvAbkxqEfE1Lh7yyHkIUqeiKnTFZSwePbWRH9L5jaw==
# SIG # End signature block
