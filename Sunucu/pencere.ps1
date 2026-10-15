# =====================================================================
#  ASSIST - UYGULAMA PENCERESI
#  ---------------------------------------------------------------
#  Assist'i tarayici sekmesi yerine adres cubugu olmayan, ekrani
#  kaplayan ayri bir pencerede acar (Edge, yoksa Chrome "--app").
#  Pencere zaten aciksa yenisi acilmaz: one getirilir; gidilecek yer
#  (-Hedef, ornek "ac=kayit|aksiyon") sunucuya birakilir, sayfa
#  odaklaninca alir.
#    -Hedef  : sayfa icinde gidilecek yer (# sonrasi)
#    -Url    : Windows bildiriminden gelen "assist:<hedef>" adresi
#    -Kapat  : acik Assist penceresini kapatir (sunucu calismaya devam eder)
#    -Ilk    : Assist.bat'tan acilis; yeni pencerede acilis ekrani Outlook/MiCollab hazir olana kadar kalir (/api/acilis/durum)
#  Kullananlar: Assist.bat, tepsi.ps1, sunucu (bildirim + Cikis dugmesi).
#  GUVENLIK (15.10.2026): yeni pencere acilmadan once sunucu anahtarli kanitla dogrulanir;
#  oturum anahtari sayfaya adresle (?k=) verilir (Sunucu\guvenlik.ps1). "assist:" adresinden
#  gelen hedef yalniz izin verilen karakterlerle kabul edilir.
#  Dosya ASCII'dir.
# =====================================================================
param([string]$Hedef = "", [string]$Url = "", [int]$Port = 8096, [switch]$Kapat, [switch]$Ilk, [switch]$UstKaldir)
. (Join-Path $PSScriptRoot "guvenlik.ps1")

$BASLIK = "STFA Assist"   # Portal\is_takip.html <title> ile ayni olmali (pencere bu adla bulunur)

if (-not ('AssistPencere3' -as [type])) {
  Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public static class AssistPencere3 {
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] static extern void keybd_event(byte k, byte s, uint f, UIntPtr e);
  [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr ust, int x, int y, int cx, int cy, uint f);
  // Uygulama penceresinin basligi yalniz sayfa basligidir; normal tarayici penceresinde
  // sonuna " - Microsoft Edge" vb. eklenir, o yuzden tam esitlik aranir.
  public static IntPtr Bul(string baslik) {
    IntPtr bulunan = IntPtr.Zero;
    EnumWindows(delegate (IntPtr h, IntPtr l) {
      if (!IsWindowVisible(h)) return true;
      StringBuilder c = new StringBuilder(64); GetClassName(h, c, 64);
      if (c.ToString() != "Chrome_WidgetWin_1") return true;
      StringBuilder t = new StringBuilder(256); GetWindowText(h, t, 256);
      if (t.ToString() == baslik) { bulunan = h; return false; }
      return true;
    }, IntPtr.Zero);
    return bulunan;
  }
  public static void OneGetir(IntPtr h) {
    if (IsIconic(h)) ShowWindow(h, 9);       // SW_RESTORE (simge durumundaysa)
    keybd_event(0x12, 0, 0, UIntPtr.Zero);   // Alt bas-birak: Windows arka plandaki islemin pencereyi one almasina izin verir
    keybd_event(0x12, 0, 2, UIntPtr.Zero);
    SetForegroundWindow(h);
  }
  public static void Buyut(IntPtr h) { ShowWindow(h, 3); }   // SW_MAXIMIZE
  // Ilk acilista acilis ekrani donerken pencere en ustte: arkada acilan Outlook one cikip gorunmesin (15.10.2026)
  public static void Ustte(IntPtr h, bool ac) { SetWindowPos(h, new IntPtr(ac ? -1 : -2), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0010); }   // TOPMOST / NOTOPMOST; boyut-yer-odak degismez
  public static void Kapat(IntPtr h) { PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); }   // WM_CLOSE
}
"@
}

$h = [AssistPencere3]::Bul($BASLIK)
if ($Kapat) { if ($h -ne [IntPtr]::Zero) { [AssistPencere3]::Kapat($h) }; return }
if ($UstKaldir) { if ($h -ne [IntPtr]::Zero) { [AssistPencere3]::Ustte($h, $false) }; return }   # acilis ekrani kapandi

if ($Url -match '^assist:(.*)$') { $Hedef = [Uri]::UnescapeDataString($Matches[1].TrimStart('/')) }
# Hedef yalniz sayfa ici konum ("ac=<kayit>|<aksiyon>", "hat=<id>", "kisiler" ...): baska karakter kabul edilmez
if ($Hedef -and $Hedef -notmatch '^[A-Za-z0-9=|_.-]{1,160}$') { $Hedef = "" }
$adres = "http://localhost:$Port/"
$anahtar = Assist-Anahtar-Oku
# Anahtar gonderilmeden / sayfaya verilmeden once porttaki sunucunun gercekten bu kullanicinin
# Assist sunucusu oldugu dogrulanir (taklit sunucu anahtari ele geciremesin).
$dogru = (Assist-Sunucu-Dogrula $Port 8000).durum -eq "tamam"

if ($h -ne [IntPtr]::Zero) {
  if ($Hedef -and $dogru) {
    try {
      $r = [System.Net.WebRequest]::Create($adres + "api/pencere/hedef")
      $r.Method = "POST"; $r.ContentType = "application/json"; $r.Timeout = 5000; $r.Headers.Add("X-IsTakip", "1"); $r.Headers.Add("X-Assist-Anahtar", $anahtar)
      $b = [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json @{ hedef = $Hedef } -Compress))
      $r.ContentLength = $b.Length; $s = $r.GetRequestStream(); $s.Write($b, 0, $b.Length); $s.Close(); $r.GetResponse().Close()
    } catch {}
  }
  [AssistPencere3]::OneGetir($h)
  return
}

if (-not $dogru) {
  Add-Type -AssemblyName System.Windows.Forms
  [void][System.Windows.Forms.MessageBox]::Show("$Port portundaki uygulama Assist sunucusu olarak dogrulanamadi. Assist'i masaustundeki simgeden yeniden baslatin.", "Assist", "OK", "Warning")
  return
}
$u = $adres + "?app=1" + $(if ($Ilk) { "&ilk=1" } else { "" }) + "&v=" + [DateTime]::Now.Ticks + "&k=" + $anahtar + $(if ($Hedef) { "#" + [Uri]::EscapeDataString($Hedef) } else { "" })
$tarayici = @(
  (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"),
  (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
  (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
  (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe")
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
if ($tarayici) {
  Start-Process -FilePath $tarayici -ArgumentList ("--app=" + $u), "--start-maximized"
  # Edge uygulama penceresinin son boyutunu hatirlar, --start-maximized'i yok sayabilir:
  # pencere belirince ekrani kaplayacak sekilde buyutulur (en cok 15 sn beklenir).
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 250
    $h = [AssistPencere3]::Bul($BASLIK)
    if ($h -ne [IntPtr]::Zero) { [AssistPencere3]::Buyut($h); [AssistPencere3]::OneGetir($h); if ($Ilk) { [AssistPencere3]::Ustte($h, $true) }; break }
  }
}
else { try { Start-Process $u } catch { Start-Process explorer.exe -ArgumentList $u } }   # anahtarli adres yalniz bu kullanicinin tarayicisina gider   # Edge/Chrome yoksa normal tarayici sekmesi

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBlbmNnDECWfTBd
# CbnTQcfKA83epaKM3D/pe5gIonbavqCCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIFp/Fnwn7TQ2
# NMnQA+Nhv1gEFevZAPcQ0B8fq/QJLv8QMA0GCSqGSIb3DQEBAQUABIIBgDRDBPff
# yrM1p0gmVSNwtorqAcOjHiMrhY//K34VA6zmMGD9d4ekuz/kDsWX7njGJAskY54A
# wmupr5D2vbHulTH/Fqi5b3GjCrIuIWKRRq6Kp7QQnySvPCgWCp57SqMUNVPRYu1f
# ruG+nXrl6gBMbJhamUCq/9/LTUd0ZzT1VJBvIMojjMvxdzBHvWrG0l/T7rjD1dlh
# fbmDZKmPGSqHUf3IJ/H1OW/0OQYYM8JSTOAvwt1iHeHUaQxQkTjQcvgte7Crjfq7
# 6XOZualLIJTD0cx2NDPLhBWUSsuAnZGAUnJ9FQP5oVNmMXzv9+EywOTCmQLfnd2L
# pOxwmuFTvac5S5bA0EN8gTGo+6E9XUdDNeePwWDaMYvyC8gX0nTI6d7nAkK3lw26
# lI946ocYkLGniEYM7VoQsuErwLovvkLqdO5li7OXOLg7pJoWJs4uuDd9Zh0OLkhx
# WT3C+qVOyPZc4VW2vkjPeiejQdSf9pTnsu+yvZzbtEzPLMIPGrVuZ7s1dqGCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxMjI1MDdaMC8GCSqGSIb3DQEJBDEi
# BCAZ7nbor6/eueG2zZ/CRwe/pbCH3g+HBP6Uw5vAxkcuKzANBgkqhkiG9w0BAQEF
# AASCAgCKEbeeiCBJt+Gca4QhcBhnXM31VkmRRwTJdqBSqRvrezB9FVC0SmQ9hTZP
# bd2+XWt+7ZMZ/jed5Kizk/QLCLF09dM2G9VJIcJqBFC7FeBVtFtTHF1wvnzHfKYt
# /i4OxCdmwCOu3bETe0upk1Q+jDCUNVO4n+Mib2w6C67jNMYTCCRqYVkT4eLz/KAi
# K6fzzyzLK3+tn5a1nKxD3Uy7aVvbDSQ8XjFGBgXuagKmSctDhTSsRdQzuEvstmOq
# 81Czyice9H6q/WthBIBVtqSmVXo3R0kJIczHD+VXNGS7ky1/Hsyzqi4mPJY8c7wx
# zeyFC4IG0pWpJRbSH2M60H/uXTKLxEwgIFjzF1ELleRfcRk2HI8Dqac6rEH2Aiy+
# pdhnGeQ6ETpqhRrp9QXM+Om4jLhYIy77FPozlrqRoCleMXiSwWyWJFa/1hh2vajB
# FdPoem9QdWWsf24MsXMDrS3QiVSKjDataXE8hhQZUlgwUTD0ihf01f2dLLA/BQFy
# No0FhfgsK3XcXtcUtXigXnPeprrNX/BmfbO6vt9ais/wrPlK9LCoWQg8JUThxz/z
# WXqKai2jmpGR+U9B01gKIBVE/gmR34vpg587FJoqoWBlY9eZUDa2oyzDMd16HGd1
# rDnYx+EWEGy3xcoGGFXWWxkWY7LQ2H8YaCXYJtBzYMX+9nDH4w==
# SIG # End signature block
