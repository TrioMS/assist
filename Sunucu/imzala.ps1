# =====================================================================
#  ASSIST - KOD IMZALAMA (YALNIZ BAKIM SORUMLUSU CALISTIRIR)
#  ---------------------------------------------------------------
#  Guvenlik incelemesi bulgu 1 (15.10.2026): paylasilan klasordeki kod
#  imzasiz calisiyordu; klasore yazabilen herkes tum kullanicilarda kod
#  calistirabiliyordu. Bu arac:
#    1) Kod imzalama sertifikasini bulur, yoksa olusturur
#       (CurrentUser\My, "CN=STFA Assist Kod Imzalama", 5 yil,
#       ozel anahtar DISA AKTARILAMAZ - yalniz bu bilgisayarda kalir)
#    2) Genel sertifikayi Sunucu\assist_imza.cer olarak yazar
#       (kullanicilar imza_kur.ps1 ile buna guvenir; ozel anahtar icermez)
#    3) Sunucu\*.ps1 dosyalarini imzalar (sunucu dosyasi EN SON)
#    4) Portal\ klasoru icin imzali katalog (Sunucu\assist.cat) uretir
#
#  ONEMLI: Her kod ya da sayfa degisikliginden SONRA yeniden calistirin.
#  Imzasi bozulan dosya, imza zorunlu kipteki kullanicilarda calismaz /
#  yuklenmez (eski surum calismaya devam eder).
#  Kullanim:  IsTakip\Assist Yayinla.bat  (ya da)
#             powershell -ExecutionPolicy Bypass -File Sunucu\imzala.ps1
#             [-ParmakIzi <kurumsal sertifika parmak izi>] [-ZamanDamgasiz] [-YeniSertifika]
#  KORUMA: Sunucu\assist_imza.cer zaten varsa ve bu bilgisayardaki sertifika
#  ondan FARKLIYSA imzalama yapilmaz (kullanicilar o parmak izine kilitli;
#  baska sertifikayla imzalamak herkesin Assist'ini durdurur). Bilerek
#  sertifika degistirmek icin -YeniSertifika verilir; sonra herkes
#  "Assist Guvenlik Kur.bat"i yeniden calistirmalidir.
#  Kurumsal (BT'nin verdigi) kod imzalama sertifikasi varsa -ParmakIzi ile
#  onu kullanin: o zaman kullanicilarda kok sertifika onayi gerekmez.
#  Dosya ASCII'dir.
# =====================================================================
param([string]$ParmakIzi = "", [switch]$ZamanDamgasiz, [switch]$YeniSertifika)
$ErrorActionPreference = "Stop"
$S = $PSScriptRoot
$kok = Split-Path -Parent $S
$KONU = "CN=STFA Assist Kod Imzalama"
$ZD = "http://timestamp.digicert.com"

# 1) Sertifika
if ($ParmakIzi) {
  $cert = Get-Item -LiteralPath ("Cert:\CurrentUser\My\" + $ParmakIzi.ToUpperInvariant())
} else {
  $cert = @(Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { $_.Subject -eq $KONU -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
    Sort-Object NotAfter -Descending) | Select-Object -First 1
  if (-not $cert -and (Test-Path -LiteralPath (Join-Path $S "assist_imza.cer")) -and -not $YeniSertifika) {
    Write-Host ""
    Write-Host "DURDURULDU: Bu bilgisayarda Assist imzalama sertifikasi yok." -ForegroundColor Red
    Write-Host "Assist'i yalniz sertifikayi olusturan bakim sorumlusunun bilgisayarindan yayinlayin." -ForegroundColor Red
    Write-Host "(Burada yeni sertifika olusturulursa imza zorunlu tum bilgisayarlarda Assist durur.)" -ForegroundColor Red
    exit 1
  }
  if (-not $cert) {
    Write-Host "Kod imzalama sertifikasi yok; olusturuluyor (5 yil, ozel anahtar disa aktarilamaz)..."
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $KONU -CertStoreLocation Cert:\CurrentUser\My `
      -NotAfter (Get-Date).AddYears(5) -KeyExportPolicy NonExportable -KeyAlgorithm RSA -KeyLength 3072 -HashAlgorithm SHA256
  }
}
if (-not $cert.HasPrivateKey) { throw "Sertifikanin ozel anahtari bu bilgisayarda yok: $($cert.Thumbprint)" }
Write-Host ("Sertifika : " + $cert.Subject + "  (" + $cert.NotAfter.ToString("dd.MM.yyyy") + " tarihine kadar)")
Write-Host ("Parmak izi: " + $cert.Thumbprint)

# 2) Genel sertifika (ozel anahtar YOK). Var olan sertifikadan farkliysa durulur (koruma, yukariya bak).
$cer = Join-Path $S "assist_imza.cer"
if ((Test-Path -LiteralPath $cer) -and -not $YeniSertifika) {
  $eski = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $cer
  if ($eski.Thumbprint -ne $cert.Thumbprint) {
    Write-Host ""
    Write-Host ("DURDURULDU: Kullanicilar " + $eski.Thumbprint + " parmak izine kilitli;") -ForegroundColor Red
    Write-Host ("bu bilgisayarin sertifikasi " + $cert.Thumbprint + ". Hicbir dosya degistirilmedi.") -ForegroundColor Red
    Write-Host "Bilerek sertifika degistirecekseniz -YeniSertifika ile calistirin." -ForegroundColor Red
    exit 1
  }
}
Export-Certificate -Cert $cert -FilePath $cer -Type CERT | Out-Null

# Zaman damgasi sunucusuna ulasilamazsa damgasiz imzalanir (imza sertifika suresi boyunca gecerli kalir)
if (-not $ZamanDamgasiz) {
  try { $null = Invoke-WebRequest -Uri $ZD -Method Head -UseBasicParsing -TimeoutSec 8 } catch {
    if (-not $_.Exception.Response) { Write-Host "Zaman damgasi sunucusuna ulasilamadi; damgasiz imzalanacak." -ForegroundColor Yellow; $ZamanDamgasiz = $true }
  }
}
function Imzala([string]$yol) {
  # Satir sonlari CRLF'ye esitlenir (16.10.2026): Set-AuthenticodeSignature eski imza blogunu silerken onunde CRLF
  # varsayip 2 karakter atiyor; blogun onunde yalniz LF varsa dosyanin son "}"si siliniyor, sunucu yeniden yuklemede
  # "Missing closing '}'" ile cokuyordu. Icerik degismez, yalniz satir sonlari.
  if ($yol -like "*.ps1") {
    $ham = [System.IO.File]::ReadAllText($yol)
    $crlf = [regex]::Replace($ham, "\r?\n", "`r`n")
    if ($crlf -ne $ham) { [System.IO.File]::WriteAllText($yol, $crlf, (New-Object System.Text.UTF8Encoding $false)) }
  }
  $p = @{ FilePath = $yol; Certificate = $cert; HashAlgorithm = "SHA256" }
  if (-not $ZamanDamgasiz) { $p.TimestampServer = $ZD }
  $s = Set-AuthenticodeSignature @p
  if (-not $s.SignerCertificate -or [string]$s.Status -eq "HashMismatch" -or [string]$s.Status -eq "NotSigned") {
    throw ("Imzalanamadi: " + $yol + " (" + $s.Status + " " + $s.StatusMessage + ")")
  }
  Write-Host ("  imzalandi: " + (Split-Path -Leaf $yol) + "  [" + $s.Status + "]")
}

# 3) Betikler: sunucu dosyasi EN SON (calisan sunucular onu gorunce yeniden yuklenir;
#    o anda guvenlik.ps1 ve katalog da imzali olmali)
$sunucu = Join-Path $S "istakip_sunucu.ps1"
$betikler = @(Get-ChildItem -LiteralPath $S -Filter "*.ps1" | Where-Object { $_.FullName -ne $sunucu } | ForEach-Object { $_.FullName })
foreach ($b in $betikler) { Imzala $b }

# 4) Sayfa katalogu (HTML Authenticode ile imzalanamaz; katalog imzalanir)
$cat = Join-Path $S "assist.cat"
if (Test-Path -LiteralPath $cat) { Remove-Item -LiteralPath $cat -Force }
New-FileCatalog -Path (Join-Path $kok "Portal") -CatalogFilePath $cat -CatalogVersion 2.0 | Out-Null
Imzala $cat

Imzala $sunucu

# Kendi bilgisayarinda dogrulama (kok guveni yoksa durum "UnknownError" gorunur: imza_kur.ps1'i bu bilgisayarda da calistirin)
$durum = @(Get-ChildItem -LiteralPath $S -Filter "*.ps1") + @(Get-Item -LiteralPath $cat) | ForEach-Object { [string](Get-AuthenticodeSignature -LiteralPath $_.FullName).Status } | Sort-Object -Unique
Write-Host ""
Write-Host ("Imza durumlari: " + ($durum -join ", "))
Write-Host "YAYINLANDI. Assist kullananlar birkac saniye icinde yeni surume gecer." -ForegroundColor Green
Write-Host ("Parmak izi (yeni kullanicilar 'Assist Guvenlik Kur.bat'ta bunu gorur): " + $cert.Thumbprint) -ForegroundColor Cyan

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDiHYk2uPbv3tX5
# zzrjvRCKqLKZllnXBbBuUgz8lYCWuaCCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEILPgau1aP2D7
# 9A/rALq3Lt/lFeSAo6hcgggB3lFR+xVUMA0GCSqGSIb3DQEBAQUABIIBgFAEJQVX
# 3hD5271K9VmIolZb0NITEr0V+xka4hKt/etoPXK5qbgX1LW4WVLuhXif7eysac6R
# WjENRpVdOz9csOAMZ7EL8ii7RApho+ckbEAfewU1VOTqqVwOGtxIbH/ssAXIaPhc
# SCRKSGorf9amDRTQddWQ6//BAiigD7ANPecRzhWcb90PHDtRHAiweE1NlDrZYZNq
# +784qug741T93podiGKi+1s0V4tuTicD1et9BtTsodFKdcT7xTeuKt3Bb2C2BO/l
# kkFXnJrQRfIo+9+YyeP6oKljRGEgesuSF3oooPfVoJVeEOv6BCciTtPinaXuDP3A
# jSD893CmRj+9w61qyPFQUjASrBjkg1rsDX5BD4njyWCN4KzAHapQp8Kbo7/DMflj
# YpxCseiYix03tIBv9e9GooOmOHv1aGXiLQ58X7XPOQcUX0Yx5f4CoPhEIAK2TOqN
# jee93bovTNMUdgq4u1eKwVXHTJL/nFGJ6ENW2Lel/fjlF3zQstr/pzUvaKGCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTYxMTI4NDJaMC8GCSqGSIb3DQEJBDEi
# BCBguChf8Y51WLXqVmPaABe8pkZa8Ii1AaB0v6I8OfO0bDANBgkqhkiG9w0BAQEF
# AASCAgArNixAsbTw1CBYgA4I2SyTGIaStKPT/+z1dhXjFeV5rN+djcfi/S7yIL6h
# qwOmGQXFz70ggsXskfG/P9ATZyf1fFY/MA9NT4UVwAeNCmgI5IOw9Fnr0ENUYpjm
# lbybfJllaXcsS3UNFrAstBZUrc9NNkgI9hHn1kYznLiarc2D98Af/Pn0IH632qd+
# tJSg9wcrKFsVRvXWSBStDg6npWOUpZuYT86a3fJZGOOMMbXEYyWplvq9IaQXMqaN
# UE9nMuNDon1duZDXXVkazn8arp3moLKOPOqmVV83P2VBzalHk5UKYvbTST1GOqnt
# 8pYriX9HZ3GmrrBZMJzsrLng7HyGo+c+tfl0wiLkmVo5hzBim+DcANkMW68qO5YV
# elgMowOzQOJekvfOlRAcxU7NjKFloXn4ggK4cakP6QwRZ2Ehh0wPOS5rp1ElUkmN
# 09jJblbDg3tFAWHMJdvpa4SqeyZvh92kapjbNWEPZx3z+4pNymkzAn5Z+iNiAOZA
# YCCv0M/RKRYWZVmzGuRlWIoOcW190SSVC+hHV6g89ZnNOkXWyaLIlzUABpw6Zkbq
# 3DytX1iY89EklQpChcsSPoT3H5mnS0/++wKuTH6S3ipZSUykKOPXbdlzUDZ6OqDo
# eg0AcWH5FJ+Ty6qpotolbIUBOg5aVt+s9/x2OnJGtg/wqIQCkw==
# SIG # End signature block
