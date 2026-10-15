# =====================================================================
#  ASSIST (IS TAKIP) - BILDIRIM ALANI (SISTEM TEPSISI) SIMGESI
#  ---------------------------------------------------------------
#  Sunucu gizli calisir (pencere yok); bu betik saatin yanina kucuk bir
#  simge koyar:
#    cift tik      -> uygulamayi kendi penceresinde ac
#    sag tik menu  -> Ac / (Baslat) / Kapat
#  Sunucu durursa simge yazisi "durdu" olur, menude "Baslat" cikar.
#  Tek kopya calisir (bat her calistiginda yenisi acilmaz).
#  Dosya ASCII'dir; Turkce harfler {XXXX} kodu ile yazilir.
# =====================================================================
param([int]$Port = 8096)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function U([string]$s) { return [regex]::Replace($s, '\{([0-9A-F]{4})\}', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) }) }

# Tek kopya
$yeni = $false
# Mutex adi 15.10.2026 guvenlik guncellemesinde degisti: anahtarsiz eski tepsi sureci (eski mutex) yenisini engellemez,
# yeni tepsi acilinca bu kullanicinin eski tepsi surecini kapatir (eski surec API'den artik 403 aliyordu).
$mutex = New-Object System.Threading.Mutex($true, "Local\IsTakipTepsi2", [ref]$yeni)
if (-not $yeni) { return }
try {
  $benKim = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  foreach ($p in @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'")) {
    if ($p.ProcessId -eq $PID -or [string]$p.CommandLine -notmatch 'tepsi\.ps1') { continue }
    $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner
    if (($o.Domain + "\" + $o.User) -eq $benKim) { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue }
  }
} catch {}

$kok = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "guvenlik.ps1")   # oturum anahtari + sunucu kaniti (15.10.2026)
$bat = Assist-Baslatici $kok   # imza zorunluysa yerel baslatici
$yerel = Join-Path $env:LOCALAPPDATA "Assist"; if (-not (Test-Path -LiteralPath $yerel)) { New-Item -ItemType Directory -Path $yerel -Force | Out-Null }   # yazilan dosyalar BILGISAYARDA (paylasilan klasor salt okunur olabilir)
$log = Join-Path $yerel ("_baslangic_" + $env:COMPUTERNAME + ".log")
$adres = "http://localhost:$Port/"
$AD = "STFAssist"   # simgenin uzerine gelince gorunen ad (kullanici istegi)

# Simge: Assist logosu (assist_simge.png, Assist.svg'den uretildi; 15.10.2026).
# Kare simgeye en-boy orani korunarak ortalanir. Dosya yoksa kirmizi kare + onay isareti.
$bmp = New-Object System.Drawing.Bitmap 32, 32
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.Clear([System.Drawing.Color]::Transparent)
$simgeDos = Join-Path $PSScriptRoot "assist_simge.png"
if (Test-Path -LiteralPath $simgeDos) {
  $isaret = [System.Drawing.Image]::FromFile($simgeDos)
  $h = [int][math]::Round(32 * $isaret.Height / $isaret.Width)
  $g.DrawImage($isaret, (New-Object System.Drawing.Rectangle(0, [int]((32 - $h) / 2), 32, $h)))
  $isaret.Dispose()
} else {
  $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(216, 35, 42))), 2, 2, 28, 28)
  $kalem = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 4
  $kalem.StartCap = [System.Drawing.Drawing2D.LineCap]::Round; $kalem.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
  $g.DrawLines($kalem, [System.Drawing.Point[]]@((New-Object System.Drawing.Point 9, 16), (New-Object System.Drawing.Point 14, 21), (New-Object System.Drawing.Point 23, 11)))
}
$g.Dispose()
$ikon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())

# Yeni bildirim varken simgenin sag alt kosesine kirmizi yuvarlak + sayi cizilir (0 ise temel simge).
Add-Type -Namespace IsTakip -Name TepsiYerel -MemberDefinition '[DllImport("user32.dll")] public static extern bool DestroyIcon(System.IntPtr h);'
$script:sonIkonSay = -1; $script:sayiliHandle = [IntPtr]::Zero
function Simge-Say([int]$n) {
  if ($n -eq $script:sonIkonSay) { return }
  $script:sonIkonSay = $n
  $eski = $script:sayiliHandle; $script:sayiliHandle = [IntPtr]::Zero
  if ($n -le 0) { $ni.Icon = $ikon }
  else {
    $b2 = New-Object System.Drawing.Bitmap $bmp
    $g2 = [System.Drawing.Graphics]::FromImage($b2)
    $g2.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g2.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g2.FillEllipse((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(216, 35, 42))), 13, 13, 19, 19)
    $g2.DrawEllipse((New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 1.5), 13, 13, 19, 19)
    $yazi = if ($n -gt 9) { "9+" } else { [string]$n }
    $f = New-Object System.Drawing.Font("Segoe UI", $(if ($n -gt 9) { 8 } else { 11 }), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object System.Drawing.StringFormat; $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $g2.DrawString($yazi, $f, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(13, 13.5, 19, 19)), $sf)
    $g2.Dispose()
    $script:sayiliHandle = $b2.GetHicon(); $b2.Dispose()
    $ni.Icon = [System.Drawing.Icon]::FromHandle($script:sayiliHandle)
  }
  if ($eski -ne [IntPtr]::Zero) { [void][IsTakip.TepsiYerel]::DestroyIcon($eski) }
}

$ni = New-Object System.Windows.Forms.NotifyIcon
$ni.Icon = $ikon
$ni.Text = $AD
$ni.Visible = $true

# Yalniz anahtarli kanitla dogrulanan sunucu "calisiyor" sayilir (taklit sunucunun PID'i kapatilmaz).
function Saglik {
  $d = Assist-Sunucu-Dogrula $Port 2000
  if ($d.durum -eq "tamam") { return $d.saglik } else { return $null }
}
# Tani kaydi: simgeye gelen tiklamalar ve hatalar (sunucu kaydindan ayri; sunucu
# yeniden baslayinca o kayit sifirlaniyor).
$tlog = Join-Path $yerel ("_tepsi_" + $env:COMPUTERNAME + ".log")
function Kayit([string]$m) { try { Add-Content -LiteralPath $tlog -Value ((Get-Date).ToString("dd.MM.yyyy HH:mm:ss") + " " + $m) } catch {} }
# DIKKAT: islev adi "Ac" OLAMAZ - PowerShell'de "ac" = Add-Content takma adi, islevin onune gecer;
# cift tik/menu bu yuzden hic tarayici acmiyordu (13.10.2026).
# Tarayici kapaliyken de acilsin: once varsayilan tarayici, olmazsa explorer uzerinden.
# Assist uygulama penceresinde acilir (Sunucu\pencere.ps1: adres cubugu yok; aciksa one gelir).
$pencere = Join-Path $PSScriptRoot "pencere.ps1"
function Tarayici-Ac {
  try { & $pencere -Port $Port; Kayit "Ac: uygulama penceresi" }
  catch {
    Kayit ("Ac hatasi: " + $_.Exception.Message)
    $u = $adres + "?v=" + [DateTime]::Now.Ticks
    try { Start-Process $u } catch { try { Start-Process explorer.exe -ArgumentList $u } catch {} }
  }
}

$menu = New-Object System.Windows.Forms.ContextMenuStrip
# Varsayilan 9 punto menu cok kucuk kaliyordu (kullanici, 13.10.2026): 12 punto + genis satir
$menu.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$menu.ShowImageMargin = $false
# En ustte "Bildirimler": Windows bildirimleri kaybolsa da burada son 48 saatinkiler ve cevapsiz
# aramalar durur; ogeye tiklayinca Assist ilgili yerde acilir. Menu acilinca "goruldu" sayilir.
$mBld = New-Object System.Windows.Forms.ToolStripMenuItem (U "Bildirimler")
[void]$menu.Items.Add($mBld)
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$mAc = $menu.Items.Add((U "Assist'i A{00C7}"))
$mAc.Font = New-Object System.Drawing.Font($menu.Font, [System.Drawing.FontStyle]::Bold)
$mBaslat = $menu.Items.Add((U "Sunucuyu ba{015F}lat"))
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$mKapat = $menu.Items.Add("Kapat")
$mBaslat.Visible = $false
foreach ($oge in $menu.Items) { if ($oge -is [System.Windows.Forms.ToolStripMenuItem]) { $oge.Padding = New-Object System.Windows.Forms.Padding(10, 6, 24, 6) } }
$ni.ContextMenuStrip = $menu

$mAc.add_Click({ Tarayici-Ac })

# Sunucu API'si (X-IsTakip basligi: sunucu yalniz bu basligi tasiyan yerel istegi kabul eder)
function Api([string]$yol, [bool]$post = $false, [string]$govde = "{}") {
  try {
    $r = [System.Net.WebRequest]::Create($adres.TrimEnd('/') + $yol)
    $r.Timeout = 8000; $r.Headers.Add("X-IsTakip", "1"); $r.Headers.Add("X-Assist-Anahtar", (Assist-Anahtar-Oku))
    if ($post) { $r.Method = "POST"; $r.ContentType = "application/json"; $b = [Text.Encoding]::UTF8.GetBytes($govde); $r.ContentLength = $b.Length; $s = $r.GetRequestStream(); $s.Write($b, 0, $b.Length); $s.Close() }
    $c = $r.GetResponse(); $sr = New-Object System.IO.StreamReader($c.GetResponseStream(), [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $c.Close()
    return ($t | ConvertFrom-Json)
  } catch { return $null }
}
function Hedef-Ac([string]$hedef) {
  try { & $pencere -Port $Port -Hedef $hedef }   # "ac=kayit|aksiyon" gibi
  catch {
    $u = $adres + "?v=" + [DateTime]::Now.Ticks + "#" + [Uri]::EscapeDataString($hedef)
    try { Start-Process $u } catch { try { Start-Process explorer.exe -ArgumentList $u } catch {} }
  }
}
$kalin = New-Object System.Drawing.Font($menu.Font, [System.Drawing.FontStyle]::Bold)
function Bildirim-Yenile {
  $j = Api "/api/tepsi/ozet"
  $mBld.DropDownItems.Clear()
  if (-not $j) { Simge-Say 0; $mBld.Text = U "Bildirimler"; [void]$mBld.DropDownItems.Add((U "Sunucuya ula{015F}{0131}lam{0131}yor")); return }
  $n = [int]$j.yeni
  Simge-Say $n
  $mBld.Text = if ($n) { (U "Bildirimler") + " (" + $n + " yeni)" } else { U "Bildirimler" }
  $mBld.Font = if ($n) { $kalin } else { $menu.Font }
  $ni.Text = if ($n) { $AD + " - " + $n + " yeni bildirim" } else { $AD }
  # Uc grup: E-postalar (okunmamis), Cevapsiz aramalar (son 48 saat), Gorevler (geciken/yaklasan ilk 5). Ogeye tiklayinca ilgili yer.
  $oge = { param($ust, [string]$metin, [string]$hedef, [bool]$k)
    $x = New-Object System.Windows.Forms.ToolStripMenuItem $metin
    $x.Tag = $hedef; if ($k) { $x.Font = $kalin }
    $x.Padding = New-Object System.Windows.Forms.Padding(6, 4, 18, 4)
    $x.add_Click({ param($s, $e) Hedef-Ac ([string]$s.Tag) })
    [void]$ust.DropDownItems.Add($x) }
  $kisa = { param([string]$s, [int]$n) if ($s.Length -gt $n) { $s.Substring(0, $n - 1) + [char]0x2026 } else { $s } }

  # E-postalar: Gelen Kutusu'nda okunmamislar (son 48 saat, en yeni 8); tiklayinca Outlook'ta acilir.
  $po = @($j.postalar | ForEach-Object { $_ }); $poYeni = @($po | Where-Object { $_.yeni }).Count; $poSay = [int]$j.postaSay
  $mPo = New-Object System.Windows.Forms.ToolStripMenuItem ("E-postalar" + $(if ($poYeni) { " (" + $poYeni + " yeni)" } elseif ($poSay) { " (" + $poSay + (U " okunmam{0131}{015F})") } else { "" }))
  if ($poYeni) { $mPo.Font = $kalin }
  [void]$mBld.DropDownItems.Add($mPo)
  if (-not $po.Count) { $x = $mPo.DropDownItems.Add((U "Son 48 saatte okunmam{0131}{015F} posta yok")); $x.Enabled = $false }
  foreach ($o in $po) {
    $x = New-Object System.Windows.Forms.ToolStripMenuItem ([string]$o.saat + "   " + (& $kisa ([string]$o.konu) 64))   # gonderen yok (Outlook guvenlik uyarisi)
    $x.Tag = (ConvertTo-Json @{ id = [string]$o.id; store = [string]$o.store } -Compress); if ($o.yeni) { $x.Font = $kalin }
    $x.Padding = New-Object System.Windows.Forms.Padding(6, 4, 18, 4)
    $x.add_Click({ param($s, $e) [void](Api "/api/posta/ac" $true ([string]$s.Tag)) })
    [void]$mPo.DropDownItems.Add($x)
  }
  [void]$mPo.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)); & $oge $mPo (U "Posta Kutusu'nda a{00E7}") "posta" $false

  $cev = @($j.cevapsiz | ForEach-Object { $_ }); $cevYeni = @($cev | Where-Object { $_.yeni }).Count
  $mCev = New-Object System.Windows.Forms.ToolStripMenuItem ((U "Cevaps{0131}z aramalar") + $(if ($cevYeni) { " (" + $cevYeni + " yeni)" } elseif ($cev.Count) { " (" + $cev.Count + ")" } else { "" }))
  if ($cevYeni) { $mCev.Font = $kalin }
  [void]$mBld.DropDownItems.Add($mCev)
  if (-not $cev.Count) { $x = $mCev.DropDownItems.Add((U "Son 48 saatte cevaps{0131}z arama yok")); $x.Enabled = $false }
  foreach ($o in $cev) { & $oge $mCev ([string]$o.saat + "   " + ([string]$o.metin -replace '^.*\u00B7\s*', '')) "kisiler" ([bool]$o.yeni) }
  if ($cev.Count) { [void]$mCev.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)); & $oge $mCev (U "Ki{015F}iler'de a{00E7} (geri ara)") "kisiler" $false }

  $gv = @($j.gorevler | ForEach-Object { $_ }); $gSay = [int]$j.gorevSay
  $mGor = New-Object System.Windows.Forms.ToolStripMenuItem ((U "G{00F6}revler") + $(if ($gSay) { " (" + $gSay + ")" } else { "" }))
  [void]$mBld.DropDownItems.Add($mGor)
  if (-not $gv.Count) { $x = $mGor.DropDownItems.Add((U "Geciken ya da yakla{015F}an g{00F6}rev yok")); $x.Enabled = $false }
  foreach ($o in $gv) { & $oge $mGor ([string]$o.kalan + "   " + [string]$o.no + "  " + (& $kisa ([string]$o.baslik) 48)) ([string]$o.hedef) ([bool]$o.gecikti) }
  if ($gSay) { [void]$mGor.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)); & $oge $mGor ((U "T{00FC}m{00FC}n{00FC} g{00F6}r") + " (" + $gSay + ")") "ana" $false }
}
# Liste acilinca yeni olanlar goruldu sayilir (simgedeki sayi silinir); liste yine durur.
$mBld.add_DropDownOpened({ if ($script:sonIkonSay -gt 0) { [void](Api "/api/tepsi/goruldu" $true); Simge-Say 0; $ni.Text = $AD; $mBld.Text = U "Bildirimler"; $mBld.Font = $menu.Font } })
$menu.add_Opening({ Bildirim-Yenile })

# Tek sol tik: menu (bildirimler en ustte) acilir; cift tik: Assist acilir. Tek tik, cift tik
# suresi kadar beklenip cift tik gelmezse islenir (iki islem birden olmasin).
$tekTik = New-Object System.Windows.Forms.Timer
$tekTik.Interval = [System.Windows.Forms.SystemInformation]::DoubleClickTime
$tekTik.add_Tick({
  $tekTik.Stop()
  try { $ni.GetType().GetMethod("ShowContextMenu", [System.Reflection.BindingFlags]"NonPublic,Instance").Invoke($ni, $null) } catch { Kayit ("menu hatasi: " + $_.Exception.Message) }
})
$script:ciftTik = [datetime]::MinValue
$ni.add_DoubleClick({ $tekTik.Stop(); $script:ciftTik = Get-Date; Kayit "cift tik"; Tarayici-Ac })
$ni.add_MouseUp({ param($s, $e)
  if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) { Kayit "sag tik" }
  elseif ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and ((Get-Date) - $script:ciftTik).TotalMilliseconds -gt 600) { $tekTik.Stop(); $tekTik.Start() }
})
$menu.add_Opening({ Kayit "menu aciliyor" })
$mBaslat.add_Click({ Start-Process -FilePath $bat -WindowStyle Minimized })
$mKapat.add_Click({
  $c = [System.Windows.Forms.MessageBox]::Show(
    (U "Assist kapat{0131}ls{0131}n m{0131}?`n`nKapat{0131}l{0131}rsa hat{0131}rlatma ve bildirim gelmez."),
    $AD, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
  if ($c -ne [System.Windows.Forms.DialogResult]::Yes) { return }
  $j = Saglik
  if ($j -and $j.pid) { Stop-Process -Id $j.pid -Force -ErrorAction SilentlyContinue }
  $ni.Visible = $false
  [System.Windows.Forms.Application]::Exit()
})

# MiCollab pencere isleri AYRI surectedir (Sunucu\micollab_izle.ps1): UI Automation bu
# simgeyle ayni is parcaciginda calisinca simge tiklamalara cevap vermiyordu (13.10.2026).

# Sunucuyu 15 sn'de bir yokla: duruma gore simge yazisi ve "Baslat" secenegi
$zam = New-Object System.Windows.Forms.Timer
$zam.Interval = 15000
$script:tikSay = 0
$zam.add_Tick({
  $calisiyor = [bool](Saglik)
  $mBaslat.Visible = -not $calisiyor
  if (-not $calisiyor) { $ni.Text = $AD + " - sunucu durdu"; Simge-Say 0 }
  elseif (($script:tikSay++ % 2) -eq 0 -and -not $menu.Visible) { Bildirim-Yenile }   # bildirimler 30 sn'de bir
})
$zam.Start()
# Yeni posta aninda: 5 sn'de bir Gelen Kutusu surumu yoklanir, degisince bildirimler hemen tazelenir
$script:postaRev = -1
$poz = New-Object System.Windows.Forms.Timer
$poz.Interval = 5000
$poz.add_Tick({
  $r = Api "/api/posta/rev"
  if ($r -and $script:postaRev -ge 0 -and [int]$r.rev -ne $script:postaRev -and -not $menu.Visible) { Bildirim-Yenile }
  if ($r) { $script:postaRev = [int]$r.rev }
})
$poz.Start()
Bildirim-Yenile

Kayit ("tepsi basladi (pid " + $PID + ", STA=" + ([Threading.Thread]::CurrentThread.ApartmentState) + ")")
try { [System.Windows.Forms.Application]::Run() } catch { Kayit ("tepsi hatasi: " + $_.Exception.Message) }
Kayit "tepsi kapandi"
$ni.Dispose()
$mutex.ReleaseMutex()

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCD8AfeI5eMwsVq
# o9mFzZx5O3WVowy5cLd/SVl7oRA9pqCCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIC563lFkrddu
# +L7FLKSvfhQW7/dJ+FHgePOrPlSYSPmyMA0GCSqGSIb3DQEBAQUABIIBgA2VoRyU
# s9H4Uc41QCOA6Kmp2n2EAW3YC3jWG23ZDlum98notPP9oFqnECS4vZaVFB4oTPH1
# /kpszcVTPU0Em4RjCiHJ3RiBMq5b+wbKPpcw7arwWZ44tYKL7cWJQysxoQ+NymZh
# Lysj9FxGHGz4hYlVrtY+m3isx8cGLEDwFmBk0+tRd3+liiC+cQzDyWVqHEgi/JxL
# mTCIbak7YJhwd+lfsEW9A4UTUAD1+5iFdwQ9FKRFke88J2vg+U5bvhaE0/REccvY
# KKUDo9l8MSKBD+GRVFG2nwi8fdQL1FtxA5sK+LpQ3PaEP2oRsS0T/Icr4TOrxRFr
# OG2ahZ1utyFgkeMZz7lqdEt8T4vx/UkyDDJMfrdtrQFDM18yLxcAZICQpu8+jk4n
# qdtJsWnXnxMGZsrdYpIf3XuRCKooxhfJ/Ju8v16Ti37cmGrlWrR7VO+cVgFZgK3p
# X552TC3s3R9JgFr4QOK5izlV5KFeOr3TLmFQIGuB67j9D9OfFhZSbe26/6GCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxMjI1MDhaMC8GCSqGSIb3DQEJBDEi
# BCAWou8jTDrfAbiYKllIz1LTZgh0IVoTdhUVmfoQGHQgXTANBgkqhkiG9w0BAQEF
# AASCAgCgYz+MzRfMzQSNmxhcM3fYu3ZdyX8ethzVWGRP+o5uML/TVyzd6Aru+tDC
# b7xCslC80fe7pG8ZDJgBHz2TrXY98xAeIfWmU49jUkzh0OvKu/VllOYDbwrH7Fbe
# YYxtn05ZgJ46iwtx+jiBgHvhipbCi5tJzuoLrzqpW7eqYR3YIF3o3I4Z1eqGt2Pc
# 3X6BXxd8K+tg73boT/IqQIB1+6xD3i41WhXnkSLP4QLZKTGHy7i20xoQv8hkZaTk
# l6rYOvh9FId5P3yH4dNDR72AGjKCa0cX7+9QzuHPjM7m02dan6pmHKkgunSfOdrm
# Zzqjz53/X9Y/rcLS9ZpHZGLASP7q4KZ3gbKU5WlhhnP4AIzoB42L6ticQ1iZ1v15
# +6WjUWLanl9OZCFiUoQ7K72lXJeRHDf0gcqdd2QhJ3fAZbXFY3waq/8HVHFQUUFE
# 01Mph4UUwScKLX86h1jmhAuUfxHLk5pJ6+bUS4B7CXsQ70CG25qQM6u+mljLNScl
# YohiQslgGdObAy2/906tLsUBTEJkbyHXmpC/VPcUXbn9lNCaOzNpZVA81c6WHHAm
# cTzpDA3JoPzVpVUksxxv11p5yQo6jHvbtWJgBIHTfB2LEkHZ+De0LNC/j7Xsi8vW
# 2qOdsM5tvVV0tocbKtWuG/rHBy/NhZlZY6OHadyvvb48DcwnPA==
# SIG # End signature block
