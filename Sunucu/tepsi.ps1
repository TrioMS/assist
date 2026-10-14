# =====================================================================
#  ASSIST (IS TAKIP) - BILDIRIM ALANI (SISTEM TEPSISI) SIMGESI
#  ---------------------------------------------------------------
#  Sunucu gizli calisir (pencere yok); bu betik saatin yanina kucuk bir
#  simge koyar:
#    cift tik      -> uygulamayi tarayicida ac
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
$mutex = New-Object System.Threading.Mutex($true, "Local\IsTakipTepsi", [ref]$yeni)
if (-not $yeni) { return }

$kok = Split-Path -Parent $PSScriptRoot
$bat = Join-Path $kok "Assist.bat"
$log = Join-Path $PSScriptRoot ("_baslangic_" + $env:COMPUTERNAME + ".log")
$adres = "http://localhost:$Port/"
$AD = "STFAssist"   # simgenin uzerine gelince gorunen ad (kullanici istegi)

# Simge: kurum logosunun bayrak isareti (tepsi_simge.png, portaldaki logodan kesildi).
# Kare simgeye en-boy orani korunarak ortalanir. Dosya yoksa kirmizi kare + onay isareti.
$bmp = New-Object System.Drawing.Bitmap 32, 32
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.Clear([System.Drawing.Color]::Transparent)
$simgeDos = Join-Path $PSScriptRoot "tepsi_simge.png"
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

function Saglik {
  try {
    $r = [System.Net.WebRequest]::Create($adres + "__saglik")
    $r.Timeout = 2000
    $c = $r.GetResponse(); $s = New-Object System.IO.StreamReader($c.GetResponseStream()); $t = $s.ReadToEnd(); $c.Close()
    return ($t | ConvertFrom-Json)
  } catch { return $null }
}
# Tani kaydi: simgeye gelen tiklamalar ve hatalar (sunucu kaydindan ayri; sunucu
# yeniden baslayinca o kayit sifirlaniyor).
$tlog = Join-Path $PSScriptRoot ("_tepsi_" + $env:COMPUTERNAME + ".log")
function Kayit([string]$m) { try { Add-Content -LiteralPath $tlog -Value ((Get-Date).ToString("dd.MM.yyyy HH:mm:ss") + " " + $m) } catch {} }
# DIKKAT: islev adi "Ac" OLAMAZ - PowerShell'de "ac" = Add-Content takma adi, islevin onune gecer;
# cift tik/menu bu yuzden hic tarayici acmiyordu (13.10.2026).
# Tarayici kapaliyken de acilsin: once varsayilan tarayici, olmazsa explorer uzerinden.
function Tarayici-Ac {
  $u = $adres + "?v=" + [DateTime]::Now.Ticks
  try { Start-Process $u; Kayit "Ac: tarayici acildi" }
  catch { Kayit ("Ac hatasi: " + $_.Exception.Message); try { Start-Process explorer.exe -ArgumentList $u } catch {} }
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
function Api([string]$yol, [bool]$post = $false) {
  try {
    $r = [System.Net.WebRequest]::Create($adres.TrimEnd('/') + $yol)
    $r.Timeout = 8000; $r.Headers.Add("X-IsTakip", "1")
    if ($post) { $r.Method = "POST"; $r.ContentType = "application/json"; $b = [Text.Encoding]::UTF8.GetBytes("{}"); $r.ContentLength = $b.Length; $s = $r.GetRequestStream(); $s.Write($b, 0, $b.Length); $s.Close() }
    $c = $r.GetResponse(); $sr = New-Object System.IO.StreamReader($c.GetResponseStream(), [Text.Encoding]::UTF8); $t = $sr.ReadToEnd(); $c.Close()
    return ($t | ConvertFrom-Json)
  } catch { return $null }
}
function Hedef-Ac([string]$hedef) {
  $u = $adres + "?v=" + [DateTime]::Now.Ticks + "#" + [Uri]::EscapeDataString($hedef)   # "ac=kayit|aksiyon" gibi
  try { Start-Process $u } catch { try { Start-Process explorer.exe -ArgumentList $u } catch {} }
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
  # Iki grup: Cevapsiz aramalar (son 48 saat) ve Gorevler (geciken/yaklasan ilk 5). Ogeye tiklayinca ilgili yer.
  $oge = { param($ust, [string]$metin, [string]$hedef, [bool]$k)
    $x = New-Object System.Windows.Forms.ToolStripMenuItem $metin
    $x.Tag = $hedef; if ($k) { $x.Font = $kalin }
    $x.Padding = New-Object System.Windows.Forms.Padding(6, 4, 18, 4)
    $x.add_Click({ param($s, $e) Hedef-Ac ([string]$s.Tag) })
    [void]$ust.DropDownItems.Add($x) }
  $kisa = { param([string]$s, [int]$n) if ($s.Length -gt $n) { $s.Substring(0, $n - 1) + [char]0x2026 } else { $s } }

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
Bildirim-Yenile

Kayit ("tepsi basladi (pid " + $PID + ", STA=" + ([Threading.Thread]::CurrentThread.ApartmentState) + ")")
try { [System.Windows.Forms.Application]::Run() } catch { Kayit ("tepsi hatasi: " + $_.Exception.Message) }
Kayit "tepsi kapandi"
$ni.Dispose()
$mutex.ReleaseMutex()
