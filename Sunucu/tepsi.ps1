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
$mAc = $menu.Items.Add((U "Assist'i A{00C7}"))
$mAc.Font = New-Object System.Drawing.Font($menu.Font, [System.Drawing.FontStyle]::Bold)
$mBaslat = $menu.Items.Add((U "Sunucuyu ba{015F}lat"))
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$mKapat = $menu.Items.Add("Kapat")
$mBaslat.Visible = $false
foreach ($oge in $menu.Items) { if ($oge -is [System.Windows.Forms.ToolStripMenuItem]) { $oge.Padding = New-Object System.Windows.Forms.Padding(10, 6, 24, 6) } }
$ni.ContextMenuStrip = $menu

$mAc.add_Click({ Tarayici-Ac })
$ni.add_DoubleClick({ Kayit "cift tik"; Tarayici-Ac })
$ni.add_MouseUp({ param($s, $e) if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) { Kayit "sag tik" } })
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
$zam.add_Tick({
  $calisiyor = [bool](Saglik)
  $mBaslat.Visible = -not $calisiyor
  $ni.Text = if ($calisiyor) { $AD } else { $AD + " - sunucu durdu" }
})
$zam.Start()

Kayit ("tepsi basladi (pid " + $PID + ", STA=" + ([Threading.Thread]::CurrentThread.ApartmentState) + ")")
try { [System.Windows.Forms.Application]::Run() } catch { Kayit ("tepsi hatasi: " + $_.Exception.Message) }
Kayit "tepsi kapandi"
$ni.Dispose()
$mutex.ReleaseMutex()
