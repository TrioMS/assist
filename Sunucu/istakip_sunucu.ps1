# =====================================================================
#  ASSIST (IS TAKIP) - YEREL SUNUCU
#  ---------------------------------------------------------------
#  Sifir bagimlilik: System.Net.HttpListener + Outlook COM. Node/npm
#  kurulumu gerektirmez (Hukuk portali sunucusuyla ayni kalip).
#
#  Iki kok:
#    portalKok : uygulamanin KODU (bu klasor, buluta senkron olabilir)
#    veriKok   : KAYITLAR (posta ozetleri, .msg kopyalari) - varsayilan
#                kullanicinin profilinde, OneDrive DISI.
#
#  OUTLOOK BAGLANTISI: connector yok; bu bilgisayarda acik olan Outlook
#  masaustu uygulamasina COM ile baglanilir. Posta sunucudan disari
#  CIKMAZ - yalniz bu makinedeki tarayiciya verilir.
#
#  Dinleme yalnizca localhost'tur. /api uclari ozel bir baslik ister
#  (X-IsTakip); baska bir web sayfasi tarayici uzerinden bu uclara
#  istek atamasin diye.
#
#  NOT: Bu dosya bilincli olarak ASCII'dir (PowerShell 5.1 BOM'suz UTF-8
#  dosyayi yanlis okur). Turkce metinler portal tarafinda durur.
# =====================================================================
$ErrorActionPreference = "Stop"
$SURUM = "0.1.0"

$betikKok = $PSScriptRoot
$BETIK_ZAMAN = (Get-Item -LiteralPath $PSCommandPath).LastWriteTimeUtc.Ticks

# --- Ayar ------------------------------------------------------------
# Ayar dosyasi MAKINEYE OZELDIR: klasor paylasilirsa herkesin kendi
# dosyasi olur, kimse otekinin veri yolunu ezmez.
$yerel = Join-Path $env:LOCALAPPDATA "Assist"; if (-not (Test-Path -LiteralPath $yerel)) { New-Item -ItemType Directory -Path $yerel -Force | Out-Null }   # yazilan dosyalar BILGISAYARDA (paylasilan klasor salt okunur olabilir)
$ayarDos = Join-Path $yerel ("istakip_ayar." + $env:COMPUTERNAME + ".json")
$eskiAyar = Join-Path $betikKok ("istakip_ayar." + $env:COMPUTERNAME + ".json")   # 14.10.2026 oncesi yeri: bir kez tasinir
if (-not (Test-Path -LiteralPath $ayarDos) -and (Test-Path -LiteralPath $eskiAyar)) { try { Copy-Item -LiteralPath $eskiAyar -Destination $ayarDos } catch {} }
$varsayilan = [ordered]@{
  port    = 8096
  veriKok = (Join-Path $env:USERPROFILE "IsTakipVeri")
}
if (Test-Path -LiteralPath $ayarDos) {
  $ayar = Get-Content -LiteralPath $ayarDos -Raw -Encoding UTF8 | ConvertFrom-Json
} else {
  $ayar = New-Object psobject -Property $varsayilan
  ($ayar | ConvertTo-Json) | Out-File -LiteralPath $ayarDos -Encoding utf8
}
foreach ($k in $varsayilan.Keys) {
  if (-not ($ayar.PSObject.Properties.Name -contains $k) -or -not [string]$ayar.$k) {
    $ayar | Add-Member -NotePropertyName $k -NotePropertyValue $varsayilan[$k] -Force
  }
}
$port      = [int]$ayar.port
$portalKok = (Resolve-Path (Join-Path $betikKok "..")).Path
$webKok    = Join-Path $portalKok "Portal"
$veriKok   = [string]$ayar.veriKok
# Gelistirme/test: gercek kayitlara dokunmadan ayri port ve klasorle calistirmak icin.
if ($env:ISTAKIP_PORT) { $port = [int]$env:ISTAKIP_PORT }
if ($env:ISTAKIP_VERI) { $veriKok = $env:ISTAKIP_VERI }
foreach ($d in @($veriKok, (Join-Path $veriKok "posta"), (Join-Path $veriKok "yedek"), (Join-Path $veriKok "belge"))) {
  if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d | Out-Null }
}
$veriDos = Join-Path $veriKok "isler.dat"      # SIFRELI (asagiya bak)
$eskiDos = Join-Path $veriKok "isler.json"     # 13.10.2026 oncesi duz metin; acilista sifrelenir
$revDos  = Join-Path $veriKok "rev.txt"
$postaKok = Join-Path $veriKok "posta"
$belgeKok = Join-Path $veriKok "belge"         # hatirlaticiya baglanan belgeler (sifreli)
# Acilmasina izin verilen belge turleri: calistirilabilir dosya (exe, bat, ps1, js...) ASLA acilmaz.
$BELGE_UZ = @(".pdf",".doc",".docx",".xls",".xlsx",".ppt",".pptx",".txt",".rtf",".csv",".msg",".eml",
              ".jpg",".jpeg",".png",".gif",".bmp",".tif",".tiff",".udf",".odt",".ods",".zip")
$acKok   = Join-Path $env:TEMP "IsTakipAc"     # acilmak icin cozulen posta kopyalari (gecici)
$script:rev = 0
if (Test-Path -LiteralPath $revDos) { try { $script:rev = [int](Get-Content -LiteralPath $revDos -Raw).Trim() } catch { $script:rev = 0 } }

$UTF8 = New-Object System.Text.UTF8Encoding($false)

# --- Sifreleme (KVKK, 13.10.2026) ------------------------------------
# Kayitlar, yedekler ve posta kopyalari Windows DPAPI ile KULLANICI
# HESABINA bagli sifrelenir: dosya kopyalansa bile yalniz ayni Windows
# oturumunda acilir. Parola yok, kullanici hicbir sey yapmaz.
# UYARI: Windows hesabi silinir / profil sifirlanirsa kayitlar acilamaz.
# Kurumsal (etki alani) hesapta anahtar etki alaninda yedeklidir.
Add-Type -AssemblyName System.Security
$ENTROPI = $UTF8.GetBytes("IsTakip.v1")
function Sifrele([byte[]]$b) { return ,[System.Security.Cryptography.ProtectedData]::Protect($b, $ENTROPI, [System.Security.Cryptography.DataProtectionScope]::CurrentUser) }
function Coz([byte[]]$b) { return ,[System.Security.Cryptography.ProtectedData]::Unprotect($b, $ENTROPI, [System.Security.Cryptography.DataProtectionScope]::CurrentUser) }
function Bayt-Yaz([string]$yol, [byte[]]$b) {
  # Once gecici dosyaya, sonra yerine: yazma yarida kesilirse eski kayit bozulmaz.
  $tmp = $yol + ".tmp"
  [System.IO.File]::WriteAllBytes($tmp, $b)
  if (Test-Path -LiteralPath $yol) { [System.IO.File]::Replace($tmp, $yol, [NullString]::Value) } else { [System.IO.File]::Move($tmp, $yol) }
}
function Sifreli-Oku([string]$yol) { return $UTF8.GetString((Coz ([System.IO.File]::ReadAllBytes($yol)))) }
function Sifreli-Yaz([string]$yol, [string]$metin) { Bayt-Yaz $yol (Sifrele ($UTF8.GetBytes($metin))) }
function Veri-Oku { if (Test-Path -LiteralPath $veriDos) { return (Sifreli-Oku $veriDos) } else { return "" } }

# Duz metin dosyalari bir kez sifrelenir; sifreli kopya dogrulanmadan duz olan SILINMEZ.
function Sifre-Gocu {
  $n = 0
  $ciftler = @()
  if (Test-Path -LiteralPath $eskiDos) { $ciftler += ,@($eskiDos, $veriDos) }
  Get-ChildItem -LiteralPath (Join-Path $veriKok "yedek") -Filter "isler-*.json" -ErrorAction SilentlyContinue |
    ForEach-Object { $ciftler += ,@($_.FullName, ($_.FullName -replace '\.json$', '.dat')) }
  foreach ($c in $ciftler) {
    $duz = [System.IO.File]::ReadAllText($c[0], $UTF8)
    if (-not (Test-Path -LiteralPath $c[1])) { Sifreli-Yaz $c[1] $duz }
    if ((Sifreli-Oku $c[1]) -eq $duz) { Remove-Item -LiteralPath $c[0] -Force; $n++ }
    elseif ($c[0] -eq $eskiDos) {
      # Ana kayitta tutarsizlik: hicbir sey silinmez, sunucu ACILMAZ (bos veriyle acilip
      # uzerine yazilmasin). Iki dosya da yerinde kalir; elle bakilmali.
      throw "isler.json ve isler.dat ikisi de var ve farkli. Hicbir sey silinmedi; veri klasorune bakin: $veriKok"
    }
  }
  Get-ChildItem -LiteralPath $postaKok -Filter "*.msg" -ErrorAction SilentlyContinue | ForEach-Object {
    $ham = [System.IO.File]::ReadAllBytes($_.FullName)
    $hedef = $_.FullName + ".dat"
    Bayt-Yaz $hedef (Sifrele $ham)
    if ((Coz ([System.IO.File]::ReadAllBytes($hedef))).Length -eq $ham.Length) { Remove-Item -LiteralPath $_.FullName -Force; $n++ }
  }
  return $n
}
# Acmak icin cozulen posta kopyalari 1 saat sonra silinir (Outlook aciksa kilitli; sonra denenir).
function Gecici-Temizle {
  if (-not (Test-Path -LiteralPath $acKok)) { return }
  Get-ChildItem -LiteralPath $acKok -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-1) } |
    ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch {} }
}
# Hicbir kayda bagli olmayan posta kopyasi silinir (kayit silinince / saklama suresi dolunca).
# 1 saatten yeni dosyaya dokunulmaz: posta baglanip kayit henuz yazilmamis olabilir.
function Yetim-Posta-Temizle([string]$metin) {
  $kullanilan = @{}
  foreach ($m in [regex]::Matches($metin, '"msg"\s*:\s*"([0-9a-f]{32}\.msg)"')) { $kullanilan[$m.Groups[1].Value] = $true }
  Get-ChildItem -LiteralPath $postaKok -Filter "*.msg.dat" -ErrorAction SilentlyContinue |
    Where-Object { -not $kullanilan[$_.Name.Substring(0, $_.Name.Length - 4)] -and $_.LastWriteTime -lt (Get-Date).AddHours(-1) } |
    ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
  # Hatirlaticidan cikarilan / hatirlaticisi silinen belgeler
  $belgeler = @{}
  foreach ($m in [regex]::Matches($metin, '"belgeId"\s*:\s*"([0-9a-f]{32})"')) { $belgeler[$m.Groups[1].Value + ".dat"] = $true }
  Get-ChildItem -LiteralPath $belgeKok -Filter "*.dat" -ErrorAction SilentlyContinue |
    Where-Object { -not $belgeler[$_.Name] -and $_.LastWriteTime -lt (Get-Date).AddHours(-1) } |
    ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
}

# --- Hatirlatici zamanlari -------------------------------------------
# baslangic "yyyy-MM-ddTHH:mm"; tekrar: yok | gunluk | isgunu | haftalik (gunler: 0=Pazar..6) | aylik.
# Sayfadaki hatSon/hatSonraki ile AYNI kural (biri degisirse oburu de).
function Hat-GunUygun($h, [datetime]$d, [datetime]$bas) {
  if ($d.Date -lt $bas.Date) { return $false }
  if ($h.bitis -and $d.ToString("yyyy-MM-dd") -gt [string]$h.bitis) { return $false }
  switch ([string]$h.tekrar) {
    "gunluk"   { return $true }
    "isgunu"   { $g = [int]$d.DayOfWeek; return ($g -ge 1 -and $g -le 5) }
    "haftalik" { return (@($h.gunler | ForEach-Object { [int]$_ }) -contains [int]$d.DayOfWeek) }
    "aylik"    { return ($d.Day -eq [math]::Min($bas.Day, [DateTime]::DaysInMonth($d.Year, $d.Month))) }
    default    { return ($d.Date -eq $bas.Date) }
  }
}
# Simdiye kadarki son olusum (yoksa $null)
function Hat-Son($h) {
  try { $bas = [datetime]::ParseExact([string]$h.baslangic, "yyyy-MM-ddTHH:mm", $null) } catch { return $null }
  $simdi = Get-Date
  for ($i = 0; $i -le 400; $i++) {
    $d = $simdi.Date.AddDays(-$i)
    if ($d -lt $bas.Date) { break }
    if (Hat-GunUygun $h $d $bas) { $o = $d.Add($bas.TimeOfDay); if ($o -le $simdi) { return $o } }
  }
  return $null
}

# --- Yardimcilar -----------------------------------------------------
function Yaz-Cevap($ctx, [string]$metin, [string]$tur, [int]$kod) {
  if (-not $kod) { $kod = 200 }
  $b = $UTF8.GetBytes($metin)
  $ctx.Response.StatusCode = $kod
  $ctx.Response.ContentType = $tur
  $ctx.Response.Headers.Add("Cache-Control", "no-store")
  $ctx.Response.ContentLength64 = $b.Length
  $ctx.Response.OutputStream.Write($b, 0, $b.Length)
}
function Yaz-Json($ctx, $nesne, [int]$kod) {
  if (-not $kod) { $kod = 200 }
  Yaz-Cevap $ctx (ConvertTo-Json -InputObject $nesne -Depth 6 -Compress) "application/json; charset=utf-8" $kod
}
function Yaz-Hata($ctx, [string]$mesaj, [int]$kod) {
  if (-not $kod) { $kod = 500 }
  Yaz-Json $ctx @{ ok = $false; hata = $mesaj } $kod
}
function Govde-Oku($ctx) {
  $sr = New-Object System.IO.StreamReader($ctx.Request.InputStream, $UTF8)
  $s = $sr.ReadToEnd(); $sr.Close(); return $s
}
# Sorgu parametresi UTF-8 olarak okunur (.NET QueryString Turkce harfleri bozar).
function Sorgu($ctx, [string]$ad) {
  $q = $ctx.Request.Url.Query
  if (-not $q) { return "" }
  foreach ($p in $q.TrimStart('?').Split('&')) {
    $i = $p.IndexOf('=')
    if ($i -lt 0) { continue }
    if ($p.Substring(0, $i) -eq $ad) { return [Uri]::UnescapeDataString($p.Substring($i + 1).Replace('+', ' ')) }
  }
  return ""
}
# Gunluk yedek: gunun ILK kaydindan once mevcut (sifreli) dosya kopyalanir, 30 gun tutulur.
function Yedek-Al {
  if (-not (Test-Path -LiteralPath $veriDos)) { return }
  $yd = Join-Path (Join-Path $veriKok "yedek") ("isler-" + (Get-Date -Format "yyyy-MM-dd") + ".dat")
  if (-not (Test-Path -LiteralPath $yd)) {
    Copy-Item -LiteralPath $veriDos -Destination $yd
    Get-ChildItem -LiteralPath (Join-Path $veriKok "yedek") -Filter "isler-*.dat" |
      Sort-Object Name -Descending | Select-Object -Skip 30 | Remove-Item -Force -ErrorAction SilentlyContinue
  }
}
function Icerik-Turu([string]$uz) {
  switch ($uz.ToLowerInvariant()) {
    ".html" { "text/html; charset=utf-8" } ".js" { "text/javascript; charset=utf-8" }
    ".css" { "text/css; charset=utf-8" } ".json" { "application/json; charset=utf-8" }
    ".png" { "image/png" } ".svg" { "image/svg+xml" } ".ico" { "image/x-icon" }
    default { "application/octet-stream" }
  }
}

# --- Outlook ---------------------------------------------------------
$PR_MSGID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$script:ol = $null
$script:cevAnahtar = ""; $script:cevListe = @()   # MiCollab cevapsiz arama onbellegi (/api/sayaclar)
$script:logAnahtar = ""; $script:logListe = @()
# MiCollab canli kaydindan (uca.log) cevapsiz aramalar: arama "CallReceived" ile gelir; cagri
# "CallCleared" olurken durumu hala "Ringing" ise (hic acilmadan bitti) cevapsizdir.
# Kayit MiCollab'da acik oldugu icin paylasimli okunur; dosya degismedikce yeniden okunmaz.
function Log-Cevapsiz {
  $l = Join-Path $env:APPDATA "Mitel\UC\uca.log"
  if (-not (Test-Path -LiteralPath $l)) { return }
  $fi = Get-Item -LiteralPath $l
  $z = [string]$fi.Length + "|" + $fi.LastWriteTimeUtc.Ticks
  if ($script:logAnahtar -ne $z) {
    $metin = ""
    try {
      $fs = [System.IO.File]::Open($l, 'Open', 'Read', 'ReadWrite')
      $sr = New-Object System.IO.StreamReader($fs); $metin = $sr.ReadToEnd(); $sr.Close(); $fs.Close()
    } catch { return $script:logListe }
    $gelen = @{}; $sonSid = ""; $sonTur = ""; $L = @()
    foreach ($s in $metin.Split("`n")) {
      if ($s -notmatch 'CallListener') { continue }
      if ($s -match 'Call Event: EventCallID \(\d+\) ID \(\S+\) SessionID \((\S+)\)') { $sonSid = $Matches[1]; continue }
      if ($s -match '^(\d{4}/\d\d/\d\d \d\d:\d\d:\d\d)\.\d+ .*Call Event: EventType \((\w+)\) .*Number \(([^)]*)\)') {
        $sonTur = $Matches[2]
        if ($sonTur -eq 'CallReceived' -and $sonSid) { $gelen[$sonSid] = [string]$Matches[3] }
        continue
      }
      if ($sonTur -eq 'CallCleared' -and $s -match '^(\d{4}/\d\d/\d\d \d\d:\d\d:\d\d)\.\d+ .*FindExistingCall: ExistingCall: .*SessionID=(\S+) .*Status=Ringing') {
        $sid = $Matches[2]
        if ($gelen.ContainsKey($sid)) {
          $t = [DateTime]::ParseExact($Matches[1], 'yyyy/MM/dd HH:mm:ss', $null).ToUniversalTime()
          $L += @{ ad = ""; no = $gelen[$sid]; t = $t }
          $gelen.Remove($sid)
        }
        $sonTur = ""
      }
    }
    $script:logListe = $L; $script:logAnahtar = $z
  }
  return $script:logListe
}
# $sinir (UTC) sonrasindaki cevapsiz aramalar, eskiden yeniye; iki kaynak (gecmis dosyasi + canli kayit)
# birlestirilir, ayni numaradan 2 dk icindeki kayit tek sayilir. Oge: @{ ad; no; t (UTC DateTime) }
function Cevapsiz-Listesi([datetime]$sinir) {
  try {
    $f = Get-ChildItem -Path (Join-Path $env:APPDATA "Mitel\UC\*\callhistory.dat") -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($f) {
      $z = $f.FullName + "|" + $f.LastWriteTimeUtc.Ticks
      if ($script:cevAnahtar -ne $z) {
        # PS 5.1: ConvertFrom-Json diziyi TEK nesne olarak verir; once degiskene alinir, sonra acilir
        $ham = Get-Content -Raw -LiteralPath $f.FullName -Encoding UTF8 | ConvertFrom-Json
        $script:cevListe = @($ham | ForEach-Object { $_ } | Where-Object { [int]$_.Type -eq 2 } |
          ForEach-Object { @{ ad = [string]$_.Name; no = [string]$_.Number; t = [DateTime]::Parse([string]$_.Date).ToUniversalTime() } })
        $script:cevAnahtar = $z
      }
    }
  } catch {}
  # MiCollab gecmis dosyasini gecikmeli/eksik yaziyor (14.10.2026: arama kayitta yoktu);
  # canli kayit (uca.log) cagri biter bitmez yazilir -> ikisi birlestirilir.
  $hepsi = @($script:cevListe) + @(Log-Cevapsiz)
  $L = @()
  foreach ($c in ($hepsi | Where-Object { $_ } | Sort-Object { $_.t })) {
    if ($c.t -le $sinir) { continue }
    $ayni = $false
    foreach ($v in $L) { if ($v.no -eq $c.no -and [Math]::Abs(($v.t - $c.t).TotalSeconds) -lt 120) { if (-not $v.ad) { $v.ad = $c.ad }; $ayni = $true; break } }
    if (-not $ayni) { $L += @{ ad = $c.ad; no = $c.no; t = $c.t } }
  }
  return $L
}
function Outlook {
  if ($script:ol) { try { $null = $script:ol.Version; return $script:ol } catch { $script:ol = $null } }
  $script:ol = New-Object -ComObject Outlook.Application
  return $script:ol
}
function Ns { return (Outlook).GetNamespace("MAPI") }
# Gelen Kutusu ya da altindaki bir klasor ("Hukuk" veya "Hukuk/Talepler").
function Klasor([string]$yol) {
  $f = (Ns).GetDefaultFolder(6)
  if ($yol) {
    foreach ($parca in $yol.Split('/', '\')) {
      if (-not $parca.Trim()) { continue }
      $f = $f.Folders.Item($parca.Trim())
    }
  }
  return $f
}
# --- Assist'ten Assist'e talep (Exchange uzerinden, 13.10.2026) -------------
# Gonderen Assist, alicinin adresine ozel siniftaki (IPM.Note.Assist) bir ileti yollar:
# govdede okunabilir metin + "ASSIST-VERI:<base64 JSON>:SON". Alicinin Assist'i dakikada
# bir Gelen Kutusu'nda bu siniftaki iletileri bulur, Gelen Kutusu > Assist klasorune
# tasir ve talebi sifreli kuyruga (gelen.dat) yazar; sayfa kuyrugu alip "Gelen talep"
# kaydi olusturur. Alicida Assist yoksa ileti normal posta gibi okunur.
# Sinif sirket disina giden postada IPM.Note'a doner: disaridan sahte talep dusmez.
$GELEN_SINIF = "IPM.Note.Assist"
$VERI_ISARET = "ASSIST-VERI:"
$gelenDos = Join-Path $veriKok "gelen.dat"
# Ogeleri tek tek dondurur (virgul ile SARMA: borudan tek nesne gibi gecip $_.id dizi olur);
# cagiran her zaman @(Gelen-Oku) ile toplar.
function Gelen-Oku {
  if (Test-Path -LiteralPath $gelenDos) {
    try { $j = Sifreli-Oku $gelenDos; if ($j -and $j.Trim()) { $a = $j | ConvertFrom-Json; foreach ($x in $a) { $x } } } catch {}
  }
}
function Gelen-Yaz($L) { $d = @($L); Sifreli-Yaz $gelenDos $(if ($d.Count) { ConvertTo-Json -InputObject $d -Depth 8 -Compress } else { '[]' }) }
function Assist-Klasoru {
  $gk = (Ns).GetDefaultFolder(6)
  try { return $gk.Folders.Item("Assist") } catch { return $gk.Folders.Add("Assist") }
}
# Mobil Assist (Power Apps) icin gorevler: kullanicinin OneDrive'inda Assist\assist_mobil.json.
# Power Apps OneDrive baglantisiyla okuyup ParseJSON ile acar (yalniz okuma). Veri degismedikce
# (rev) ve gun donmedikce yeniden yazilmaz. Test sunucusunda (ISTAKIP_PORT) yazilmaz.
$script:mobilAnahtar = ""
$MOBIL_GELEN = if ($env:ISTAKIP_PORT) { "Assist\mobil_gelen_test" } else { "Assist\mobil_gelen" }   # test sunucusu gercek istekleri almasin
function Mobil-Disa-Aktar {
  if ($env:ISTAKIP_PORT -or -not $env:OneDriveCommercial -or -not (Test-Path -LiteralPath $veriDos)) { return }
  $an = [string]$script:rev + "|" + (Get-Date).ToString("yyyy-MM-dd")
  if ($an -eq $script:mobilAnahtar) { return }
  $V = (Veri-Oku) | ConvertFrom-Json
  if (-not $V) { return }
  $L = New-Object System.Collections.ArrayList
  $ekle = { param($tur, $no, $baslik, $ust, $termin, $sorumlu, $kid, $aid)
    $g = Gun-Fark $termin
    [void]$L.Add([ordered]@{ tur = $tur; no = [string]$no; baslik = [string]$baslik; ust = [string]$ust; termin = [string]$termin
      kalan = $(if ($null -eq $g) { 9999 } else { $g }); sorumlu = [string]$sorumlu; kid = [string]$kid; aid = [string]$aid }) }
  foreach ($k in @($V.kayitlar | ForEach-Object { $_ })) {
    $acik = ($k.durum -eq "acik" -or $k.durum -eq "beklemede")
    if ($acik -and $k.tur -ne "posta") { & $ekle $(if ($k.tur -eq "not") { "Not" } else { "Talep" }) $k.no $k.baslik "" $k.termin $k.sorumlu $k.id "" }
    if ($k.durum -eq "iptal") { continue }
    foreach ($ak in @($k.aksiyonlar | ForEach-Object { $_ })) {
      if ($ak.durum -ne "tamam") { & $ekle "Aksiyon" $ak.no $ak.metin $k.baslik $ak.termin $ak.sorumlu $k.id $ak.id }
    }
  }
  $sirali = @($L | Sort-Object { $_.kalan }, { $_.no })
  $yg = [int]$V.ayar.yaklasanGun; if ($yg -le 0) { $yg = 7 }
  $cikti = [ordered]@{ surum = 1; guncelleme = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss"); ad = [string]$V.ayar.adSoyad
    yaklasanGun = $yg; telOnek = $(if ($null -ne $V.ayar.telOnek) { [string]$V.ayar.telOnek } else { "1" }); gorevler = $sirali }
  $kl = Join-Path $env:OneDriveCommercial "Assist"
  if (-not (Test-Path -LiteralPath $kl)) { New-Item -ItemType Directory -Path $kl | Out-Null }
  $klG = Join-Path $env:OneDriveCommercial $MOBIL_GELEN; if (-not (Test-Path -LiteralPath $klG)) { New-Item -ItemType Directory -Path $klG | Out-Null }   # telefonun biraktigi istekler
  $metin = ConvertTo-Json -InputObject $cikti -Depth 5 -Compress
  [System.IO.File]::WriteAllText((Join-Path $kl "assist_mobil.json"), $metin, $UTF8)
  # Ayni icerik .txt olarak da: Power Automate "dosya icerigini al" duz metni dogrudan verir (json'u base64 sarar)
  [System.IO.File]::WriteAllText((Join-Path $kl "assist_mobil.txt"), $metin, $UTF8)
  $script:mobilAnahtar = $an
}
# Bu Windows oturumunun kurumsal e-postasi (kucuk harf, onbellekli): once etki alani, olmazsa Outlook.
$script:benEposta = $null
function Ben-Eposta {
  if ($null -ne $script:benEposta) { return $script:benEposta }
  $e = ""
  try { Add-Type -AssemblyName System.DirectoryServices.AccountManagement; $e = [string][System.DirectoryServices.AccountManagement.UserPrincipal]::Current.EmailAddress } catch {}
  if (-not $e) { try { $e = [string](Ns).CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {} }
  $script:benEposta = $e.Trim().ToLowerInvariant()
  return $script:benEposta
}


function Gelen-Talep-Kontrol {
  $gk = (Ns).GetDefaultFolder(6)
  $bulunan = @(); foreach ($it in $gk.Items.Restrict("[MessageClass] = '$GELEN_SINIF'")) { $bulunan += $it }   # tasirken koleksiyon degismesin
  if (-not $bulunan.Count) { return }
  $L = New-Object System.Collections.ArrayList
  foreach ($x in @(Gelen-Oku)) { [void]$L.Add($x) }
  $hedef = Assist-Klasoru; $yeni = 0
  foreach ($it in $bulunan) {
    try {
      $govde = [string]$it.Body
      $i = $govde.IndexOf($VERI_ISARET); if ($i -lt 0) { continue }
      $j = $govde.IndexOf(":SON", $i); if ($j -lt 0) { continue }
      $b64 = ($govde.Substring($i + $VERI_ISARET.Length, $j - $i - $VERI_ISARET.Length)) -replace '\s', ''
      $t = $UTF8.GetString([Convert]::FromBase64String($b64)) | ConvertFrom-Json
      $gonderen = [string]$it.SenderName; $ep = Gonderen-Eposta $it; $alinma = Tarih-Iso $it.ReceivedTime
      $msgId = ""; try { $msgId = [string]$it.PropertyAccessor.GetProperty($PR_MSGID) } catch {}
      $tasinan = $it.Move($hedef)
      [void]$L.Add([ordered]@{ id = [guid]::NewGuid().ToString("N"); gonderen = $gonderen; gonderenEposta = $ep; alinma = $alinma; talep = $t;
        posta = @{ id = [string]$tasinan.EntryID; store = [string]$hedef.StoreID; msgId = $msgId; konu = [string]$tasinan.Subject; kimden = $gonderen; tarih = $alinma } })
      $yeni++
    } catch { Write-Host ("  Gelen talep okunamadi: " + $_.Exception.Message) -ForegroundColor Yellow }
  }
  if ($yeni) {
    Gelen-Yaz $L
    Write-Host ("  " + $yeni + " gelen talep alindi")
    try { Bildirim-Goster $BLD_AD ($(if ($yeni -eq 1) { "1 yeni talep geldi" } else { "$yeni yeni talep geldi" })) "talep" } catch {}
  }
}

# Toplanti odalari: Exchange "All Rooms" adres listesi (sirkette 15. kat Pera / Galata).
function Oda-Listesi {
  $L = @()
  foreach ($al in (Ns).AddressLists) {
    if ([string]$al.Name -notmatch '^All Rooms$|Odalar$') { continue }
    foreach ($ae in $al.AddressEntries) { $L += $ae }
  }
  return $L
}
function Tarih-Iso($d) { try { return ([datetime]$d).ToString("yyyy-MM-ddTHH:mm:ss") } catch { return "" } }
function Gonderen-Eposta($it) {
  try {
    if ($it.SenderEmailType -eq "EX") {
      $eu = $it.Sender.GetExchangeUser()
      if ($eu) { return [string]$eu.PrimarySmtpAddress }
    }
    return [string]$it.SenderEmailAddress
  } catch { return "" }
}
function Bayrak-Termin($it) {
  try {
    if ($it.FlagStatus -ne 2) { return "" }
    $d = [datetime]$it.TaskDueDate
    if ($d.Year -ge 4000) { return "" }
    return $d.ToString("yyyy-MM-dd")
  } catch { return "" }
}
function Posta-Ayrinti($it) {
  $govde = [string]$it.Body
  if ($govde.Length -gt 40000) { $govde = $govde.Substring(0, 40000) }
  $msgId = ""; try { $msgId = [string]$it.PropertyAccessor.GetProperty($PR_MSGID) } catch {}
  $ekler = @(); try { foreach ($e in $it.Attachments) { $ekler += [string]$e.FileName } } catch {}
  return @{ ok = $true; id = [string]$it.EntryID; store = [string]$it.Parent.StoreID; msgId = $msgId
    konu = [string]$it.Subject; kimden = [string]$it.SenderName; eposta = (Gonderen-Eposta $it)
    kime = [string]$it.To; bilgi = [string]$it.CC; tarih = (Tarih-Iso $it.ReceivedTime)
    bayrakTermin = (Bayrak-Termin $it); govde = $govde; ekler = @($ekler) }
}
function Posta-Bul([string]$id, [string]$store, [string]$msgId) {
  $ns = Ns
  if ($id) {
    try { if ($store) { return $ns.GetItemFromID($id, $store) } else { return $ns.GetItemFromID($id) } } catch {}
  }
  # EntryID tutmadiysa (posta tasinmis olabilir) Internet Message-ID ile
  # Gelen Kutusu agacinda aranir.
  if ($msgId) {
    $filtre = '@SQL="' + $PR_MSGID + '" = ''' + $msgId.Replace("'", "''") + ''''
    $kuyruk = New-Object System.Collections.Queue
    $kuyruk.Enqueue($ns.GetDefaultFolder(6))
    try { $kuyruk.Enqueue($ns.GetDefaultFolder(5)) } catch {}
    $bakilan = 0
    while ($kuyruk.Count -gt 0 -and $bakilan -lt 200) {
      $f = $kuyruk.Dequeue(); $bakilan++
      try { $r = $f.Items.Restrict($filtre); if ($r.Count -gt 0) { return $r.Item(1) } } catch {}
      try { foreach ($alt in $f.Folders) { $kuyruk.Enqueue($alt) } } catch {}
    }
  }
  return $null
}

# --- Masaustu bildirimi (Windows) -------------------------------------
# Bildirimi TARAYICI degil bu sunucu verir: sekme kapaliyken de gelsin diye
# (13.10.2026). Sunucu her dakika isler.json'a bakar, ayarlardaki hatirlatma
# kurallarina gore Windows bildirimi gosterir. Bildirime tiklaninca tarayicida
# ilgili is acilir (#ac=<kayit>|<aksiyon>).
# Ayni hatirlatma bir kez gonderilir; gonderilenler bildirim.json'da tutulur
# (isler.json'a YAZILMAZ: sayfanin kaydiyla cakismasin).
# Turkce harfler {XXXX} (unicode kodu) ile yazilir; dosya ASCII kalmali.
function U([string]$s) { return [regex]::Replace($s, '\{([0-9A-F]{4})\}', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) }) }
$BLD_AD = "Assist"
$BLD_UYG = "IsTakip.Bildirim"
$bldDos = Join-Path $veriKok "bildirim.json"
$script:bldSon = [datetime]::MinValue
$script:postaSon = [datetime]::MinValue
$script:postaBekleyen = 0
$script:bildirimHazir = $false
try {
  $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
  $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
  # Bildirimde kaynak adi "Assist" gorunsun (kullanici kaydi, yonetici gerekmez)
  $rk = "HKCU:\Software\Classes\AppUserModelId\$BLD_UYG"
  if (-not (Test-Path $rk)) { New-Item -Path $rk -Force | Out-Null }
  Set-ItemProperty -Path $rk -Name DisplayName -Value $BLD_AD
  $script:bildirimHazir = $true
} catch { Write-Host ("  Windows bildirimi kullanilamiyor: " + $_.Exception.Message) -ForegroundColor Yellow }

# "assist:<hedef>" adresi Sunucu\pencere.ps1'i calistirir (kullanici kaydi, yonetici gerekmez).
# Duz PowerShell (gizli): konsolu "headless" acan komutu kurumsal guvenlik yazilimi zararli sayiyor (15.10.2026).
$script:assistAdres = $false
try {
  $rk = "HKCU:\Software\Classes\assist"
  $komut = 'powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + (Join-Path $betikKok "pencere.ps1") + '" -Port ' + $port + ' -Url "%1"'
  if (-not (Test-Path "$rk\shell\open\command")) { New-Item -Path "$rk\shell\open\command" -Force | Out-Null }
  Set-ItemProperty -Path $rk -Name "(default)" -Value "URL:Assist"
  Set-ItemProperty -Path $rk -Name "URL Protocol" -Value ""
  Set-ItemProperty -Path "$rk\shell\open\command" -Name "(default)" -Value $komut
  $script:assistAdres = $true
} catch { Write-Host ("  assist: adresi kaydedilemedi: " + $_.Exception.Message) -ForegroundColor Yellow }

# Acilis ekrani bekler: Outlook / MiCollab KULLANILABILIR mi? (15.10.2026 olcumu: Outlook ~6, MiCollab ~5 sn)
#  - Outlook: Windows "Application" gunlugune eklentiler yuklenince Outlook olay 45 yazar.
#  - MiCollab: uca.log'a oturum acilinca LoginViewModel satirlari yazilir.
#  Yuklu degilse beklenmez; 40 sn'den uzun suredir calisiyorsa hazir sayilir.
function Uygulama-Hazir([string]$ad) {
  if ($ad -eq "outlook") {
    if (-not (Test-Path -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE")) { return $true }
    $p = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) | Select-Object -First 1
  } else {
    if (-not (Test-Path -LiteralPath (Join-Path $env:APPDATA "Mitel\UC"))) { return $true }
    $p = @(Get-Process -Name UCA -ErrorAction SilentlyContinue) | Select-Object -First 1
  }
  if (-not $p) { return $false }
  $bas = $null; try { $bas = $p.StartTime } catch {}
  if (-not $bas -or ((Get-Date) - $bas).TotalSeconds -ge 40) { return $true }
  if ($ad -eq "outlook") {
    $ev = Get-WinEvent -FilterHashtable @{ LogName = "Application"; ProviderName = "Outlook"; Id = 45; StartTime = $bas } -MaxEvents 1 -ErrorAction SilentlyContinue
    return [bool]$ev
  }
  try {
    $fs = [System.IO.File]::Open((Join-Path $env:APPDATA "Mitel\UC\uca.log"), "Open", "Read", "ReadWrite")
    try { if ($fs.Length -gt 300000) { [void]$fs.Seek(-300000, "End") }; $son = (New-Object System.IO.StreamReader($fs)).ReadToEnd() } finally { $fs.Close() }
    foreach ($m in [regex]::Matches($son, '(?m)^(\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2})\.\d+ \S+\s+\[[^\]]*\] Mitel\.Communicator\.Views\.ViewModels\.LoginViewModel - (OfficeIntegration|Finished operation)')) {
      # tarih metin olarak KARSILASTIRILMAZ: "/" Turkce bolge ayarinda "." olur (15.10.2026 hatasi)
      if ([datetime]::ParseExact($m.Groups[1].Value, "yyyy/MM/dd HH:mm:ss", [Globalization.CultureInfo]::InvariantCulture) -ge $bas.AddSeconds(-1)) { return $true }
    }
  } catch { return $true }
  return $false
}

function Xml-Kacis([string]$s) { return [System.Security.SecurityElement]::Escape($s) }
# Gosterilen her bildirim bir gecmise de yazilir (Windows bildirimi kaybolunca tepsi simgesinden
# "Bildirimler" ile yeniden gorulur). Metinlerde is adi yok (yalniz sayi/saat), dosya sifresiz;
# son 7 gun, en cok 60 kayit.
$bldGecmisDos = Join-Path $veriKok "bildirim_gecmis.json"
$tepsiDos = Join-Path $veriKok "tepsi.json"   # tepside bildirimlerin son gorulme zamani
function Bildirim-Gecmisi {
  if (-not (Test-Path -LiteralPath $bldGecmisDos)) { return @() }
  try { $ham = [System.IO.File]::ReadAllText($bldGecmisDos, $UTF8) | ConvertFrom-Json; return @($ham | ForEach-Object { $_ }) } catch { return @() }
}
function Bildirim-Kaydet([string]$metin, [string]$hedef) {
  $sinir = (Get-Date).AddDays(-7).ToString("o")
  $L = @(Bildirim-Gecmisi | Where-Object { [string]$_.t -gt $sinir }) + @([pscustomobject]@{ t = (Get-Date).ToString("o"); metin = $metin; hedef = $hedef })
  if ($L.Count -gt 60) { $L = $L[($L.Count - 60)..($L.Count - 1)] }
  try { [System.IO.File]::WriteAllText($bldGecmisDos, (ConvertTo-Json -InputObject @($L) -Depth 3), $UTF8) } catch {}
}
function Bildirim-Goster([string]$baslik, [string]$metin, [string]$hedef) {
  Bildirim-Kaydet $metin $hedef
  if (-not $script:bildirimHazir) { return }
  $url = if ($script:assistAdres) { "assist:" + [Uri]::EscapeDataString($hedef) } else { "http://localhost:$port/?b=" + [DateTime]::Now.Ticks + "#" + $hedef }
  $xml = '<toast activationType="protocol" launch="' + (Xml-Kacis $url) + '"><visual><binding template="ToastGeneric">' +
         '<text>' + (Xml-Kacis $baslik) + '</text><text>' + (Xml-Kacis $metin) + '</text></binding></visual></toast>'
  $d = New-Object Windows.Data.Xml.Dom.XmlDocument
  $d.LoadXml($xml)
  [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($BLD_UYG).Show((New-Object Windows.UI.Notifications.ToastNotification $d))
}
function Gun-Fark([string]$t) {
  if (-not $t) { return $null }
  try { return ([datetime]::ParseExact($t.Substring(0, 10), "yyyy-MM-dd", $null) - (Get-Date).Date).Days } catch { return $null }
}
# Bayrakli / kategorili postalarin kimlikleri (govde okunmaz, hizli)
function Posta-Kimlikleri($p) {
  $f = Klasor ([string]$p.klasor)
  $items = $f.Items
  if ($p.kaynak -eq "bayrakli") { $items = $items.Restrict("[FlagStatus] <> 0") }
  elseif ($p.kaynak -eq "kategori" -and $p.kategori) { $items = $items.Restrict("[Categories] = '" + ([string]$p.kategori).Replace("'", "''") + "'") }
  $items.Sort("[ReceivedTime]", $true)
  $gun = [int]$p.gun; if ($gun -le 0) { $gun = 14 }
  $sinir = (Get-Date).Date.AddDays(-$gun)
  $ids = New-Object System.Collections.ArrayList
  for ($i = 1; $i -le $items.Count -and $ids.Count -lt 300; $i++) {
    $it = $items.Item($i)
    if ($it.Class -ne 43) { continue }
    if ($p.kaynak -eq "bayrakli" -and $it.FlagStatus -ne 2) { continue }
    if ($p.kaynak -ne "kategori" -and [datetime]$it.ReceivedTime -lt $sinir) { break }
    [void]$ids.Add([string]$it.EntryID)
  }
  return ,$ids
}
function Bildirim-Kontrol {
  if (-not $script:bildirimHazir -or -not (Test-Path -LiteralPath $veriDos)) { return }
  $V = (Veri-Oku) | ConvertFrom-Json
  if (-not $V -or -not $V.ayar -or -not $V.ayar.masaustu) { return }
  $A = $V.ayar
  $bugun = (Get-Date).ToString("yyyy-MM-dd")
  $hg = @($A.hatirlatGun | ForEach-Object { [int]$_ })
  $gonderilen = @{}
  if (Test-Path -LiteralPath $bldDos) {
    try { ([System.IO.File]::ReadAllText($bldDos, $UTF8) | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $gonderilen[$_.Name] = [string]$_.Value } } catch {}
  }
  # Her yeni hatirlatma bir "tur" tasir: 0 termini gecti, 1 bugun, 2 yaklasiyor, 3 posta.
  # Bildirimde IS ADI YAZMAZ (KVKK: kilit ekraninda okunabilir) - yalniz sayilar.
  $yeni = New-Object System.Collections.ArrayList
  $ekle = {
    param($id, $termin)
    $g = Gun-Fark $termin
    if ($null -eq $g) { return }
    if ($g -lt 0) { $an = if ($A.gecikenHerGun) { "$id|gecti|$bugun" } else { "$id|gecti|$termin" }; $tur = 0 }
    elseif ($hg -contains $g) { $an = "$id|$g|$termin"; $tur = if ($g -eq 0) { 1 } else { 2 } }
    else { return }
    if (-not $gonderilen.ContainsKey($an)) { [void]$yeni.Add(@{ an = $an; tur = $tur }) }
  }
  $bagli = @{}
  foreach ($k in @($V.kayitlar)) {
    # DIKKAT: PowerShell degisken adlarinda buyuk/kucuk harf ayirmaz; $a/$p
    # kullanmak ayar ($A) ve posta ayarini ($P) ezer (13.10.2026).
    foreach ($po in @($k.postalar)) { if ($po.id) { $bagli[[string]$po.id] = $true } }
    $acik = ($k.durum -eq "acik" -or $k.durum -eq "beklemede")
    if ($acik -and $k.tur -ne "posta" -and $k.termin) { & $ekle $k.id $k.termin }
    if ($k.durum -eq "iptal") { continue }
    foreach ($ak in @($k.aksiyonlar)) {
      if ($ak.durum -eq "tamam" -or -not $ak.termin) { continue }
      & $ekle $ak.id $ak.termin
    }
  }
  # Kayda eklenmemis bayrakli / kategorili postalar (kontrol araliginda bir)
  $P = $A.posta
  if ($P -and ($P.kaynak -eq "bayrakli" -or $P.kaynak -eq "kategori")) {
    $dk = [int]$A.kontrolDk; if ($dk -lt 5) { $dk = 15 }
    if (((Get-Date) - $script:postaSon).TotalMinutes -ge $dk) {
      $script:postaSon = Get-Date
      try {
        $yok = @{}; foreach ($y in @($V.yoksay)) { $yok[[string]$y] = $true }
        $script:postaBekleyen = @((Posta-Kimlikleri $P) | Where-Object { -not $bagli[$_] -and -not $yok[$_] }).Count
      } catch {}
    }
    if ($script:postaBekleyen -gt 0) {
      $an = "posta|" + $script:postaBekleyen + "|" + $bugun
      if (-not $gonderilen.ContainsKey($an)) { [void]$yeni.Add(@{ an = $an; tur = 3; n = $script:postaBekleyen }) }
    }
  }
  # Hatirlaticilar: zamani gelmis (son 24 saat) ve henuz bildirilmemis olusumlar.
  # Bildirimde BASLIK YAZMAZ (kilit ekrani); yalniz saat. Tiklaninca hatirlatici acilir.
  $hatYeni = New-Object System.Collections.ArrayList
  foreach ($h in @($V.hatirlaticilar)) {
    if (-not $h -or $h.aktif -eq $false) { continue }
    $o = Hat-Son $h
    if (-not $o -or $o -lt (Get-Date).AddHours(-24)) { continue }
    $an = "H|" + $h.id + "|" + $o.ToString("yyyy-MM-ddTHH:mm")
    if (-not $gonderilen.ContainsKey($an)) { [void]$hatYeni.Add(@{ an = $an; id = [string]$h.id; saat = $o.ToString("HH:mm") }) }
  }
  if ($hatYeni.Count -gt 3) { Bildirim-Goster $BLD_AD ([string]$hatYeni.Count + (U " hat{0131}rlat{0131}c{0131}n{0131}n zaman{0131} geldi")) "hat" }
  else { foreach ($x in $hatYeni) { Bildirim-Goster $BLD_AD ((U "Hat{0131}rlat{0131}c{0131} {00B7} ") + $x.saat) ("hat=" + $x.id) } }
  foreach ($x in $hatYeni) { [void]$yeni.Add(@{ an = $x.an; tur = 9 }) }
  if ($yeni.Count -eq 0) { return }
  $say = @(0, 0, 0, 0)
  foreach ($b in $yeni) { if ($b.tur -eq 3) { $say[3] = $b.n } elseif ($b.tur -le 2) { $say[$b.tur]++ } }
  $parca = @()
  if ($say[0]) { $parca += [string]$say[0] + (U " i{015F}in termini ge{00E7}ti") }
  if ($say[1]) { $parca += [string]$say[1] + (U " i{015F} bug{00FC}n bitiyor") }
  if ($say[2]) { $parca += [string]$say[2] + (U " i{015F}in termini yakla{015F}{0131}yor") }
  if ($say[3]) { $parca += [string]$say[3] + " posta kayda eklenmedi" }
  $hedef = if ($say[0] + $say[1] + $say[2] -eq 0) { "posta" } else { "ana" }
  if ($parca.Count) { Bildirim-Goster $BLD_AD ($parca -join (U " {00B7} ")) $hedef }
  # Gonderilenleri yaz; 45 gunden eskiler silinir
  $esik = (Get-Date).AddDays(-45).ToString("yyyy-MM-dd")
  $yaz = [ordered]@{}
  foreach ($k in $gonderilen.Keys) { if ($gonderilen[$k] -ge $esik) { $yaz[$k] = $gonderilen[$k] } }
  foreach ($b in $yeni) { $yaz[$b.an] = $bugun }
  [System.IO.File]::WriteAllText($bldDos, (ConvertTo-Json -InputObject $yaz -Compress), $UTF8)
}

# --- MiCollab (sirket telefonu) ---------------------------------------
# Klasor, bilgisayarda kayitli tel: isleyicisinden bulunur (yoksa varsayilan kurulum yeri).
function MiCollab-Klasor {
  foreach ($k in @("HKCU:\Software\Classes\tel\shell\open\command", "HKLM:\SOFTWARE\Classes\tel\shell\open\command")) {
    try { $c = [string](Get-ItemProperty -LiteralPath $k -ErrorAction Stop).'(default)'
          if ($c -match 'Mitel' -and $c -match '"?([^"]+\\)ucadialer\.exe') { return $Matches[1] } } catch {}
  }
  return "C:\Program Files (x86)\Mitel\Unified Communicator Advanced 7.3\"
}
# Outlook toplanti penceresinde Teams eklentisinin dugmesine basar; govdeye Teams baglantisi
# gelince $true. Pencere, toplantinin Inspector basligiyla bulunur (ayni adli baska pencereye basilmasin).
function Teams-Toplanti-Ekle($t) {
  Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
  $AE = [System.Windows.Automation.AutomationElement]
  $baslik = ""; try { $baslik = [string]$t.GetInspector.Caption } catch {}
  $pen = $null
  for ($i = 0; $i -lt 20 -and -not $pen; $i++) {
    foreach ($w in $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)) {
      if ($w.Current.ClassName -eq 'rctrl_renwnd32' -and $baslik -and $w.Current.Name -eq $baslik) { $pen = $w; break }
    }
    if (-not $pen) { Start-Sleep -Milliseconds 500 }
  }
  if (-not $pen) { Write-Host "  Teams: toplanti penceresi bulunamadi"; return $false }
  $dug = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
  $hedef = $null
  for ($i = 0; $i -lt 10 -and -not $hedef; $i++) {
    foreach ($d in $pen.FindAll([System.Windows.Automation.TreeScope]::Descendants, $dug)) { if ($d.Current.Name -match '^\s*(Teams Meeting|Teams Toplant)') { $hedef = $d; break } }
    if (-not $hedef) { Start-Sleep -Milliseconds 500 }
  }
  if (-not $hedef) { Write-Host "  Teams: 'Teams Meeting' dugmesi bulunamadi (eklenti kapali olabilir)"; return $false }
  try { $hedef.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() } catch { return $false }
  for ($i = 0; $i -lt 30; $i++) { Start-Sleep -Seconds 1; if (([string]$t.Body) -match 'teams\.microsoft\.com/l/meetup-join') { return $true } }
  Write-Host "  Teams: baglanti 30 sn icinde gelmedi"
  return $false
}

# MiCollab arama penceresindeki "gorusmeyi bitir" dugmesine basar; basilan dugmenin adini dondurur.
# Oncelik sirasi: acik "bitir/kapat" adlari once, genel "iptal" en son. Ana liste penceresinin
# kapat (X) dugmesine dokunulmaz (o pencere gorusme degil, MiCollab'i gizler).
function MiCollab-Gorusme-Bitir {
  Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
  $AE = [System.Windows.Automation.AutomationElement]
  $dugme = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
  # DIKKAT: dizide virgul "+"dan once islenir; birlestirilen her oge PARANTEZ icinde olmali
  # (13.10.2026: son desen ikiye bolunup "Not enough )'s" hatasi verdi).
  $oncelik = @(
    '(?i)^\s*(end\s*call|hang\s*up|hangup|disconnect|drop|end)\s*$',
    '(?i)(end\s*call|hang\s*up|hangup|disconnect)',
    ('(?i)(bitir|sonland|' + (U "g{00F6}r{00FC}{015F}meyi kapat") + '|aramay' + (U "{0131}") + ' kapat)'),
    ('(?i)^\s*(cancel|' + (U "iptal") + ')\s*$')
  )
  $adaylar = @(); $adlar = @()
  foreach ($p in $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)) {
    $sur = ""; try { $sur = (Get-Process -Id $p.Current.ProcessId -ErrorAction Stop).ProcessName } catch {}
    if (@("UCA", "UCASoftphoneManager", "53xxPCApp") -notcontains $sur) { continue }
    foreach ($d in $p.FindAll([System.Windows.Automation.TreeScope]::Descendants, $dugme)) {
      $ad = [string]$d.Current.Name; $aid = [string]$d.Current.AutomationId
      $adlar += ("[" + [string]$p.Current.Name + "] " + $ad + $(if ($aid) { " #" + $aid } else { "" }))
      for ($i = 0; $i -lt $oncelik.Count; $i++) { if ($ad -match $oncelik[$i] -or $aid -match $oncelik[$i]) { $adaylar += ,@($i, $d, $ad); break } }
    }
  }
  if ($adaylar.Count) {
    $sec = $adaylar | Sort-Object { $_[0] } | Select-Object -First 1
    try { $sec[1].GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); Write-Host ("  Gorusme kapatildi: " + $sec[2]); return [string]$sec[2] } catch {}
  }
  Write-Host ("  Gorusme kapatma dugmesi bulunamadi. MiCollab dugmeleri: " + (($adlar | Select-Object -First 60) -join " | "))
  return ""
}
# Turkce harfleri sadelestir (rehber aramasi icin): c g i I o s u
function Harf-Sadele([string]$s) {
  $harita = @{ 0x00E7 = 'c'; 0x00C7 = 'C'; 0x011F = 'g'; 0x011E = 'G'; 0x0131 = 'i'; 0x0130 = 'I'; 0x00F6 = 'o'; 0x00D6 = 'O'; 0x015F = 's'; 0x015E = 'S'; 0x00FC = 'u'; 0x00DC = 'U' }
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $s.ToCharArray()) { $k = [int]$ch; if ($harita.ContainsKey($k)) { [void]$sb.Append($harita[$k]) } else { [void]$sb.Append($ch) } }
  return $sb.ToString()
}

# --- Kisi rehberi: MiCollab + Outlook kisileri + sirket rehberi (GAL) -------
# Uc kaynak tek listede birlesir (ayni kisi e-postasina ya da adina gore tek kayit).
# Dahili/telefon onceligi: MiCollab > sirket rehberi > Outlook kisileri.
# Liste kisisel veri: sifreli onbellege (rehber.dat) yazilir, gunde bir yenilenir.
$rehberDos = Join-Path $veriKok "rehber.dat"
# Rehber kurali degisince yukselt: eski surumlu onbellek 24 saati beklemeden yeniden kurulur
# (13.10.2026: MiCollab'in 107 kisisi eski onbellekte TEK kayitta birlesik kalmisti).
$REHBER_SURUM = 2
# Numaradan rehberdeki ad (rehber onbellegi; dosya degismedikce yeniden cozulmez). Bulunamazsa "".
$script:rehAdAnahtar = ""; $script:rehAd = @{}
function Rehber-Ad([string]$no) {
  $r = ($no -replace '\D', ''); if (-not $r -or -not (Test-Path -LiteralPath $rehberDos)) { return "" }
  $z = [string](Get-Item -LiteralPath $rehberDos).LastWriteTimeUtc.Ticks
  if ($script:rehAdAnahtar -ne $z) {
    $script:rehAd = @{}
    try { foreach ($k in @(((Sifreli-Oku $rehberDos) | ConvertFrom-Json).kisiler | ForEach-Object { $_ })) { $t = ([string]$k.tel -replace '\D', ''); if ($t -and -not $script:rehAd.ContainsKey($t)) { $script:rehAd[$t] = [string]$k.ad } } } catch {}
    $script:rehAdAnahtar = $z
  }
  if ($script:rehAd.ContainsKey($r)) { return $script:rehAd[$r] } else { return "" }
}

function Rehber-Kur {
  $epostaya = @{}; $ada = @{}; $liste = New-Object System.Collections.ArrayList
  $ekle = {
    param([string]$ad, [string]$eposta, [string]$tel, [string]$kaynak)
    $ad = $ad.Trim(); $eposta = $eposta.Trim(); $tel = $tel.Trim()
    if ($eposta -notmatch '^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$') { $eposta = "" }
    if (($tel -replace '[^\d]', '').Length -gt 15 -or $ad.Length -gt 120) { return }   # birlesik/bozuk kayit (tek numara degil)
    if (-not $ad -and -not $eposta) { return }
    if (-not $ad) { $ad = $eposta }
    $adAnahtar = (Harf-Sadele $ad).ToLowerInvariant() -replace '\s+', ' '
    $k = $null
    if ($eposta -and $epostaya.ContainsKey($eposta.ToLowerInvariant())) { $k = $epostaya[$eposta.ToLowerInvariant()] }
    elseif ($ada.ContainsKey($adAnahtar)) { $k = $ada[$adAnahtar] }
    if (-not $k) {
      $k = @{ ad = $ad; eposta = $eposta; tel = $tel; kaynak = @($kaynak) }
      [void]$liste.Add($k); $ada[$adAnahtar] = $k
    } else {
      if (-not $k.eposta -and $eposta) { $k.eposta = $eposta }
      if (-not $k.tel -and $tel) { $k.tel = $tel }
      if ($k.kaynak -notcontains $kaynak) { $k.kaynak += $kaynak }
    }
    if ($k.eposta) { $epostaya[$k.eposta.ToLowerInvariant()] = $k }
  }
  # 1) MiCollab (dahili numaralar burada)
  Get-ChildItem -LiteralPath (Join-Path $env:APPDATA "Mitel\UC") -Recurse -Filter "contacts.dat" -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      # DIKKAT: PS 5.1'de ConvertFrom-Json diziyi TEK nesne olarak verir; once degiskene alinmali,
      # yoksa @(...) tum listeyi tek kayit sanir (13.10.2026: 107 kisiden 1'i gorunuyordu).
      $mcKisiler = [System.IO.File]::ReadAllText($_.FullName, $UTF8) | ConvertFrom-Json
      foreach ($c in $mcKisiler) {
        $ad = (([string]$c.FirstName + " " + [string]$c.LastName)).Trim(); if (-not $ad) { $ad = [string]$c.LogName }
        $ep = [string]$c.DefaultEmailAddress; if (-not $ep -and $c.Emails) { $ep = [string]@($c.Emails)[0].EmailAddress }
        $tl = [string]$c.DefaultPhoneNumber; if (-not $tl -and $c.Phones) { $tl = [string]@($c.Phones)[0].Number }
        & $ekle $ad $ep $tl "MiCollab"
      }
    } catch {}
  }
  # 2) Sirket rehberi (yalniz kullanici kayitlari; dagitim listeleri haric)
  # Sunucu tek is parcacikli: buyuk rehberde tarama dakikalarca surup sunucuyu
  # kilitliyordu (13.10.2026, bat "tanimlanamayan uygulama" dedi) -> en cok 40 sn.
  $gSaat = [System.Diagnostics.Stopwatch]::StartNew()
  try {
    foreach ($ae in (Ns).GetGlobalAddressList().AddressEntries) {
      if ($gSaat.Elapsed.TotalSeconds -gt 40) { Write-Host "  Sirket rehberi 40 sn'de tamamlanamadi; alinabilenler kullanildi." -ForegroundColor Yellow; break }
      if ($ae.AddressEntryUserType -ne 0) { continue }
      try { $eu = $ae.GetExchangeUser(); if ($eu) { & $ekle ([string]$eu.Name) ([string]$eu.PrimarySmtpAddress) ([string]$eu.BusinessTelephoneNumber) "Rehber" } } catch {}
    }
  } catch {}
  # 3) Outlook kisileri
  try {
    foreach ($it in (Ns).GetDefaultFolder(10).Items) {
      if ($it.Class -ne 40) { continue }
      $tl = [string]$it.BusinessTelephoneNumber; if (-not $tl) { $tl = [string]$it.MobileTelephoneNumber }
      & $ekle ([string]$it.FullName) ([string]$it.Email1Address) $tl "Outlook"
    }
  } catch {}
  $sirali = @($liste | Sort-Object { $_.ad })
  $json = ConvertTo-Json -InputObject @{ surum = $REHBER_SURUM; zaman = (Get-Date).ToString("o"); kisiler = $sirali } -Depth 4 -Compress
  Sifreli-Yaz $rehberDos $json
  return $json
}

# Duz metin kayitlar varsa once sifrelenir. Hata olursa sunucu acilmaz (veri korunur).
$gocSayi = Sifre-Gocu
Gecici-Temizle
# Assist kisayolu (masaustu, Baslangic). DIKKAT (15.10.2026): konsolu gizli (headless) acan kisayolu
# kurumsal guvenlik yazilimi SILIYOR (zararli yazilim kalibi); bat dogrudan acilir.
function Assist-Kisayol([string]$lnk) {
  $k = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk)
  $k.TargetPath = Join-Path $portalKok "Assist.bat"; $k.Arguments = ""
  $k.WorkingDirectory = $portalKok; $k.WindowStyle = 7; $k.Description = "STFA Assist"
  $k.IconLocation = (Join-Path $portalKok "Assist.ico") + ",0"
  $k.Save()
}
# Ad degisikligi (Is Takip -> Assist, 13.10.2026): eski otomatik baslat kisayolu yeni bat'a tasinir.
try {
  $eskiLnk = Join-Path ([Environment]::GetFolderPath("Startup")) "Is Takip.lnk"
  if (Test-Path -LiteralPath $eskiLnk) {
    Assist-Kisayol (Join-Path ([Environment]::GetFolderPath("Startup")) "Assist.lnk")
    Remove-Item -LiteralPath $eskiLnk -Force
  }
} catch {}
# Logo degisti (15.10.2026, AssistLogo.ico -> Assist.ico; yeni ad = Windows simge onbellegi yenilenir):
# masaustu ve Baslangic klasorundeki Assist kisayollari yeni simgeye gecirilir.
try {
  $yeniSimge = Join-Path $portalKok "Assist.ico"
  if (Test-Path -LiteralPath $yeniSimge) {
    $ws = New-Object -ComObject WScript.Shell
    foreach ($kl in @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Startup"))) {
      foreach ($l in @(Get-ChildItem -LiteralPath $kl -Filter "*.lnk" -ErrorAction SilentlyContinue)) {
        $k = $ws.CreateShortcut($l.FullName)
        if ($k.IconLocation -match 'AssistLogo\.ico|AssistSimge\.ico') { $k.IconLocation = $yeniSimge + ",0"; $k.Save() }
      }
    }
  }
} catch {}
# Windows acilisinda otomatik baslat VARSAYILAN ACIK (15.10.2026): ilk calismada Baslangic kisayolu konur.
# Kullanici sonra Ayarlar'dan kapatirsa "otobaslatIlk" isareti sayesinde bir daha kendiliginden acilmaz.
try {
  if (-not $env:ISTAKIP_PORT -and -not $ayar.otobaslatIlk) {
    $bLnk = Join-Path ([Environment]::GetFolderPath("Startup")) "Assist.lnk"
    if (-not (Test-Path -LiteralPath $bLnk)) { Assist-Kisayol $bLnk }
    $ayar | Add-Member -NotePropertyName otobaslatIlk -NotePropertyValue $true -Force
    [System.IO.File]::WriteAllText($ayarDos, ($ayar | ConvertTo-Json -Depth 4), $UTF8)
  }
} catch {}# Outlook/MiCollab/Teams'i bat baslatir (Sunucu\uygulamalar.ps1): sunucu zaten aciksa da calissin diye.

# --- Dinleyici -------------------------------------------------------
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
try { $listener.Prefixes.Add("http://127.0.0.1:$port/") } catch {}
$listener.Start()

Write-Host ""
Write-Host "  ASSIST SUNUCUSU  v$SURUM"
Write-Host "  ------------------------------------------------"
Write-Host "  Adres  : http://localhost:$port/"
Write-Host "  Veri   : $veriKok  (sifreli - Windows hesabina bagli)"
if ($gocSayi -gt 0) { Write-Host "  $gocSayi duz metin dosya sifrelendi; duz halleri silindi." -ForegroundColor Cyan }
Write-Host "  Ayar   : $ayarDos"
Write-Host ""
Write-Host "  Bu pencere acik kaldikca uygulama calisir. Kapatmak icin pencereyi kapatin."
Write-Host ""

while ($listener.IsListening) {
  # Istek beklerken saniyede bir betik dosyasina bakilir. Dosya guncellenmisse
  # sunucu kendini yeniden yukler: sayfa yeni surume gecip eski sunucuya yeni
  # bir uc sordugunda "Bilinmeyen uc" hatasi alinmasin (13.10.2026).
  try { $bekle = $listener.BeginGetContext($null, $null) } catch { break }
  $yenidenYukle = $false
  while (-not $bekle.AsyncWaitHandle.WaitOne(1000)) {
    # Dakikada bir hatirlatma kontrolu (Windows bildirimi)
    if (((Get-Date) - $script:bldSon).TotalSeconds -ge 60) {
      $script:bldSon = Get-Date
      try { Bildirim-Kontrol } catch { Write-Host ("  Bildirim hatasi: " + $_.Exception.Message) -ForegroundColor Yellow }
      if (-not $env:ISTAKIP_PORT) { try { Gelen-Talep-Kontrol } catch { Write-Host ("  Gelen talep hatasi: " + $_.Exception.Message) -ForegroundColor Yellow } }
      try { Gecici-Temizle } catch {}
      try { Mobil-Disa-Aktar } catch { Write-Host ("  Mobil disa aktarma hatasi: " + $_.Exception.Message) -ForegroundColor Yellow }
    }
    try { $z = (Get-Item -LiteralPath $PSCommandPath).LastWriteTimeUtc.Ticks } catch { $z = $BETIK_ZAMAN }
    if ($z -ne $BETIK_ZAMAN) { $yenidenYukle = $true; break }
  }
  if ($yenidenYukle) {
    Write-Host "  Sunucu dosyasi guncellendi - yeniden yukleniyor..." -ForegroundColor Cyan
    Start-Sleep -Milliseconds 500   # dosya yazimi bitsin
    try { $listener.Stop(); $listener.Close() } catch {}
    & $PSCommandPath
    return
  }
  try { $ctx = $listener.EndGetContext($bekle) } catch { break }
  try {
    $yol   = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath)
    $metod = $ctx.Request.HttpMethod

    if ($yol -eq "/__saglik") {
      Yaz-Json $ctx @{ ok = $true; urun = "istakip"; surum = $SURUM; portalKok = $portalKok; veriKok = $veriKok;
                       pid = $PID; betikZaman = [string]$BETIK_ZAMAN }
      $ctx.Response.OutputStream.Close(); continue
    }

    if ($yol.StartsWith("/api/")) {
      # Istek kaydi: sunucu bir istekte takilirsa kayittaki SON satir hangisi oldugunu gosterir.
      if (($yol -ne "/api/pencere/hedef" -or $metod -ne "GET") -and $yol -ne "/api/acilis/durum") { Write-Host ("  > " + (Get-Date).ToString("HH:mm:ss") + " " + $metod + " " + $yol + $(if ($yol -ne "/api/kisi/coz") { $ctx.Request.Url.Query })) }   # kisi adi kayda yazilmaz; pencerenin 2 sn yoklamasi yazilmaz
      # Baska bir sayfadan gelen istek ozel basligi gonderemez (CORS izni yok).
      $koken = [string]$ctx.Request.Headers["Origin"]
      $kokenOk = (-not $koken) -or ($koken -eq "http://localhost:$port") -or ($koken -eq "http://127.0.0.1:$port")
      if (-not $ctx.Request.IsLocal -or $ctx.Request.Headers["X-IsTakip"] -ne "1" -or -not $kokenOk) {
        Yaz-Hata $ctx "Erisim reddedildi." 403
        $ctx.Response.OutputStream.Close(); continue
      }

      # ---------- veri ----------
      if ($yol -eq "/api/veri" -and $metod -eq "GET") {
        $ham = "null"
        $ham = Veri-Oku; if (-not $ham.Trim()) { $ham = "null" }
        Yaz-Cevap $ctx ('{"rev":' + $script:rev + ',"veri":' + $ham + ',"veriKok":' + (ConvertTo-Json $veriKok) + '}') "application/json; charset=utf-8" 200
      }
      elseif ($yol -eq "/api/veri" -and $metod -eq "POST") {
        # Iyimser kilit: iki sekme ayni anda yazarsa biri digerini sessizce ezmesin.
        $istenen = [int](Sorgu $ctx "rev")
        if ($istenen -ne $script:rev) { Yaz-Json $ctx @{ ok = $false; hata = "rev"; rev = $script:rev } 409 }
        else {
          $govde = Govde-Oku $ctx
          try { $null = $govde | ConvertFrom-Json } catch { Yaz-Hata $ctx "Gecersiz veri." 400; $ctx.Response.OutputStream.Close(); continue }
          Yedek-Al
          Sifreli-Yaz $veriDos $govde
          try { Yetim-Posta-Temizle $govde } catch {}
          $script:rev++
          [System.IO.File]::WriteAllText($revDos, [string]$script:rev, $UTF8)
          Yaz-Json $ctx @{ ok = $true; rev = $script:rev }
        }
      }

      # ---------- belge yukle (hatirlaticiya baglanir; sifreli saklanir) ----------
      elseif ($yol -eq "/api/belge/yukle" -and $metod -eq "POST") {
        $ad = [System.IO.Path]::GetFileName([Uri]::UnescapeDataString([string]$ctx.Request.Headers["X-Ad"]))
        $uz = [System.IO.Path]::GetExtension($ad).ToLowerInvariant()
        if ($BELGE_UZ -notcontains $uz) { Yaz-Hata $ctx "Bu dosya turu eklenemez." 400 }
        else {
          $ms = New-Object System.IO.MemoryStream
          $ctx.Request.InputStream.CopyTo($ms)
          $b = $ms.ToArray()
          if ($b.Length -gt 26214400) { Yaz-Hata $ctx "Belge 25 MB'tan buyuk." 413 }
          else {
            $id = [guid]::NewGuid().ToString("N")
            Bayt-Yaz (Join-Path $belgeKok ($id + ".dat")) (Sifrele $b)
            Yaz-Json $ctx @{ ok = $true; id = $id; boyut = $b.Length }
          }
        }
      }

      # ---------- belge ac: gecici klasore cozulur, 1 saat sonra silinir ----------
      elseif ($yol -eq "/api/belge/ac" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $id = [string]$g.id
        $ad = [System.IO.Path]::GetFileName([string]$g.ad)
        $uz = [System.IO.Path]::GetExtension($ad).ToLowerInvariant()
        $kaynak = Join-Path $belgeKok ($id + ".dat")
        if ($id -notmatch '^[0-9a-f]{32}$' -or $BELGE_UZ -notcontains $uz -or -not (Test-Path -LiteralPath $kaynak)) { Yaz-Hata $ctx "Belge bulunamadi." 404 }
        else {
          if (-not (Test-Path -LiteralPath $acKok)) { New-Item -ItemType Directory -Path $acKok | Out-Null }
          $hedef = Join-Path $acKok ($id.Substring(0, 8) + "_" + $ad)
          [System.IO.File]::WriteAllBytes($hedef, (Coz ([System.IO.File]::ReadAllBytes($kaynak))))
          Start-Process -FilePath $hedef
          Yaz-Json $ctx @{ ok = $true }
        }
      }

      # ---------- telefondan (Power Apps) gelen istekler ----------
      # Telefon OneDrive\Assist\mobil_gelen\<kimlik>.json birakir: islem = not | aksiyon | tamamla.
      # GET: bekleyen dosyalar; postasi olanlarda posta Outlook'ta internetMessageId ile bulunup Assist'in
      # posta bilgisine (EntryID, depo) cevrilir. Sayfa kaydi olusturur, sonra POST .../sil ile dosyalari siler.
      elseif ($yol -eq "/api/mobil/gelen") {
        $L = @()
        $kl = Join-Path $env:OneDriveCommercial $MOBIL_GELEN
        if ($env:OneDriveCommercial -and (Test-Path -LiteralPath $kl)) {
          foreach ($f in @(Get-ChildItem -LiteralPath $kl -Filter "*.json" -File | Sort-Object LastWriteTime | Select-Object -First 30)) {
            $v = $null; try { $v = [System.IO.File]::ReadAllText($f.FullName, $UTF8) | ConvertFrom-Json } catch {}
            $po = $null
            if ($v -and $v.posta -and $v.posta.internetMessageId) {
              try {
                $mid = ([string]$v.posta.internetMessageId).Replace("'", "''")
                $it = $null
                foreach ($it2 in (Ns).GetDefaultFolder(6).Items.Restrict("@SQL=""" + $PR_MSGID + """ = '" + $mid + "'")) { $it = $it2; break }
                if ($it) {
                  $ep = ""; try { $ep = [string]$it.Sender.GetExchangeUser().PrimarySmtpAddress } catch {}; if (-not $ep) { $ep = [string]$it.SenderEmailAddress }
                  $po = @{ id = [string]$it.EntryID; store = [string]$it.Parent.StoreID; msgId = [string]$v.posta.internetMessageId; konu = [string]$it.Subject
                           kimden = [string]$it.SenderName; eposta = $ep; tarih = ([datetime]$it.ReceivedTime).ToString("yyyy-MM-ddTHH:mm") }
                }
              } catch {}
            }
            $L += @{ dosya = $f.Name; veri = $v; posta = $po }
          }
        }
        Yaz-Json $ctx @{ ok = $true; gelen = @($L) }
      }
      elseif ($yol -eq "/api/mobil/gelen/sil" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $n = 0
        $kl = Join-Path $env:OneDriveCommercial $MOBIL_GELEN
        foreach ($ad in @($g.dosyalar | ForEach-Object { [string]$_ })) {
          if ($ad -notmatch '^[A-Za-z0-9_.-]+\.json$') { continue }   # yalniz bu klasordeki dosya adi
          $p = Join-Path $kl $ad; if (Test-Path -LiteralPath $p) { try { Remove-Item -LiteralPath $p -Force; $n++ } catch {} }
        }
        Yaz-Json $ctx @{ ok = $true; silinen = $n }
      }

      # ---------- kisiye ozel ad gorunumu ("Hos Geldin <ad>") ----------
      # Paylasilan klasorde IsTakip\kisi_stilleri.json (OneDrive ile herkese gider; GitHub'a GITMEZ).
      # Her Assist yalniz KENDI e-postasina ait stili uygular. Dosyayi yalniz "yonetici" degistirir:
      # yonetici dosyada yazilidir ve ilk kaydeden kisidir (kodda kimse adiyla yazili degil).
      elseif ($yol -eq "/api/kisi-stil") {
        $stilDos = Join-Path $portalKok "kisi_stilleri.json"
        $ben = Ben-Eposta
        $d = $null; if (Test-Path -LiteralPath $stilDos) { try { $d = [System.IO.File]::ReadAllText($stilDos, $UTF8) | ConvertFrom-Json } catch {} }
        $yon = if ($d -and $d.yonetici) { ([string]$d.yonetici).ToLowerInvariant() } else { "" }
        # Yonetici dosyada yaziliysa yalniz o kisi. Dosya yoksa (ilk kurulum / silinmis) bolum yalniz
        # ?yonetici=1 ile acilan sayfada gorunur: kisayoldan acan kullanicilar hic gormez.
        $yetkili = $ben -and (($yon -and $yon -eq $ben) -or (-not $yon -and (Sorgu $ctx "yonetici") -eq "1"))
        if ($metod -eq "POST") {
          if (-not $yetkili) { Yaz-Hata $ctx "Bu ayari yalniz yonetici degistirebilir." 403 }
          else {
            $g = (Govde-Oku $ctx) | ConvertFrom-Json
            $L = @()
            foreach ($k in @($g.kisiler | ForEach-Object { $_ })) {
              $ep = ([string]$k.eposta).Trim().ToLowerInvariant()
              if ($ep -notmatch '^[^@\s]+@[^@\s]+\.[a-z]{2,}$') { continue }
              $yt = [string]$k.yaziTipi; if ($yt -notin @("duz", "el", "genis", "serif")) { $yt = "duz" }
              $rk = [string]$k.renk; if ($rk -notmatch '^#[0-9A-Fa-f]{6}$') { $rk = "" }
              $by = [int]$k.boyut; if ($by -lt 16 -or $by -gt 44) { $by = 0 }
              # Tek seferlik mesaj: metin degisince yeni kimlik alir (kisinin Assist'i bu kimligi bir kez gosterir)
              $ms = ([string]$k.mesaj).Trim(); if ($ms.Length -gt 600) { $ms = $ms.Substring(0, 600) }
              $eski = if ($d) { @($d.kisiler | ForEach-Object { $_ }) | Where-Object { ([string]$_.eposta).ToLowerInvariant() -eq $ep } | Select-Object -First 1 } else { $null }
              $mid = if (-not $ms) { "" } elseif ($eski -and [string]$eski.mesaj -eq $ms -and $eski.mesajId) { [string]$eski.mesajId } else { [string][DateTime]::Now.Ticks }
              $L += [ordered]@{ eposta = $ep; yaziTipi = $yt; renk = $rk; kalin = [bool]$k.kalin; italik = [bool]$k.italik; boyut = $by; mesaj = $ms; mesajId = $mid }
            }
            $yeni = [ordered]@{ yonetici = $(if ($yon) { $yon } else { $ben }); kisiler = @($L) }
            [System.IO.File]::WriteAllText($stilDos, (ConvertTo-Json -InputObject $yeni -Depth 4), $UTF8)
            Write-Host ("  Kisi stilleri kaydedildi (" + $L.Count + " kisi)")
            Yaz-Json $ctx @{ ok = $true; sayi = $L.Count }
          }
        } else {
          $kendi = $null
          if ($d -and $ben) { $kendi = @($d.kisiler | ForEach-Object { $_ }) | Where-Object { ([string]$_.eposta).ToLowerInvariant() -eq $ben } | Select-Object -First 1 }
          Yaz-Json $ctx @{ ok = $true; stil = $kendi; yetkili = [bool]$yetkili; kisiler = $(if ($yetkili -and $d) { @($d.kisiler | ForEach-Object { $_ }) } else { @() }) }
        }
      }

      # ---------- kullanicinin adi (ana ekran "Hos Geldin") ----------
      # Once kurumsal hesap (etki alani) gorunen adi, olmazsa Outlook, olmazsa Windows kullanici adi.
      elseif ($yol -eq "/api/kullanici") {
        if (-not $script:kullaniciAd) {
          $ad = ""
          try { Add-Type -AssemblyName System.DirectoryServices.AccountManagement; $ad = [string][System.DirectoryServices.AccountManagement.UserPrincipal]::Current.DisplayName } catch {}
          if (-not $ad) { try { $ad = [string](Outlook).Session.CurrentUser.Name } catch {} }
          if (-not $ad) { $ad = [string]$env:USERNAME }
          $script:kullaniciAd = $ad
        }
        Yaz-Json $ctx @{ ok = $true; ad = $script:kullaniciAd }
      }

      # ---------- Teams / telefon: TARAYICI DEGIL SUNUCU acar ----------
      # Tarayici msteams:/tel: baglantisinda her seferinde "uygulama acilsin mi" diye
      # soruyordu (13.10.2026). Sunucu acinca bu onay cikmaz. Guvenlik: yalniz gecerli
      # e-postadan Teams baglantisi, yalniz rakamdan telefon aramasi kurulur.
      elseif ($yol -eq "/api/iletisim" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $tur = [string]$g.tur; $deger = [string]$g.deger
        if (($tur -eq "teams-yaz" -or $tur -eq "teams-ara") -and $deger -match '^[^@\s"<>]+@[^@\s"<>]+\.[A-Za-z]{2,}$') {
          $yolu = if ($tur -eq "teams-yaz") { "chat" } else { "call" }
          Start-Process ("msteams:/l/" + $yolu + "/0/0?users=" + [Uri]::EscapeDataString($deger))
          Yaz-Json $ctx @{ ok = $true }
        } elseif ($tur -eq "tel" -and ($deger -replace '[^\d+]', '') -match '^\+?\d{2,15}$') {
          $no = $deger -replace '[^\d+]', ''
          $dialer = Join-Path (MiCollab-Klasor) "ucadialer.exe"
          # MiCollab (ThirdPartyDialHandler.DialImp) 3-5 haneli dahiliyi CIPLAK ister ("222"),
          # uzun numarayi da ciplak kabul eder; bu yuzden "tel:" oneki eklenmez (13.10.2026).
          if (Test-Path -LiteralPath $dialer) { Start-Process -FilePath $dialer -ArgumentList $no } else { Start-Process ("tel:" + $no) }
          Write-Host ("  Tel: MiCollab'a " + $no.Length + " haneli numara gonderildi")
          Yaz-Json $ctx @{ ok = $true }
        } elseif ($tur -eq "tel-kapat") {
          # Gorusmeyi kapat. MiCollab'in komut kanali (mailslot) yalniz "dial"i isler, "hangup"u
          # yok sayar (13.10.2026 denendi). Bu yuzden MiCollab'in arama penceresindeki bitir
          # dugmesine UI Automation ile basilir. Dugme bulunamazsa penceredeki dugme adlari kayda
          # yazilir (adlar MiCollab surumune/diline gore degisebilir; buradaki listeye eklenir).
          $sonuc = MiCollab-Gorusme-Bitir
          if ($sonuc) { Yaz-Json $ctx @{ ok = $true; dugme = $sonuc } }
          else { Yaz-Hata $ctx "MiCollab'da kapatma dugmesi bulunamadi; gorusmeyi MiCollab'dan kapatin." 404 }
        } else { Yaz-Hata $ctx "Gecersiz istek." 400 }
      }

      # ---------- balon sayaclari: okunmamis posta + cevapsiz arama ----------
      # okunmamis: Outlook Gelen Kutusu'nun okunmamis sayisi (-1 = Outlook'a ulasilamadi).
      # cevapsiz: MiCollab arama gecmisinde (callhistory.dat, JSON; Type 2 = cevapsiz) ?son= zamanindan
      # sonraki cevapsiz aramalar. Dosya degismedikce yeniden okunmaz.
      elseif ($yol -eq "/api/sayaclar") {
        $okunmamis = -1
        try { $okunmamis = [int](Ns).GetDefaultFolder(6).UnReadItemCount } catch {}
        $son = Sorgu $ctx "son"; $cev = @()
        if ($son) {
          try { $cev = @(Cevapsiz-Listesi ([DateTime]::Parse($son).ToUniversalTime()) | ForEach-Object { @{ ad = $_.ad; no = $_.no; t = $_.t.ToString("o") } }) } catch {}
        }
        Yaz-Json $ctx @{ ok = $true; okunmamis = $okunmamis; cevapsiz = $cev }
      }

      # ---------- rehberi Outlook kisilerine yaz (iPhone'da gelen aramada isim icin) ----------
      # Sirket rehberindeki numarali kisiler Outlook "Kisiler" klasorune "Assist Rehber" kategorisiyle
      # yazilir; Outlook mobil "Kisileri Kaydet" / Exchange esitlemesi bunlari telefona tasir.
      # Dahili (3-5 hane) mobil kuralla: onek + dahili (1 + 646 = 1646, is telefonu). Uzun numara +90'li
      # cep olarak. YALNIZ bu kategorideki kisiler eklenir/guncellenir/silinir; kullanicinin kendi
      # eklediklerine dokunulmaz (ayni e-postali kendi kisisi varsa o kisi atlanir).
      # deneme=true: hicbir sey yazmadan sayilari dondurur. kaldir=true: kategorideki kisilerin hepsini siler.
      elseif ($yol -eq "/api/rehber/telefona" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $onek = ([string]$g.onek) -replace '[^\d+]', ''
        $KAT = "Assist Rehber"
        $f = (Ns).GetDefaultFolder(10)   # olFolderContacts
        $bizim = @{}; $kendi = @{}
        foreach ($it in $f.Items) {
          if ($it.Class -ne 40) { continue }   # olContact
          $ep = ([string]$it.Email1Address).ToLowerInvariant()
          if (([string]$it.Categories) -match [regex]::Escape($KAT)) { $bizim[$(if ($ep) { $ep } else { "ad:" + [string]$it.FullName })] = $it }
          elseif ($ep) { $kendi[$ep] = $true }
        }
        if ($g.kaldir) {
          $n = 0
          if (-not $g.deneme) { foreach ($it in @($bizim.Values)) { try { $it.Delete(); $n++ } catch {} } } else { $n = $bizim.Count }
          Write-Host ("  Telefon rehberi: " + $n + " Assist kisisi Outlook'tan " + $(if ($g.deneme) { "silinecek" } else { "silindi" }))
          Yaz-Json $ctx @{ ok = $true; silinen = $n; deneme = [bool]$g.deneme }
        } else {
          $ben = ""; try { $ben = ([string](Ns).CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress).ToLowerInvariant() } catch {}
          $kisiler = @(); try { $kisiler = @(((Sifreli-Oku $rehberDos) | ConvertFrom-Json).kisiler | ForEach-Object { $_ }) } catch {}
          $ekle = 0; $guncel = 0; $atla = 0; $ayni = 0; $yeniAnahtar = @{}
          foreach ($k in $kisiler) {
            $ad = [string]$k.ad; $ep = ([string]$k.eposta).ToLowerInvariant(); $ham = ([string]$k.tel) -replace '\D', ''
            if (-not $ham -or -not $ad) { continue }
            if ($ad -match '(?i)(c|\u00E7)ay ocag|ocag\u0131|toplant(\u0131|i) odas|\bt\.\s*odas|meeting room') { continue }   # kisi olmayan kayitlar
            if ($ep -and $ep -eq $ben) { continue }
            if ($ep -and $kendi.ContainsKey($ep)) { $atla++; continue }
            $isTel = ""; $cep = ""
            if ($ham.Length -le 5) { $isTel = $onek + $ham }
            else {
              $c = $ham -replace '^00', ''; if ($c.StartsWith("0")) { $c = "90" + $c.Substring(1) }; if ($c.Length -eq 10) { $c = "90" + $c }
              $cep = "+" + $c
            }
            $an = if ($ep) { $ep } else { "ad:" + $ad }
            if ($yeniAnahtar.ContainsKey($an)) { continue }
            $yeniAnahtar[$an] = $true
            if ($g.deneme) { if ($bizim.ContainsKey($an)) { $guncel++ } else { $ekle++ }; continue }
            $it = $bizim[$an]; $yeni = -not $it
            if ($yeni) { $it = $f.Items.Add(2) }   # olContactItem
            $deg = $yeni -or [string]$it.FullName -ne $ad -or [string]$it.BusinessTelephoneNumber -ne $isTel -or [string]$it.MobileTelephoneNumber -ne $cep
            if (-not $deg) { $ayni++; continue }
            $it.FullName = $ad; if ($ep) { $it.Email1Address = $ep }
            $it.BusinessTelephoneNumber = $isTel; $it.MobileTelephoneNumber = $cep
            $it.CompanyName = "STFA"; $it.Categories = $KAT
            try { $it.Save(); if ($yeni) { $ekle++ } else { $guncel++ } } catch {}
          }
          $sil = 0
          foreach ($an in @($bizim.Keys)) { if (-not $yeniAnahtar.ContainsKey($an)) { if (-not $g.deneme) { try { $bizim[$an].Delete() } catch {} }; $sil++ } }
          Write-Host ("  Telefon rehberi: +" + $ekle + " ~" + $guncel + " -" + $sil + " (atlanan " + $atla + ")" + $(if ($g.deneme) { " DENEME" } else { "" }))
          Yaz-Json $ctx @{ ok = $true; deneme = [bool]$g.deneme; eklenen = $ekle; guncellenen = $guncel; ayni = $ayni; silinen = $sil; atlanan = $atla; toplam = $yeniAnahtar.Count }
        }
      }

      # ---------- telefon rehberi: .vcf (kurum politikasi Outlook'tan kisi aktarmayi kapattiysa) ----------
      # "Assist Rehber" kategorisindeki Outlook kisileri tek vCard dosyasi olur. GET: indir;
      # POST {eposta:true}: dosya ekli e-posta KULLANICININ KENDISINE gonderilir (telefonda acip
      # "Tum kisileri ekle"). Baska alici kabul edilmez.
      elseif ($yol -eq "/api/rehber/vcf") {
        $kac = { param([string]$s) $s.Replace('\', '\\').Replace(',', '\,').Replace(';', '\;') }
        $sb = New-Object System.Text.StringBuilder; $n = 0
        foreach ($it in (Ns).GetDefaultFolder(10).Items.Restrict("[Categories] = 'Assist Rehber'")) {
          $fn = [string]$it.FullName; if (-not $fn) { continue }
          [void]$sb.Append("BEGIN:VCARD`r`nVERSION:3.0`r`n")
          [void]$sb.Append("N:" + (& $kac ([string]$it.LastName)) + ";" + (& $kac ([string]$it.FirstName)) + ";;;`r`nFN:" + (& $kac $fn) + "`r`nORG:STFA`r`n")
          if ($it.BusinessTelephoneNumber) { [void]$sb.Append("TEL;TYPE=WORK,VOICE:" + $it.BusinessTelephoneNumber + "`r`n") }
          if ($it.MobileTelephoneNumber) { [void]$sb.Append("TEL;TYPE=CELL:" + $it.MobileTelephoneNumber + "`r`n") }
          if ($it.Email1Address) { [void]$sb.Append("EMAIL;TYPE=INTERNET,WORK:" + $it.Email1Address + "`r`n") }
          [void]$sb.Append("CATEGORIES:Assist Rehber`r`nEND:VCARD`r`n"); $n++
        }
        if (-not $n) { Yaz-Hata $ctx "Once 'Rehberi Outlook kisilerine yaz' ile kisileri olusturun." 404 }
        elseif ($metod -eq "POST") {
          $dos = Join-Path $env:TEMP "STFA Rehber.vcf"
          [System.IO.File]::WriteAllText($dos, $sb.ToString(), $UTF8)
          $ben = ""; try { $ben = [string](Ns).CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {}
          if (-not $ben) { Yaz-Hata $ctx "Kendi e-posta adresiniz bulunamadi." 500 }
          else {
            $m = (Outlook).CreateItem(0)
            $m.To = $ben; $m.Subject = "STFA Rehber (" + $n + " kisi) - telefona ekle"
            $m.Body = (U "Telefonda bu postay{0131} a{00E7}{0131}p ekteki 'STFA Rehber.vcf' dosyas{0131}na dokunun, 'T{00FC}m ki{015F}ileri ekle' deyin.`r`nDahili numaralar mobil {00F6}nekiyle yaz{0131}ld{0131} (646 -> 1646).")
            [void]$m.Attachments.Add($dos)
            $m.Send()
            Remove-Item -LiteralPath $dos -Force -ErrorAction SilentlyContinue
            Write-Host ("  Telefon rehberi: " + $n + " kisilik vcf kullanicinin kendisine gonderildi")
            Yaz-Json $ctx @{ ok = $true; sayi = $n; eposta = $ben }
          }
        } else {
          $b = $UTF8.GetBytes($sb.ToString())
          $ctx.Response.AddHeader("Content-Disposition", "attachment; filename=""STFA Rehber.vcf""")
          $ctx.Response.ContentType = "text/vcard; charset=utf-8"; $ctx.Response.ContentLength64 = $b.Length
          $ctx.Response.OutputStream.Write($b, 0, $b.Length)
        }
      }

      # ---------- tepsi simgesi: bildirim ozeti ----------
      # Son 48 saatin bildirimleri (gecmisten) + cevapsiz aramalar, yeniden eskiye en cok 15 oge.
      # yeni: tepside son "goruldu"dan (ve cevapsizlarda Assist'te gorulmeden) sonra gelenler.
      elseif ($yol -eq "/api/tepsi/ozet") {
        $gor = [datetime]::MinValue
        try { $gor = [datetime]::Parse([string]([System.IO.File]::ReadAllText($tepsiDos, $UTF8) | ConvertFrom-Json).gorulme) } catch {}
        $cevSon = [datetime]::MinValue
        try { $vv = (Veri-Oku) | ConvertFrom-Json; if ($vv.cevapsizSon) { $cevSon = [datetime]::Parse([string]$vv.cevapsizSon) } } catch {}
        $sinir = (Get-Date).AddHours(-48); $og = @()
        foreach ($b in (Bildirim-Gecmisi)) {
          try { $t = [datetime]::Parse([string]$b.t) } catch { continue }
          if ($t -lt $sinir) { continue }
          $og += @{ t = $t; metin = [string]$b.metin; hedef = [string]$b.hedef; yeni = ($t -gt $gor) }
        }
        try {
          foreach ($c in (Cevapsiz-Listesi $sinir.ToUniversalTime())) {
            $t = $c.t.ToLocalTime(); $ad = $c.ad; if (-not $ad) { $ad = Rehber-Ad $c.no }; if (-not $ad) { $ad = $c.no }
            $og += @{ t = $t; metin = (U "Cevaps{0131}z arama {00B7} ") + $ad; hedef = "kisiler"; yeni = ($t -gt $gor -and $t -gt $cevSon) }
          }
        } catch {}
        $og = @($og | Sort-Object { $_.t } -Descending | Select-Object -First 15)
        $saatYaz = { param($t) if ($t.Date -eq (Get-Date).Date) { $t.ToString("HH:mm") } else { $t.ToString("dd.MM HH:mm") } }
        $cikti = @($og | ForEach-Object { @{ saat = (& $saatYaz $_.t); metin = $_.metin; hedef = $_.hedef; yeni = [bool]$_.yeni } })
        # Gorevler: geciken / bugun / yaklasan (Ayarlar'daki gun) acik isler ve aksiyonlar, termine gore; ilk 5.
        # Tepsi menusu yalniz kullanici acinca gorunur (kilit ekranindaki Windows bildirimi degil): baslik yazilir.
        $gor2 = @()
        try {
          $yg = [int]$vv.ayar.yaklasanGun; if ($yg -le 0) { $yg = 7 }
          $ekleG = { param($kid, $aid, $no, $baslik, $termin)
            $gf = Gun-Fark $termin; if ($null -eq $gf -or $gf -gt $yg) { return }
            $kalan = if ($gf -lt 0) { [string](-$gf) + (U " g{00FC}n gecikti") } elseif ($gf -eq 0) { U "Bug{00FC}n" } elseif ($gf -eq 1) { U "Yar{0131}n" } else { [string]$gf + (U " g{00FC}n kald{0131}") }
            $script:gorTmp += @{ termin = [string]$termin; no = [string]$no; baslik = [string]$baslik; kalan = $kalan; hedef = "ac=" + $kid + "|" + $aid; gecikti = ($gf -lt 0) } }
          $script:gorTmp = @()
          foreach ($k in @($vv.kayitlar | ForEach-Object { $_ })) {
            $acik = ($k.durum -eq "acik" -or $k.durum -eq "beklemede")
            if ($acik -and $k.tur -ne "posta" -and $k.termin) { & $ekleG $k.id "" $k.no $k.baslik $k.termin }
            if ($k.durum -eq "iptal") { continue }
            foreach ($ak in @($k.aksiyonlar | ForEach-Object { $_ })) { if ($ak.durum -ne "tamam" -and $ak.termin) { & $ekleG $k.id $ak.id $ak.no $ak.metin $ak.termin } }
          }
          $gor2 = @($script:gorTmp | Sort-Object { $_.termin })
        } catch {}
        # E-postalar: Gelen Kutusu'ndaki okunmamislar (son 48 saat, en yeni 8). Tepsi menusu son acildiktan
        # sonra gelenler "yeni" sayilir ve simgedeki sayiya eklenir. Tiklayinca Outlook'ta acilir (/api/posta/ac).
        $posta = @(); $postaSay = 0
        try {
          $gk = (Ns).GetDefaultFolder(6); $postaSay = [int]$gk.UnReadItemCount
          $its = $gk.Items.Restrict("[UnRead] = True"); $its.Sort("[ReceivedTime]", $true)
          $it = $its.GetFirst(); $n = 0
          while ($it -and $n -lt 8) {
            $rt = $null; try { $rt = [datetime]$it.ReceivedTime } catch {}
            if ($rt -and $rt -lt $sinir) { break }
            if ($it.Class -eq 43) {
              $posta += @{ saat = $(if ($rt) { & $saatYaz $rt } else { "" }); kimden = [string]$it.SenderName; konu = [string]$it.Subject
                           id = [string]$it.EntryID; store = [string]$it.Parent.StoreID; yeni = [bool]($rt -and $rt -gt $gor) }
              $n++
            }
            $it = $its.GetNext()
          }
        } catch {}
        Yaz-Json $ctx @{ ok = $true; yeni = @($og | Where-Object { $_.yeni }).Count + @($posta | Where-Object { $_.yeni }).Count; ogeler = $cikti;
          cevapsiz = @($cikti | Where-Object { $_.hedef -eq "kisiler" }); gorevSay = $gor2.Count; gorevler = @($gor2 | Select-Object -First 5);
          postaSay = $postaSay; postalar = @($posta) }
      }
      elseif ($yol -eq "/api/tepsi/goruldu" -and $metod -eq "POST") {
        try { [System.IO.File]::WriteAllText($tepsiDos, (@{ gorulme = (Get-Date).ToString("o") } | ConvertTo-Json), $UTF8) } catch {}
        Yaz-Json $ctx @{ ok = $true }
      }

      # ---------- kisi rehberi (MiCollab + Outlook + sirket rehberi) ----------
      # Onbellek 24 saatten yeniyse ondan verilir; ?yenile=1 ile yeniden kurulur.
      elseif ($yol -eq "/api/rehber") {
        $json = ""
        $taze = (Test-Path -LiteralPath $rehberDos) -and ((Get-Item -LiteralPath $rehberDos).LastWriteTime -gt (Get-Date).AddHours(-24))
        if ($taze -and (Sorgu $ctx "yenile") -ne "1") { try { $json = Sifreli-Oku $rehberDos } catch { $json = "" } }
        if ($json -and $json -notmatch ('"surum":' + $REHBER_SURUM + '[,}]')) { $json = "" }
        if (-not $json) { $json = Rehber-Kur }
        Yaz-Cevap $ctx ('{"ok":true,"rehber":' + $json + '}') "application/json; charset=utf-8" 200
      }

      # ---------- kisi: sirket adres defterinden e-posta ve is telefonu ----------
      # Rehberde adlar Turkce harfsiz olabiliyor ("Ayse Yilmaz"): once oldugu gibi,
      # bulunamazsa harfler sadelestirilerek aranir.
      elseif ($yol -eq "/api/kisi/coz") {
        $ad = Sorgu $ctx "ad"; $bulunan = $null
        # DIKKAT: Select-Object ciktisi PSObject sarili gelir; COM'a oyle verilince
        # CreateRecipient sunucuyu sonsuza dek kilitliyordu (13.10.2026) -> [string] sart.
        $adaylar = @([string]$ad)
        $sade = [string](Harf-Sadele $ad); if ($sade -cne $ad) { $adaylar += $sade }
        foreach ($aday in $adaylar) {
          if (-not $aday) { continue }
          try {
            $r = (Ns).CreateRecipient([string]$aday); $null = $r.Resolve()
            if ($r.Resolved) { $bulunan = $r.AddressEntry; break }
          } catch {}
        }
        if (-not $bulunan) { Yaz-Json $ctx @{ ok = $false; hata = "Rehberde bulunamadi." } }
        else {
          $eu = $null; try { $eu = $bulunan.GetExchangeUser() } catch {}
          if ($eu) { Yaz-Json $ctx @{ ok = $true; ad = [string]$eu.Name; eposta = [string]$eu.PrimarySmtpAddress; tel = [string]$eu.BusinessTelephoneNumber } }
          else { Yaz-Json $ctx @{ ok = $true; ad = [string]$bulunan.Name; eposta = [string]$bulunan.Address; tel = "" } }
        }
      }

      # ---------- acilis ekrani: Outlook / MiCollab hazir mi? (Ayarlar'da kapali olan beklenmez) ----------
      elseif ($yol -eq "/api/acilis/durum") {
        $L = [ordered]@{}
        foreach ($ad in "outlook", "micollab") {
          $istendi = if ($ayar.baslat -and $null -ne $ayar.baslat.$ad) { [bool]$ayar.baslat.$ad } else { $true }
          $L[$ad] = @{ istendi = $istendi; hazir = $(if ($istendi) { Uygulama-Hazir $ad } else { $true }) }
        }
        Yaz-Json $ctx @{ ok = $true; uygulamalar = $L }
      }

      # ---------- uygulama penceresi (Sunucu\pencere.ps1) ----------      # Tepsi / bildirim acik pencereyi one getirirken gidilecek yeri buraya birakir; sayfa
      # odaklaninca alir (GET bir kez verir, sonra siler). Cikis dugmesi pencereyi kapatir;
      # sunucu ve tepsi simgesi calismaya devam eder.
      elseif ($yol -eq "/api/pencere/hedef") {
        if ($metod -eq "POST") { $script:pencereHedef = [string]((Govde-Oku $ctx) | ConvertFrom-Json).hedef; Yaz-Json $ctx @{ ok = $true } }
        else { $h = [string]$script:pencereHedef; $script:pencereHedef = ""; Yaz-Json $ctx @{ ok = $true; hedef = $h } }
      }
      elseif ($yol -eq "/api/pencere/kapat" -and $metod -eq "POST") {
        & (Join-Path $betikKok "pencere.ps1") -Kapat
        Yaz-Json $ctx @{ ok = $true }
      }

      # ---------- Windows acilisinda otomatik baslat ----------
      # Kullanicinin Baslangic klasorune bat dosyasinin kisayolu konur / kaldirilir.
      # Bat sunucuyu ve tarayiciyi acar; sunucu tarayicidan bagimsiz calisir.
      elseif ($yol -eq "/api/otobaslat") {
        $lnk = Join-Path ([Environment]::GetFolderPath("Startup")) "Assist.lnk"
        if ($metod -eq "POST") {
          $g = (Govde-Oku $ctx) | ConvertFrom-Json
          if ($g.ac) { Assist-Kisayol $lnk }   # penceresiz bat + Assist logosu
          elseif (Test-Path -LiteralPath $lnk) { Remove-Item -LiteralPath $lnk -Force }
        }
        Yaz-Json $ctx @{ ok = $true; acik = (Test-Path -LiteralPath $lnk) }
      }

      # ---------- Outlook takvimi: Assist Takvim'inde toplantilar gorunsun ----------
      # ?bas=yyyy-MM-dd&gunSayi=N (<=62). Tekrarlayan toplantilarin o araliga dusen her olusu
      # gelir (IncludeRecurrences). Yalniz okunur; tiklaninca /api/takvim/ac Outlook'ta acar.
      elseif ($yol -eq "/api/takvim" -and $metod -eq "GET") {
        $bas = [datetime]::Today
        try { $bas = [datetime]::ParseExact((Sorgu $ctx "bas"), "yyyy-MM-dd", $null) } catch {}
        $n = 42; try { $n = [math]::Min(62, [math]::Max(1, [int](Sorgu $ctx "gunSayi"))) } catch {}
        $bit = $bas.AddDays($n)
        $items = (Ns).GetDefaultFolder(9).Items     # olFolderCalendar
        $items.IncludeRecurrences = $true
        $items.Sort("[Start]")
        # Restrict tarihleri Windows bolge bicimini ister ("g"); sinirda kalanlar asagida yeniden suzulur
        $suz = $items.Restrict("[Start] < '" + $bit.ToString("g") + "' AND [End] > '" + $bas.ToString("g") + "'")
        $L = New-Object System.Collections.ArrayList
        $it = $suz.GetFirst()
        while ($it -and $L.Count -lt 600) {
          try {
            $s = [datetime]$it.Start; $e = [datetime]$it.End
            if ($s -lt $bit -and $e -gt $bas) {
              $o = @{ id = [string]$it.EntryID; baslik = [string]$it.Subject; bas = $s.ToString("yyyy-MM-ddTHH:mm"); bit = $e.ToString("yyyy-MM-ddTHH:mm");
                tumGun = [bool]$it.AllDayEvent; yer = [string]$it.Location; durum = [int]$it.BusyStatus; toplanti = ([int]$it.MeetingStatus -ne 0) }
              # Kisa araliklarda (anasayfa Toplanti penceresi) ayrintilar da: duzenleyen, katilimci sayisi, Teams baglantisi
              if ($n -le 7) {
                try { $o.duzenleyen = [string]$it.Organizer } catch {}
                try { $o.katilimci = [int]$it.Recipients.Count } catch {}
                $tl = ""
                try { $tl = [string]$it.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/SkypeTeamsMeetingUrl") } catch {}
                if (-not $tl) { try { $m = [regex]::Match([string]$it.Body, 'https://teams\.microsoft\.com/l/meetup-join/[^\s>"]+'); if ($m.Success) { $tl = $m.Value } } catch {} }
                $o.teams = $tl
              }
              [void]$L.Add($o)
            }
          } catch {}
          $it = $suz.GetNext()
        }
        Yaz-Cevap $ctx ('{"ok":true,"olaylar":' + $(if ($L.Count) { ConvertTo-Json -InputObject @($L) -Depth 4 -Compress } else { '[]' }) + '}') "application/json; charset=utf-8" 200
      }
      elseif ($yol -eq "/api/takvim/ac" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = $null; try { $it = (Ns).GetItemFromID([string]$g.id) } catch {}
        if ($it) { $it.Display($false); try { $it.GetInspector.Activate() } catch {}; Yaz-Json $ctx @{ ok = $true } } else { Yaz-Hata $ctx "Takvim ogesi bulunamadi." 404 }
      }

      # ---------- toplanti odalari: Outlook "All Rooms" listesi + doluluk (30 dk dilim) ----------
      # Doluluk her odanin Outlook serbest/mesgul bilgisinden okunur: 0 bos, 1 belirsiz, 2 dolu, 3 disarida.
      # ?bas=yyyy-MM-dd&gunSayi=N (1-62): gunde 48 dilim, N gun art arda (hafta 7, ay gorunumu 42).
      # Outlook tek cagrida ~28 gun verir; daha uzun aralik 28'er gunluk parcalarla alinir.
      elseif ($yol -eq "/api/oda" -and $metod -eq "GET") {
        $gun = [datetime]::Today
        try { $gun = [datetime]::ParseExact((Sorgu $ctx "bas"), "yyyy-MM-dd", $null) } catch {}
        $n = 7; try { $n = [math]::Min(62, [math]::Max(1, [int](Sorgu $ctx "gunSayi"))) } catch {}
        $L = New-Object System.Collections.ArrayList
        foreach ($ae in (Oda-Listesi)) {
          $sb = New-Object System.Text.StringBuilder
          for ($p = 0; $p -lt $n; $p += 28) {
            $parca = [math]::Min(28, $n - $p)
            $fb = ""; try { $fb = [string]$ae.GetFreeBusy($gun.AddDays($p), 30, $true) } catch {}
            [void]$sb.Append($fb.PadRight($parca * 48, '0').Substring(0, $parca * 48))
          }
          $ep = ""; try { $ep = [string]$ae.GetExchangeUser().PrimarySmtpAddress } catch {}
          [void]$L.Add(@{ ad = [string]$ae.Name; eposta = $ep; dilim = $sb.ToString() })
        }
        Yaz-Json $ctx @{ ok = $true; bas = $gun.ToString("yyyy-MM-dd"); gunSayi = $n; odalar = @($L) }
      }
      # Rezervasyon: oda KAYNAK, katilimcilar ZORUNLU alici olarak eklenir ve davet GONDERILIR
      # (kullanici "Rezerve et" ile istedi, 13.10.2026: ilgililere posta gitsin). Oda posta kutusu
      # daveti kabul edince rezervasyon kesinlesir. gonder=false gelirse yalniz taslak acilir.
      elseif ($yol -eq "/api/oda/rezerve" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $oda = $null
        foreach ($ae in (Oda-Listesi)) { try { if ([string]$ae.GetExchangeUser().PrimarySmtpAddress -eq [string]$g.eposta) { $oda = $ae; break } } catch {} }
        try { $bas = [datetime]::ParseExact([string]$g.bas, "yyyy-MM-ddTHH:mm", $null); $bit = [datetime]::ParseExact([string]$g.bit, "yyyy-MM-ddTHH:mm", $null) } catch { $bas = $null }
        if (-not $oda) { Yaz-Hata $ctx "Oda bulunamadi." 400 }
        elseif (-not $bas -or $bit -le $bas) { Yaz-Hata $ctx "Saat araligi gecersiz." 400 }
        else {
          $t = (Outlook).CreateItem(1)          # olAppointmentItem
          $t.MeetingStatus = 1                  # olMeeting: davet olarak acilir
          $t.Subject = [string]$g.konu
          $t.Start = $bas; $t.End = $bit
          $t.Location = [string]$oda.Name
          $alici = $t.Recipients.Add([string]$g.eposta); $alici.Type = 3   # olResource
          $katilan = 0
          foreach ($ep in @($g.katilimcilar)) {
            $ep = [string]$ep
            if ($ep -match '^[^@\s;,]+@[^@\s;,]+\.[A-Za-z]{2,}$') { $r = $t.Recipients.Add($ep); $r.Type = 1; $katilan++ }   # olRequired
          }
          [void]$t.Recipients.ResolveAll()
          if ($g.gonder -eq $false) { $t.Display(); Yaz-Json $ctx @{ ok = $true; gonderildi = $false } }
          else { $t.Send(); Write-Host ("  Oda rezervasyonu gonderildi (" + $katilan + " katilimci)"); Yaz-Json $ctx @{ ok = $true; gonderildi = $true; katilimci = $katilan } }
        }
      }

      # ---------- Teams toplantisi (takvimdeki "+" menusu) ----------
      # Outlook'ta toplanti olusturulur, katilimcilar eklenir, pencere acilip Teams eklentisinin
      # "Teams Meeting" dugmesine UI Automation ile basilir (COM'da Teams baglantisi ekleme yolu yok;
      # 14.10.2026 denendi). Baglanti govdeye gelince davet GONDERILIR. Gelmezse gonderilmez,
      # pencere kullaniciya acik birakilir.
      elseif ($yol -eq "/api/toplanti/teams" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        try { $bas = [datetime]::ParseExact([string]$g.bas, "yyyy-MM-ddTHH:mm", $null); $bit = [datetime]::ParseExact([string]$g.bit, "yyyy-MM-ddTHH:mm", $null) } catch { $bas = $null }
        $epler = @(@($g.katilimcilar) | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^[^@\s;,]+@[^@\s;,]+\.[A-Za-z]{2,}$' })
        if (-not $bas -or $bit -le $bas) { Yaz-Hata $ctx "Saat araligi gecersiz." 400 }
        elseif (-not ([string]$g.konu).Trim()) { Yaz-Hata $ctx "Konu bos olamaz." 400 }
        elseif (-not $epler.Count) { Yaz-Hata $ctx "En az bir katilimci ekleyin." 400 }
        else {
          $t = (Outlook).CreateItem(1); $t.MeetingStatus = 1
          $t.Subject = ([string]$g.konu).Trim(); $t.Start = $bas; $t.End = $bit
          if ($g.aciklama) { $t.Body = [string]$g.aciklama }
          foreach ($ep in $epler) { $r = $t.Recipients.Add($ep); $r.Type = 1 }
          [void]$t.Recipients.ResolveAll()
          $t.Display($false)
          $link = Teams-Toplanti-Ekle $t
          if ($link) {
            $t.Send()
            Write-Host ("  Teams toplantisi gonderildi (" + $epler.Count + " katilimci)")
            Yaz-Json $ctx @{ ok = $true; gonderildi = $true; katilimci = $epler.Count }
          } else {
            Yaz-Hata $ctx "Teams baglantisi eklenemedi. Toplanti Outlook'ta acik birakildi: 'Teams Meeting' dugmesine basip Gonder'e basin." 409
          }
        }
      }

      # ---------- acilista birlikte baslayan uygulamalar (Outlook / MiCollab / Teams) ----------
      # Secim makineye ozel ayar dosyasinda (baslat) durur: bat okuyabilsin diye sifreli veride degil.
      # Bir uygulama acilinca kapaliysa hemen de baslatilir. Kurallar Sunucu\uygulamalar.ps1'de.
      elseif ($yol -eq "/api/baslat") {
        $vars = [ordered]@{ outlook = $true; micollab = $true; teams = $false }
        $simdi = [ordered]@{}
        foreach ($k in $vars.Keys) { $simdi[$k] = if ($ayar.baslat -and $null -ne $ayar.baslat.$k) { [bool]$ayar.baslat.$k } else { $vars[$k] } }
        if ($metod -eq "POST") {
          $g = (Govde-Oku $ctx) | ConvertFrom-Json
          foreach ($k in $vars.Keys) {
            if ($null -eq $g.$k) { continue }
            $yeniDeger = [bool]$g.$k
            if ($yeniDeger -and -not $simdi[$k] -and -not $env:ISTAKIP_PORT) { try { & (Join-Path $betikKok "uygulamalar.ps1") -Yalniz $k } catch {} }
            $simdi[$k] = $yeniDeger
          }
          $ayar | Add-Member -NotePropertyName baslat -NotePropertyValue ([pscustomobject]$simdi) -Force
          [System.IO.File]::WriteAllText($ayarDos, ($ayar | ConvertTo-Json -Depth 4), $UTF8)
        }
        Yaz-Json $ctx @{ ok = $true; baslat = $simdi }
      }

      # ---------- bildirim kontrolunu hemen calistir (Ayarlar > Dene) ----------
      elseif ($yol -eq "/api/bildirim/dene" -and $metod -eq "POST") {
        Bildirim-Goster $BLD_AD (U "Masa{00FC}st{00FC} bildirimi {00E7}al{0131}{015F}{0131}yor.") "ana"
        Yaz-Json $ctx @{ ok = $script:bildirimHazir }
      }
      # ---------- posta listesi ----------
      elseif ($yol -eq "/api/posta/liste") {
        $kaynak = Sorgu $ctx "kaynak"; $kategori = Sorgu $ctx "kategori"; $klasorYol = Sorgu $ctx "klasor"
        $gun = [int](Sorgu $ctx "gun"); if ($gun -le 0) { $gun = 14 }
        $enCok = 300
        try { $f = Klasor $klasorYol } catch { Yaz-Hata $ctx ("Klasor bulunamadi: " + $klasorYol) 404; $ctx.Response.OutputStream.Close(); continue }
        $items = $f.Items
        # "[FlagStatus] = 2" bu posta kutusunda HIC sonuc dondurmuyor (13.10.2026,
        # Exchange); "<> 0" calisiyor. Tamamlanmis bayraklar (1) asagida elenir.
        if ($kaynak -eq "bayrakli") { $items = $items.Restrict("[FlagStatus] <> 0") }
        elseif ($kaynak -eq "kategori" -and $kategori) { $items = $items.Restrict("[Categories] = '" + $kategori.Replace("'", "''") + "'") }
        $items.Sort("[ReceivedTime]", $true)
        $sinir = (Get-Date).Date.AddDays(-$gun)
        $liste = New-Object System.Collections.ArrayList
        $n = $items.Count
        for ($i = 1; $i -le $n -and $liste.Count -lt $enCok; $i++) {
          $it = $items.Item($i)
          if ($it.Class -ne 43) { continue }
          if ($kaynak -eq "bayrakli" -and $it.FlagStatus -ne 2) { continue }
          $alindi = [datetime]$it.ReceivedTime
          # Bayrakli postalar da "son N gun" ile sinirli: posta kutusunda yillar once
          # bayrak konmus yuzlerce posta var, listeyi onlar doldurmasin (13.10.2026).
          # Kategori kipinde sinir yok: kategori bilincli olarak islenmis bir isaret.
          if ($kaynak -ne "kategori" -and $alindi -lt $sinir) { break }
          $on = ""
          try { $on = ([string]$it.Body) -replace '\s+', ' '; if ($on.Length -gt 220) { $on = $on.Substring(0, 220) } } catch {}
          [void]$liste.Add(@{
            id = [string]$it.EntryID; store = [string]$f.StoreID
            konu = [string]$it.Subject; kimden = [string]$it.SenderName
            tarih = (Tarih-Iso $alindi); on = $on.Trim()
            bayrak = ($it.FlagStatus -eq 2); bayrakTermin = (Bayrak-Termin $it)
            okunmadi = [bool]$it.UnRead; ek = [int]$it.Attachments.Count
            kategori = [string]$it.Categories
          })
        }
        Yaz-Json $ctx @{ ok = $true; klasor = [string]$f.FolderPath; postalar = @($liste) }
      }

      # ---------- tek posta (tam metin) ----------
      elseif ($yol -eq "/api/posta/oku") {
        $it = Posta-Bul (Sorgu $ctx "id") (Sorgu $ctx "store") (Sorgu $ctx "msgId")
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else { Yaz-Json $ctx (Posta-Ayrinti $it) }
      }

      # ---------- Outlook'ta o an secili posta ----------
      # Kullanici postayi Outlook'ta secer, uygulamada "secili postadan doldur"
      # der: listede aramaya gerek kalmaz.
      elseif ($yol -eq "/api/posta/secili") {
        $ex = $null; try { $ex = (Outlook).ActiveExplorer() } catch {}
        $it = $null
        if ($ex) { try { if ($ex.Selection.Count -gt 0) { $it = $ex.Selection.Item(1) } } catch {} }
        # Posta ayri pencerede aciksa o pencerenin postasi alinir.
        if (-not $it) { try { $ins = (Outlook).ActiveInspector(); if ($ins) { $it = $ins.CurrentItem } } catch {} }
        if (-not $it) { Yaz-Hata $ctx "Outlook'ta secili posta yok." 404 }
        elseif ($it.Class -ne 43) { Yaz-Hata $ctx "Secili oge posta degil." 400 }
        else { Yaz-Json $ctx (Posta-Ayrinti $it) }
      }

      # ---------- postayi kayda bagla: .msg kopyasi + istege bagli kategori ----------
      elseif ($yol -eq "/api/posta/bagla" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = Posta-Bul ([string]$g.id) ([string]$g.store) ([string]$g.msgId)
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else {
          $msgAd = ""
          if ($g.msgSakla) {
            # Dosya adi yalniz rastgele kimlik: konu/kisi adi dosya sistemine dusmesin.
            $msgAd = ([guid]::NewGuid().ToString("N")) + ".msg"
            # Outlook dosyaya yazabilsin diye once gecici klasore, sonra sifrelenip posta\ altina.
            $gec = Join-Path $env:TEMP ("istakip_" + $msgAd)
            try {
              $it.SaveAs($gec, 9)
              Bayt-Yaz (Join-Path $postaKok ($msgAd + ".dat")) (Sifrele ([System.IO.File]::ReadAllBytes($gec)))
            } catch { $msgAd = "" }
            Remove-Item -LiteralPath $gec -Force -ErrorAction SilentlyContinue
          }
          $katSonuc = ""
          if ([string]$g.kategori) {
            try {
              $mevcut = [string]$it.Categories
              $parcalar = @($mevcut.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
              if ($parcalar -notcontains [string]$g.kategori) { $it.Categories = (($parcalar + [string]$g.kategori) -join ', '); $it.Save() }
              $katSonuc = "ok"
            } catch { $katSonuc = "hata" }
          }
          Yaz-Json $ctx @{ ok = $true; msg = $msgAd; kategori = $katSonuc }
        }
      }

      # ---------- bayrak ekle / kaldir (Outlook'taki postaya islenir) ----------
      # ---------- Gelen Kutusu: tumunu okundu yap (Posta Kutusu sayfasindaki dugme) ----------
      # Yalniz Gelen Kutusu'nun okunmamislari; posta silinmez/tasinmaz, yalniz "okundu" isaretlenir.
      elseif ($yol -eq "/api/posta/okundu" -and $metod -eq "POST") {
        $gk = (Ns).GetDefaultFolder(6); $n = 0
        $liste = @(); foreach ($it in $gk.Items.Restrict("[UnRead] = True")) { $liste += $it }
        foreach ($it in $liste) { try { $it.UnRead = $false; $n++ } catch {} }
        Write-Host ("  Posta: " + $n + " okunmamis posta okundu yapildi")
        Yaz-Json $ctx @{ ok = $true; sayi = $n; kalan = [int]$gk.UnReadItemCount }
      }
      elseif ($yol -eq "/api/posta/bayrak" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = Posta-Bul ([string]$g.id) ([string]$g.store) ([string]$g.msgId)
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else {
          # 4 = olMarkNoDate: tarihsiz bayrak. Termin uygulamada tutulur.
          if ($g.ac) { $it.MarkAsTask(4) } else { $it.ClearTaskFlag() }
          $it.Save()
          Yaz-Json $ctx @{ ok = $true; bayrak = ($it.FlagStatus -eq 2); bayrakTermin = (Bayrak-Termin $it) }
        }
      }
      # ---------- talebi bir kisiye ilet ----------
      # Outlook'ta HAZIR posta acilir; GONDERILMEZ. Kullanici kontrol edip kendisi gonderir.
      # Talebe bagli posta varsa o posta iletilir (Forward), yoksa yeni posta olusur.
      elseif ($yol -eq "/api/posta/ilet" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $kaynak = $null
        if ($g.posta -and $g.posta.id) { try { $kaynak = Posta-Bul ([string]$g.posta.id) ([string]$g.posta.store) ([string]$g.posta.msgId) } catch {} }
        if ($kaynak) { $m = $kaynak.Forward() } else { $m = (Outlook).CreateItem(0); $m.Subject = [string]$g.konu }
        $m.To = [string]$g.kime
        $metin = [string]$g.govde
        if ($m.BodyFormat -eq 2) {
          # HTML postada metin govdenin basina, mevcut bicim bozulmadan eklenir
          $ek = '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' +
                ([System.Security.SecurityElement]::Escape($metin) -replace "`r?`n", "<br>") + '</div><br>'
          $html = [string]$m.HTMLBody
          $bm = [regex]::Match($html, '<body[^>]*>', 'IgnoreCase')
          if ($bm.Success) { $m.HTMLBody = $html.Insert($bm.Index + $bm.Length, $ek) } else { $m.HTMLBody = $ek + $html }
        } else {
          $m.Body = $metin + "`r`n`r`n" + [string]$m.Body
        }
        $m.Display($false)
        try { $m.GetInspector.Activate() } catch {}
        Yaz-Json $ctx @{ ok = $true; iletildi = [bool]$kaynak }
      }

      # ---------- talebi baska bir Assist kullanicisina gonder (Exchange) ----------
      # Kullanici "Gonder"e basinca ileti GERCEKTEN gonderilir (talep iletme 13.10.2026'dan
      # beri boyle; once yalniz taslak aciliyordu). Bagli posta varsa ek olarak eklenir.
      elseif ($yol -eq "/api/talep/gonder" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $kime = [string]$g.kime
        if ($kime -notmatch '^[^@\s;,]+@[^@\s;,]+\.[A-Za-z]{2,}$') { Yaz-Hata $ctx "Gecersiz e-posta adresi." 400 }
        else {
          $b64 = [Convert]::ToBase64String($UTF8.GetBytes(($g.talep | ConvertTo-Json -Depth 6 -Compress)))
          $m = (Outlook).CreateItem(0)
          $m.MessageClass = $GELEN_SINIF
          $m.To = $kime
          $m.Subject = "[Assist] " + [string]$g.konu
          $m.Body = [string]$g.govde + "`r`n`r`n-----`r`n" + (U "Bu talep Assist ile g{00F6}nderildi; Assist kullan{0131}yorsan{0131}z portal{0131}n{0131}za kendili{011F}inden d{00FC}{015F}er.") + "`r`n" + $VERI_ISARET + $b64 + ":SON"
          if ($g.posta -and $g.posta.id) {
            try { $kaynak = Posta-Bul ([string]$g.posta.id) ([string]$g.posta.store) ([string]$g.posta.msgId); if ($kaynak) { [void]$m.Attachments.Add($kaynak, 5) } } catch {}
          }
          [void]$m.Recipients.ResolveAll()
          $m.Send()
          Write-Host "  Talep gonderildi (Assist iletisi)"
          Yaz-Json $ctx @{ ok = $true }
        }
      }
      # Gelen talep kuyrugu: sayfa alir, kayda donusturur, sonra "al" ile kuyruktan siler
      elseif ($yol -eq "/api/gelen" -and $metod -eq "GET") {
        $kuyruk = @(Gelen-Oku)
        Yaz-Cevap $ctx ('{"ok":true,"gelen":' + $(if ($kuyruk.Count) { ConvertTo-Json -InputObject $kuyruk -Depth 8 -Compress } else { '[]' }) + '}') "application/json; charset=utf-8" 200
      }
      elseif ($yol -eq "/api/gelen/al" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $alinan = @($g.ids | ForEach-Object { [string]$_ })
        Gelen-Yaz @(Gelen-Oku | Where-Object { $alinan -notcontains [string]$_.id })
        Yaz-Json $ctx @{ ok = $true }
      }

      # ---------- postayi Outlook'ta ac ----------
      elseif ($yol -eq "/api/posta/ac" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = $null
        try { $it = Posta-Bul ([string]$g.id) ([string]$g.store) ([string]$g.msgId) } catch {}
        if ($it) {
          $it.Display($false)
          try { $it.GetInspector.Activate() } catch {}
          Yaz-Json $ctx @{ ok = $true; kaynak = "outlook" }
        } else {
          # Posta kutusunda yoksa (silinmis/arsivlenmis) saklanan .msg kopyasi acilir.
          $ad = [string]$g.msg
          # Sifreli kopya gecici klasore cozulup acilir; 1 saat sonra silinir (Gecici-Temizle).
          $sifreli = if ($ad -match '^[0-9a-f]{32}\.msg$') { Join-Path $postaKok ($ad + ".dat") } else { "" }
          if ($sifreli -and (Test-Path -LiteralPath $sifreli)) {
            if (-not (Test-Path -LiteralPath $acKok)) { New-Item -ItemType Directory -Path $acKok | Out-Null }
            $yolMsg = Join-Path $acKok $ad
            [System.IO.File]::WriteAllBytes($yolMsg, (Coz ([System.IO.File]::ReadAllBytes($sifreli))))
            Start-Process -FilePath $yolMsg
            Yaz-Json $ctx @{ ok = $true; kaynak = "kopya" }
          } else { Yaz-Hata $ctx "Posta Outlook'ta bulunamadi ve kopyasi yok." 404 }
        }
      }
      else { Yaz-Hata $ctx "Bilinmeyen uc." 404 }
      $ctx.Response.OutputStream.Close(); continue
    }

    # ---------- statik dosyalar ----------
    if ($yol -eq "/") { $yol = "/is_takip.html" }
    $tam = [System.IO.Path]::GetFullPath((Join-Path $webKok $yol.TrimStart('/')))
    if ($yol -eq "/favicon.ico") { $tam = Join-Path $portalKok "Assist.ico" }   # uygulama penceresinin gorev cubugu simgesi
    if (-not ($tam.StartsWith($webKok, [StringComparison]::OrdinalIgnoreCase) -or $yol -eq "/favicon.ico") -or -not (Test-Path -LiteralPath $tam -PathType Leaf)) {
      Yaz-Cevap $ctx "Bulunamadi" "text/plain; charset=utf-8" 404
    } else {
      $b = [System.IO.File]::ReadAllBytes($tam)
      $ctx.Response.ContentType = Icerik-Turu ([System.IO.Path]::GetExtension($tam))
      $ctx.Response.Headers.Add("Cache-Control", "no-store")
      $ctx.Response.ContentLength64 = $b.Length
      $ctx.Response.OutputStream.Write($b, 0, $b.Length)
    }
    $ctx.Response.OutputStream.Close()
  } catch {
    $m = $_.Exception.Message
    Write-Host ("  HATA " + $yol + " : " + $m) -ForegroundColor Yellow
    try { Yaz-Hata $ctx $m 500 } catch {}
    try { $ctx.Response.OutputStream.Close() } catch {}
  }
}
