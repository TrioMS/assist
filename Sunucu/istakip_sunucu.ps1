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
#  GUVENLIK (15.10.2026): /api uclari ayrica OTURUM ANAHTARI ister
#  (X-Assist-Anahtar); /__saglik anahtarla kanit uretir; imza zorunlu
#  kipte imzasiz kod yeniden yuklenmez, dogrulanamayan sayfa sunulmaz.
#  Ayrinti: Sunucu\guvenlik.ps1
#
#  NOT: Bu dosya bilincli olarak ASCII'dir (PowerShell 5.1 BOM'suz UTF-8
#  dosyayi yanlis okur). Turkce metinler portal tarafinda durur.
# =====================================================================
$ErrorActionPreference = "Stop"
$SURUM = "0.1.0"

$betikKok = $PSScriptRoot
$BETIK_ZAMAN = (Get-Item -LiteralPath $PSCommandPath).LastWriteTimeUtc.Ticks
. (Join-Path $betikKok "guvenlik.ps1")
$GUV = Assist-Guvenlik-Ayar
# Oturum anahtari: sunucu kendini yeniden yukleyince (ayni surec) ayni anahtar korunur.
if ($global:AssistAnahtar -and (Assist-Anahtar-Oku) -eq $global:AssistAnahtar) { $ANAHTAR = $global:AssistAnahtar }
else { $ANAHTAR = Assist-Anahtar-Hazirla; $global:AssistAnahtar = $ANAHTAR }
$script:katalogAnahtar = ""; $script:katalogOk = $false   # sayfa katalog dogrulamasi onbellegi

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
# Bulgu 3 (15.10.2026): acilmak icin cozulen dosyaya "internetten geldi" isareti (Mark-of-the-Web, Zone 3)
# konur. Isaretsiz dosyada Office Korumali Gorunum'u ve internet makrolarini engelleme kuralini uygulamiyordu.
# Outlook'un ekleri actigi kendi gecici klasoru (OLK / Content.Outlook). Assist veri klasorune ek YAZILMAZ
# (kullanici istedi, 16.10.2026): ek Outlook'taki postadan alinir, Outlook'un klasorunde acilir.
function Outlook-Ek-Klasoru {
  try {
    $v = ([string](Outlook).Version).Split('.')[0] + ".0"
    $k = (Get-ItemProperty -Path ("HKCU:\Software\Microsoft\Office\" + $v + "\Outlook\Security") -Name OutlookSecureTempFolder -ErrorAction Stop).OutlookSecureTempFolder
    if ($k -and (Test-Path -LiteralPath $k)) { return $k }
  } catch {}
  if (-not (Test-Path -LiteralPath $acKok)) { New-Item -ItemType Directory -Path $acKok | Out-Null }
  return $acKok
}
function Internet-Isareti([string]$yol) {
  try { Set-Content -LiteralPath $yol -Stream "Zone.Identifier" -Value "[ZoneTransfer]`r`nZoneId=3" -Encoding ASCII -ErrorAction Stop }
  catch { Write-Host ("  Internet isareti konamadi: " + $_.Exception.Message) -ForegroundColor Yellow }
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
  # Bulgu 10 (15.10.2026) + G1 (16.10.2026): govde ustten sinirli okunur. ContentLength64'e GUVENILMEZ:
  # chunked (Transfer-Encoding: chunked) istekte -1 gelir ve eski kontrol atlanirdi; burada gercek
  # okunan bayt sayisi sinirlanir.
  $sinir = 52428800
  if ($ctx.Request.ContentLength64 -gt $sinir) { throw "Istek govdesi cok buyuk." }
  $tampon = New-Object byte[] 65536
  $ms = New-Object System.IO.MemoryStream
  $gs = $ctx.Request.InputStream
  try {
    while (($o = $gs.Read($tampon, 0, $tampon.Length)) -gt 0) {
      $ms.Write($tampon, 0, $o)
      if ($ms.Length -gt $sinir) { throw "Istek govdesi cok buyuk." }
    }
    return $UTF8.GetString($ms.ToArray())
  } finally { $ms.Dispose() }
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
    ".wasm" { "application/wasm" }
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
# Outlook baglantisi onbellekte tutulur. Outlook kapanip yeniden acilinca eski baglanti OLU kalir ("RPC server
# is unavailable"); .Version okumak bunu yakalamiyordu (15.10.2026: sayfa Yenile'ye basmadan postalari ve
# toplantilari getirmiyordu). Gercek bir cagri (GetNamespace) denenir; olmazsa birakilip yeniden baglanilir.
function Outlook {
  if ($script:ol) {
    try { $null = $script:ol.GetNamespace("MAPI"); return $script:ol }
    catch { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:ol) } catch {}; $script:ol = $null }
  }
  $script:ol = New-Object -ComObject Outlook.Application
  return $script:ol
}
function Ns { return (Outlook).GetNamespace("MAPI") }
# Arka plan kontrolleri (tepsi ozeti, posta sayaci, dakikalik zamanlayici) Outlook'a YALNIZ Outlook acik ve
# en az 90 sn'dir calisiyorsa dokunur (15.10.2026): acilis aninda COM ikinci bir Outlook baslatip "another
# program is using Outlook" + 0xc0000142 hatasina, Outlook guvenlik uyarisina (Allow) ve sunucunun
# kilitlenmesine (sayfa bos kaldi) yol aciyordu. Kullanicinin tikladigi istekler bu kurala bagli degil.
function Outlook-Uygun {
  $p = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) | Select-Object -First 1
  if (-not $p) { return $false }
  # 16.10.2026: 90 sn beklemek acilista yeni postalarin sayaca ~2 dk gec dusmesine yol aciyordu. Outlook
  # "yuklendi" kaydini (Application gunlugu, olay 45) yazdiysa hazirdir: acilis ekrani da bunu bekler.
  $yas = 0; try { $yas = ((Get-Date) - $p.StartTime).TotalSeconds } catch { return $false }
  if ($yas -ge 90) { return $true }
  if ($yas -lt 10) { return $false }
  if ($script:olHazirPid -eq $p.Id) { return $true }   # bu Outlook sureci icin bir kez dogrulandi
  $ev = Get-WinEvent -FilterHashtable @{ LogName = "Application"; ProviderName = "Outlook"; Id = 45; StartTime = $p.StartTime } -MaxEvents 1 -ErrorAction SilentlyContinue
  if ($ev) { $script:olHazirPid = $p.Id; return $true }
  return $false
}
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
# --- Ceviri ve ses (16.10.2026): Microsoft Translator + Azure Konusma (sirketin Azure hesabi) ---
# Kullanici karari: icerik yalniz sirketin kendi Azure aboneligindeki Microsoft hizmetine gider (ucretsiz genel
# hizmetler KULLANILMAZ). Anahtarlar ceviri.dat'ta DPAPI ile sifreli; sayfaya GERI VERILMEZ (yalniz "tanimli mi").
# Sesli okuma ve ozet tarayicida/cevrimdisi yapilir, buraya gelmez.
$ceviriDos = Join-Path $veriKok "ceviri.dat"
function Ceviri-Ayar {
  if (Test-Path -LiteralPath $ceviriDos) { try { return (Sifreli-Oku $ceviriDos) | ConvertFrom-Json } catch {} }
  return [pscustomobject]@{ anahtar = ""; bolge = ""; konusmaAnahtar = ""; konusmaBolge = "" }
}
function Tls12 { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 }
# Metni hedef dile cevirir. Doner: @{ ceviri; dil (algilanan kaynak dil) }. Uzun metin paragraf sinirinda parcalanir.
function Ceviri-Yap([string]$metin, [string]$hedef) {
  $a = Ceviri-Ayar
  if (-not [string]$a.anahtar) { throw "Ceviri anahtari tanimli degil (Ayarlar > Ceviri ve Ses)." }
  if ($hedef -notmatch '^[A-Za-z]{2,3}(-[A-Za-z]{2,4})?$') { throw "Gecersiz hedef dil." }
  Tls12
  $h = @{ "Ocp-Apim-Subscription-Key" = [string]$a.anahtar }; if ([string]$a.bolge) { $h["Ocp-Apim-Subscription-Region"] = [string]$a.bolge }
  $parcalar = New-Object System.Collections.ArrayList; $sb = New-Object System.Text.StringBuilder
  foreach ($satir in ($metin -split "`n")) {
    if ($sb.Length + $satir.Length -gt 9000 -and $sb.Length) { [void]$parcalar.Add($sb.ToString()); [void]$sb.Clear() }
    [void]$sb.Append($satir).Append("`n")
  }
  if ($sb.Length) { [void]$parcalar.Add($sb.ToString()) }
  $govde = ConvertTo-Json -InputObject @($parcalar | ForEach-Object { @{ Text = $_ } }) -Compress   # -InputObject: tek parcada da dizi
  $r = Invoke-RestMethod -Uri ("https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&to=" + $hedef) -Method Post -Headers $h `
    -ContentType "application/json; charset=utf-8" -Body $UTF8.GetBytes($govde) -TimeoutSec 40
  $cikti = New-Object System.Text.StringBuilder; $dil = ""
  foreach ($x in @($r)) { [void]$cikti.Append([string]$x.translations[0].text); if (-not $dil -and $x.detectedLanguage) { $dil = [string]$x.detectedLanguage.language } }
  return @{ ceviri = $cikti.ToString().TrimEnd(); dil = $dil }
}
# Kisa ses kaydini (WAV, 16 kHz mono PCM, en cok ~60 sn) yaziya cevirir (Azure Konusma, REST).
function Konusma-Yaz([byte[]]$wav, [string]$dil) {
  $a = Ceviri-Ayar
  $an = $(if ([string]$a.konusmaAnahtar) { [string]$a.konusmaAnahtar } else { [string]$a.anahtar })
  $bl = $(if ([string]$a.konusmaBolge) { [string]$a.konusmaBolge } else { [string]$a.bolge })
  if (-not $an -or -not $bl) { throw "Konusma anahtari ya da bolgesi tanimli degil (Ayarlar > Ceviri ve Ses)." }
  if ($bl -notmatch '^[a-z0-9]+$') { throw "Gecersiz bolge." }
  if ($dil -notmatch '^[a-z]{2}-[A-Z]{2}$') { $dil = "tr-TR" }
  Tls12
  $h = @{ "Ocp-Apim-Subscription-Key" = $an; "Accept" = "application/json" }
  $r = Invoke-RestMethod -Uri ("https://" + $bl + ".stt.speech.microsoft.com/speech/recognition/conversation/cognitiveservices/v1?language=" + $dil + "&format=simple") `
    -Method Post -Headers $h -ContentType "audio/wav; codecs=audio/pcm; samplerate=16000" -Body $wav -TimeoutSec 60
  if ([string]$r.RecognitionStatus -ne "Success") { throw ("Ses anlasilamadi (" + [string]$r.RecognitionStatus + ").") }
  return [string]$r.DisplayText
}
# Metni sese cevirir (Azure Konusma, noral ses). Bilgisayarda o dilin Windows sesi yoksa sayfa bunu kullanir. Doner: mp3 baytlari.
$SES_ADI = @{ tr = "tr-TR-EmelNeural"; en = "en-US-JennyNeural"; de = "de-DE-KatjaNeural"; fr = "fr-FR-DeniseNeural"; ar = "ar-SA-ZariyahNeural"; ru = "ru-RU-SvetlanaNeural"; es = "es-ES-ElviraNeural"; it = "it-IT-ElsaNeural" }
function Ses-Sentez([string]$metin, [string]$dil) {
  $a = Ceviri-Ayar
  $an = $(if ([string]$a.konusmaAnahtar) { [string]$a.konusmaAnahtar } else { [string]$a.anahtar })
  $bl = $(if ([string]$a.konusmaBolge) { [string]$a.konusmaBolge } else { [string]$a.bolge })
  if (-not $an -or -not $bl) { throw "Konusma anahtari ya da bolgesi tanimli degil (Ayarlar > Ceviri ve Ses)." }
  if ($bl -notmatch '^[a-z0-9]+$') { throw "Gecersiz bolge." }
  $k = ([string]$dil).ToLowerInvariant(); if ($k.Length -gt 2) { $k = $k.Substring(0, 2) }
  $ses = $(if ($SES_ADI.ContainsKey($k)) { $SES_ADI[$k] } else { $SES_ADI["en"] }); $yerel = $ses.Substring(0, 5)
  $ssml = "<speak version='1.0' xml:lang='" + $yerel + "'><voice name='" + $ses + "'>" + [System.Security.SecurityElement]::Escape($metin) + "</voice></speak>"
  Tls12
  $h = @{ "Ocp-Apim-Subscription-Key" = $an; "X-Microsoft-OutputFormat" = "audio-24khz-48kbitrate-mono-mp3"; "User-Agent" = "Assist" }
  $gec = Join-Path $env:TEMP ("assist_ses_" + [guid]::NewGuid().ToString("N") + ".mp3")
  try {
    Invoke-WebRequest -Uri ("https://" + $bl + ".tts.speech.microsoft.com/cognitiveservices/v1") -Method Post -Headers $h -ContentType "application/ssml+xml; charset=utf-8" `
      -Body $UTF8.GetBytes($ssml) -OutFile $gec -UseBasicParsing -TimeoutSec 60 | Out-Null
    $b = [System.IO.File]::ReadAllBytes($gec)
  } finally { Remove-Item -LiteralPath $gec -Force -ErrorAction SilentlyContinue }
  return ,$b
}
# --- Posta Kutusu klasor agaci (16.10.2026): Outlook'un klasor bolmesinin aynasi ---
# Her depo (kendi posta kutusu, ek/paylasilan posta kutulari, PST "Kisisel Klasorler") kok olur; altinda
# yalniz posta klasorleri (DefaultItemType 0), gizli olanlar (PR_ATTR_HIDDEN) ve Esitleme Sorunlari haric.
# Varsayilan klasorler Outlook sirasiyla once, digerleri ada gore. Ortak klasor deposu (public) alinmaz.
$PR_GIZLI = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B"
function Posta-Agac {
  $tr = [Globalization.CultureInfo]::GetCultureInfo("tr-TR")
  $sayac = @{ n = 0 }
  $depolar = New-Object System.Collections.ArrayList
  foreach ($s in (Ns).Stores) {
    try { if ([int]$s.ExchangeStoreType -eq 2) { continue } } catch {}
    $ozel = @{}
    foreach ($p in @(@(6, "gelen"), @(16, "taslak"), @(5, "gonderilmis"), @(3, "silinmis"), @(23, "onemsiz"), @(4, "giden"), @(20, "esitleme"))) {
      try { $ozel[[string]$s.GetDefaultFolder($p[0]).EntryID] = $p[1] } catch {}
    }
    $sira = @{ gelen = 0; taslak = 1; gonderilmis = 2; silinmis = 3; arsiv = 4; onemsiz = 5; giden = 6 }
    $dal = $null
    $dal = {
      param($kf, [int]$derin)
      $L = New-Object System.Collections.ArrayList
      if ($derin -gt 6) { return ,$L }
      foreach ($f in $kf.Folders) {
        if ($sayac.n -ge 500) { break }
        try {
          if ([int]$f.DefaultItemType -ne 0) { continue }
          $id = [string]$f.EntryID; $tur = [string]$ozel[$id]
          if ($tur -eq "esitleme") { continue }
          $gizli = $false; try { $gizli = [bool]$f.PropertyAccessor.GetProperty($PR_GIZLI) } catch {}
          if ($gizli) { continue }
          $ad = [string]$f.Name
          if (-not $tur -and $derin -eq 0 -and ($ad -eq "Archive" -or $ad -eq "Ar$([char]0x015F)iv")) { $tur = "arsiv" }
          $sayac.n++
          [void]$L.Add([ordered]@{ id = $id; ad = $ad; ok = [int]$f.UnReadItemCount; tur = $tur; alt = (& $dal $f ($derin + 1)) })
        } catch {}
      }
      $S = @($L | Sort-Object @{ e = { if ($sira.ContainsKey($_.tur)) { $sira[$_.tur] } else { 99 } } }, @{ e = { $_.ad.ToLower($tr) } })
      return ,$S
    }
    try {
      $kok = $s.GetRootFolder()
      [void]$depolar.Add([ordered]@{ ad = [string]$s.DisplayName; store = [string]$s.StoreID; birincil = ($s.StoreID -eq (Ns).DefaultStore.StoreID); klasorler = (& $dal $kok 0) })
    } catch {}
  }
  return ,@($depolar | Sort-Object @{ e = { if ($_.birincil) { 0 } else { 1 } } })
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
# --- ORTAK IS TAKIBI (15.10.2026) ----------------------------------------------------
# Konu + alt alta DEGISTIRILEMEZ kayit defteri (kim, ne zaman, ne yazdi); yalniz katilimcilar arasinda.
# Merkezi sunucu YOK: her katilimcinin Assist'inde tam kopya (ortak.dat, DPAPI). Esitleme Exchange ile:
# IPM.Note.Assist.Konu sinifinda ileti; govdede "ASSIST-KONU:<base64 JSON>:SON" = konu basligi + girisler.
# Alici girisleri kimlige gore birlestirir (ayni giris iki kez yazilmaz). Kurallar (kullanici, 15.10.2026):
#  - katilimci ekleme ve kapatma/acma YALNIZ konuyu acan (baslik/katilimci/durum degisikligi yalniz acandan kabul)
#  - bir giris yalniz yazarinin gonderdigi iletiyle (ya da acanin "gecmisi paylas"iyla) kabul edilir
#  - ek dosya yok; yeni katilan katildigi andan sonrasini alir, acan isterse gecmisi gonderir
#  - Gonderilemeyen paket kuyrukta (giden) bekler, dakikalik kontrolde yeniden denenir
# Simdilik icerik Exchange'de base64 (sifresiz) durur; uctan uca sifreleme sonraki adim.
$ORTAK_SINIF = "IPM.Note.Assist.Konu"
$ORTAK_ISARET = "ASSIST-KONU:"
$ortakDos = Join-Path $veriKok "ortak.dat"
$script:ortakRev = 0
function Ortak-Oku {
  $o = $null
  if (Test-Path -LiteralPath $ortakDos) { try { $o = (Sifreli-Oku $ortakDos) | ConvertFrom-Json } catch {} }
  if (-not $o) { $o = [pscustomobject]@{ konular = @(); giden = @() } }
  if ($null -eq $o.giden) { $o | Add-Member -NotePropertyName giden -NotePropertyValue @() -Force }
  $o.konular = @($o.konular | ForEach-Object { $_ }); $o.giden = @($o.giden | ForEach-Object { $_ })
  foreach ($k in $o.konular) { $k.girisler = @($k.girisler | Sort-Object { [string]$_.t }); $k.katilimcilar = @($k.katilimcilar | ForEach-Object { $_ }) }   # defter hep zaman sirasinda
  # Sifreleme (15.10.2026): bilinen acik anahtarlar, degisen anahtarlar (onay bekler), istenme zamanlari,
  # anahtari gelmemis aliciya bekleyen paketler, anahtari onaylanmamis gonderenden gelen paketler
  foreach ($ad in "anahtarlar", "degisen", "istenen") { $h = @{}; if ($o.$ad) { foreach ($p in $o.$ad.PSObject.Properties) { $h[$p.Name] = $p.Value } }; $o | Add-Member -NotePropertyName $ad -NotePropertyValue $h -Force }
  foreach ($ad in "bekleyenGiden", "bekleyenGelen", "bekleyenTalep") { $o | Add-Member -NotePropertyName $ad -NotePropertyValue @($o.$ad | Where-Object { $_ }) -Force }
  return $o
}
function Ortak-Yaz($o) { Sifreli-Yaz $ortakDos ($o | ConvertTo-Json -Depth 10 -Compress); $script:ortakRev++ }
function Ben-Kimlik {
  if ($script:benKimlik) { return $script:benKimlik }
  $ad = [string]$env:ISTAKIP_BEN_AD   # test sunucusu
  if (-not $ad) { try { Add-Type -AssemblyName System.DirectoryServices.AccountManagement; $ad = [string][System.DirectoryServices.AccountManagement.UserPrincipal]::Current.DisplayName } catch {} }
  if (-not $ad) { try { $ad = [string]((Veri-Oku) | ConvertFrom-Json).ayar.adSoyad } catch {} }
  if (-not $ad) { $ad = $env:USERNAME }
  $script:benKimlik = [pscustomobject]@{ ad = $ad; eposta = (Ben-Eposta) }
  return $script:benKimlik
}
function Ortak-Baslik($k) {
  return [ordered]@{ id = $k.id; baslik = $k.baslik; termin = $k.termin; acan = $k.acan; katilimcilar = @($k.katilimcilar); durum = $k.durum; olusturma = $k.olusturma }
}
# --- UCTAN UCA SIFRELEME (15.10.2026) -------------------------------------------------
# Her Assist'in RSA-3072 anahtar cifti var; gizli anahtar ortak_anahtar.dat'ta DPAPI ile (bilgisayardan
# cikmaz). Her konunun 64 baytlik anahtari (32 AES-256-CBC + 32 HMAC-SHA256) her alici icin onun acik
# anahtariyla RSA-OAEP ile ayri ayri kilitlenir; govde sifreli, HMAC'li ve gonderenin gizli anahtariyla
# imzali gider. Exchange'de icerik (baslik dahil) OKUNAMAZ; yalniz kimin kime ne zaman yazdigi gorunur.
# Acik anahtarlar ilk temasta "anahtar-iste"/"anahtar" iletileriyle degisilir (ilk gorulen kabul edilir);
# bir kisinin anahtari sonradan DEGISIRSE kullanici onaylayana kadar o kisiden giris kabul edilmez.
# Eski surumun sifresiz (v1) iletileri gecis icin hala alinir.
$ortakAnahtarDos = Join-Path $veriKok "ortak_anahtar.dat"
$script:benRsa = $null; $script:benAcik = ""
function Rsa-Bos { $p = New-Object System.Security.Cryptography.CspParameters 24; $r = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList $p; $r.PersistKeyInCsp = $false; return $r }
function Benim-Rsa {
  if ($script:benRsa) { return $script:benRsa }
  if (Test-Path -LiteralPath $ortakAnahtarDos) {
    $r = Rsa-Bos; $r.FromXmlString($UTF8.GetString((Coz ([System.IO.File]::ReadAllBytes($ortakAnahtarDos)))))
  } else {
    $p = New-Object System.Security.Cryptography.CspParameters 24
    $r = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 3072, $p; $r.PersistKeyInCsp = $false
    Bayt-Yaz $ortakAnahtarDos (Sifrele ($UTF8.GetBytes($r.ToXmlString($true))))
    Write-Host "  Ortak is sifreleme anahtari olusturuldu"
  }
  $script:benRsa = $r; $script:benAcik = $r.ToXmlString($false)
  return $r
}
function Benim-Acik { $null = Benim-Rsa; return $script:benAcik }
function Rsa-Acik([string]$xml) { $r = Rsa-Bos; $r.FromXmlString($xml); return $r }
function Parmak([string]$acik) {
  $h = [System.Security.Cryptography.SHA256]::Create().ComputeHash($UTF8.GetBytes($acik))
  return (($h[0..7] | ForEach-Object { $_.ToString("X2") }) -join "") -replace '(.{4})(?!$)', '$1 '
}
function B64([byte[]]$b) { return [Convert]::ToBase64String($b) }
function Bayt-Esit([byte[]]$a, [byte[]]$b) { if ($a.Length -ne $b.Length) { return $false }; $f = 0; for ($i = 0; $i -lt $a.Length; $i++) { $f = $f -bor ($a[$i] -bxor $b[$i]) }; return ($f -eq 0) }
function Aes-Sifrele([byte[]]$ak, [byte[]]$veri) {
  $aes = [System.Security.Cryptography.AesCryptoServiceProvider]::new(); $aes.KeySize = 256; $aes.Mode = "CBC"; $aes.Padding = "PKCS7"
  $aes.Key = [byte[]]($ak[0..31]); $aes.GenerateIV()
  $sif = $aes.CreateEncryptor().TransformFinalBlock($veri, 0, $veri.Length)
  $mac = [System.Security.Cryptography.HMACSHA256]::new([byte[]]($ak[32..63])).ComputeHash([byte[]]($aes.IV + $sif))
  return @{ iv = (B64 $aes.IV); sif = (B64 $sif); mac = (B64 $mac) }
}
function Aes-Coz([byte[]]$ak, [string]$iv, [string]$sif, [string]$mac) {
  $ivB = [Convert]::FromBase64String($iv); $sB = [Convert]::FromBase64String($sif)
  $beklenen = [System.Security.Cryptography.HMACSHA256]::new([byte[]]($ak[32..63])).ComputeHash([byte[]]($ivB + $sB))
  if (-not (Bayt-Esit $beklenen ([Convert]::FromBase64String($mac)))) { throw "HMAC tutmadi (ileti degistirilmis ya da anahtar yanlis)" }
  $aes = [System.Security.Cryptography.AesCryptoServiceProvider]::new(); $aes.KeySize = 256; $aes.Mode = "CBC"; $aes.Padding = "PKCS7"
  $aes.Key = [byte[]]($ak[0..31]); $aes.IV = $ivB
  return ,$aes.CreateDecryptor().TransformFinalBlock($sB, 0, $sB.Length)
}
function Imza-Verisi($z) { return $UTF8.GetBytes([string]$z.konuId + "|" + [string]$z.iv + "|" + [string]$z.sif + "|" + [string]$z.mac + "|" + ([string]$z.gonderen).ToLowerInvariant()) }
# Kisinin acik anahtarini ogrenir. Doner: yeni | ayni | degisti (degisti: onay icin bekletilir)
function Anahtar-Ogren($o, [string]$eposta, [string]$acik) {
  $eposta = $eposta.ToLowerInvariant()
  if (-not $acik -or $acik -notmatch '^<RSAKeyValue>') { return "gecersiz" }
  $var = $o.anahtarlar[$eposta]
  if (-not $var) { $o.anahtarlar[$eposta] = [pscustomobject]@{ acik = $acik; parmak = (Parmak $acik); t = (Get-Date).ToString("o") }; return "yeni" }
  if ([string]$var.acik -eq $acik) { return "ayni" }
  $o.degisen[$eposta] = [pscustomobject]@{ acik = $acik; parmak = (Parmak $acik); t = (Get-Date).ToString("o") }
  return "degisti"
}
# Tasiyici ileti. Konu satirinda baslik YOK (sifreli icerikte). Test sunucusu: ISTAKIP_TEST_POSTA klasorune dosya.
function Ortak-Ileti([string[]]$kime, [string]$json, [string]$baslik) {
  $govde = (U "Bu ileti Assist ortak i{015F} takibi i{00E7}indir; Assist kullan{0131}yorsan{0131}z kendili{011F}inden i{015F}lenir.") + "`r`n" + $ORTAK_ISARET + [Convert]::ToBase64String($UTF8.GetBytes($json)) + ":SON"
  if ($env:ISTAKIP_TEST_POSTA) {
    foreach ($a in $kime) { [System.IO.File]::WriteAllText((Join-Path $env:ISTAKIP_TEST_POSTA ([guid]::NewGuid().ToString("N") + ".json")), (@{ kime = $a; gonderen = (Ben-Eposta); govde = $govde } | ConvertTo-Json -Compress), $UTF8) }
    return
  }
  if ($env:ISTAKIP_PORT) { Write-Host ("  [test] ortak is iletisi GONDERILMEDI -> " + ($kime -join ", ")); return }
  $m = (Outlook).CreateItem(0)
  $m.MessageClass = $ORTAK_SINIF
  $m.To = ($kime -join "; ")
  $m.Subject = "[Assist] " + (U "Ortak i{015F}: g{00FC}venli ileti")
  $m.Body = $govde
  $m.DeleteAfterSubmit = $true   # Gonderilmis Ogeler dolmasin
  # Gonderenin Outlook'unda "her iletiye teslim/okundu bildirimi iste" acik olabilir: tasiyici iletiye
  # bildirim istenmez (15.10.2026: gonderene "Delivered" postasi geliyordu)
  $m.OriginatorDeliveryReportRequested = $false; $m.ReadReceiptRequested = $false
  [void]$m.Recipients.ResolveAll()
  $m.Send()
}
# Hazir JSON'u yollar; Outlook'a ulasilamazsa kuyruga (giden) birakir.
function Ortak-Teslim($o, [string[]]$kime, $zarf) {
  $json = $zarf | ConvertTo-Json -Depth 10 -Compress
  try { Ortak-Ileti $kime $json ""; return $true }
  catch {
    Write-Host ("  Ortak is gonderilemedi, kuyrukta: " + $_.Exception.Message) -ForegroundColor Yellow
    $o.giden = @($o.giden) + @([pscustomobject]@{ kime = @($kime); json = $json; baslik = ""; t = (Get-Date).ToString("o") })
    return $false
  }
}
function Konu-Anahtari($k) {
  if (-not $k.anahtar) { $b = New-Object byte[] 64; [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($b); $k | Add-Member -NotePropertyName anahtar -NotePropertyValue (B64 $b) -Force }
  return ,[Convert]::FromBase64String([string]$k.anahtar)
}
# Sifreli zarf: icerik (konu basligi + girisler) yalniz $alicilar (acik anahtari bilinen) acabilir
function Ortak-Zarf($o, $k, [string[]]$alicilar, $ic) {
  $ak = Konu-Anahtari $k
  $s = Aes-Sifrele $ak ($UTF8.GetBytes(($ic | ConvertTo-Json -Depth 10 -Compress)))
  $kilit = @{}; foreach ($e in $alicilar) { $kilit[$e] = B64 ((Rsa-Acik ([string]$o.anahtarlar[$e].acik)).Encrypt($ak, $true)) }
  $z = [ordered]@{ v = 2; tur = "veri"; konuId = $k.id; gonderen = (Ben-Eposta); acik = (Benim-Acik); anahtarlar = $kilit; iv = $s.iv; sif = $s.sif; mac = $s.mac }
  $z.imza = B64 ((Benim-Rsa).SignData((Imza-Verisi $z), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1))
  return $z
}
function Anahtar-Iste($o, [string]$e) {
  $son = $o.istenen[$e]
  if ($son -and ((Get-Date) - [datetime]::Parse([string]$son)).TotalMinutes -lt 10) { return }
  $o.istenen[$e] = (Get-Date).ToString("o")
  [void](Ortak-Teslim $o @($e) ([ordered]@{ v = 2; tur = "anahtar-iste"; gonderen = (Ben-Eposta); acik = (Benim-Acik) }))
}
# Anahtari yeni gelen kisiye bekleyen paketleri sifreleyip yollar
function Bekleyen-Gonder($o, [string]$e) {
  $kalan = @()
  foreach ($b in @($o.bekleyenGiden)) {
    if ([string]$b.kime -ne $e) { $kalan += $b; continue }
    $k = $o.konular | Where-Object { $_.id -eq $b.konuId } | Select-Object -First 1
    if (-not $k) { continue }
    [void](Ortak-Teslim $o @($e) (Ortak-Zarf $o $k @($e) ([string]$b.ic | ConvertFrom-Json)))
  }
  $o.bekleyenGiden = $kalan
}
# Paketi katilimcilara yollar. Doner: @{ ok = hepsi gitti mi; anahtar = anahtari beklenen kisi sayisi }
function Ortak-Gonder($o, [string[]]$kime, $paket) {
  $ben = Ben-Eposta
  $kime = @($kime | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() } | Where-Object { $_ -ne $ben } | Select-Object -Unique)
  if (-not $kime.Count) { return @{ ok = $true; anahtar = 0 } }
  $k = $o.konular | Where-Object { $_.id -eq $paket.konu.id } | Select-Object -First 1
  $ic = [ordered]@{ konu = $paket.konu; girisler = @($paket.girisler) }
  $bilinen = @($kime | Where-Object { $o.anahtarlar[$_] }); $bilinmeyen = @($kime | Where-Object { -not $o.anahtarlar[$_] })
  $ok = $true
  if ($bilinen.Count) { $ok = Ortak-Teslim $o $bilinen (Ortak-Zarf $o $k $bilinen $ic) }
  foreach ($e in $bilinmeyen) {
    $o.bekleyenGiden = @($o.bekleyenGiden) + @([pscustomobject]@{ kime = $e; konuId = $k.id; ic = ($ic | ConvertTo-Json -Depth 10 -Compress); t = (Get-Date).ToString("o") })
    Anahtar-Iste $o $e
  }
  return @{ ok = ($ok -and -not $bilinmeyen.Count); anahtar = $bilinmeyen.Count }
}function Ortak-Giris([string]$tur, [string]$metin) {
  $b = Ben-Kimlik
  return [pscustomobject]@{ id = [guid]::NewGuid().ToString("N"); t = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fff"); yazar = [pscustomobject]@{ ad = $b.ad; eposta = $b.eposta }; tur = $tur; metin = $metin }
}
# Gelen paketi birlestirir. $gonderen: iletinin gercek gondereni (Exchange'den). Doner: eklenen giris sayisi.
function Ortak-Al($o, $paket, [string]$gonderen) {
  $ben = Ben-Eposta; $gonderen = ([string]$gonderen).ToLowerInvariant()
  $h = $paket.konu
  if (-not $h -or -not $h.id) { return 0 }
  if (-not (@($h.katilimcilar) | Where-Object { ([string]$_.eposta).ToLowerInvariant() -eq $ben })) { return 0 }   # ben katilimci degilsem alinmaz
  $k = $o.konular | Where-Object { $_.id -eq $h.id } | Select-Object -First 1
  if (-not $k) {
    if (([string]$h.acan.eposta).ToLowerInvariant() -ne $gonderen) { return 0 }   # yeni konuyu yalniz acan getirebilir
    $k = [pscustomobject]@{ id = $h.id; baslik = $h.baslik; termin = $h.termin; acan = $h.acan; katilimcilar = @($h.katilimcilar); durum = $h.durum
      olusturma = $h.olusturma; katilma = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss"); okunan = ""; girisler = @() }
    $o.konular = @($o.konular) + @($k)
  } elseif (([string]$k.acan.eposta).ToLowerInvariant() -eq $gonderen) {
    $k.baslik = $h.baslik; $k.termin = $h.termin; $k.katilimcilar = @($h.katilimcilar); $k.durum = $h.durum
  }
  $acanGonderdi = (([string]$k.acan.eposta).ToLowerInvariant() -eq $gonderen)
  $var = @{}; foreach ($g in $k.girisler) { $var[[string]$g.id] = $true }
  $n = 0
  $gelecek = (Get-Date).AddDays(1)
  foreach ($g in @($paket.girisler | ForEach-Object { $_ })) {
    if (-not $g.id -or $var[[string]$g.id]) { continue }
    # G2 (16.10.2026): zaman GERCEK tarih olarak ayristirilir (metin karsilastirmasi bicim hilesiyle asilabiliyordu)
    $gt = [datetime]::MinValue
    if (-not $g.t -or -not [datetime]::TryParse([string]$g.t, [ref]$gt) -or $gt -gt $gelecek) { continue }
    if (([string]$g.yazar.eposta).ToLowerInvariant() -ne $gonderen) {
      # Bulgu 4 (15.10.2026): baskasinin girisi yalniz konuyu acan aktarabilir ve "aktaran" ile isaretlenir;
      # yazari dogrulanamadigi icin sayfada "aktarilan" diye gosterilir.
      if (-not $acanGonderdi) { continue }
      $g | Add-Member -NotePropertyName aktaran -NotePropertyValue $gonderen -Force
    }
    $k.girisler = @($k.girisler) + @($g); $var[[string]$g.id] = $true; $n++
  }
  $k.girisler = @($k.girisler | Sort-Object { [string]$_.t })
  return $n
}
# Outlook kurali (kullanici izni, 15.10.2026): konusu "[Assist] Ortak is:" olan ileti teslimde dogrudan
# Gelen Kutusu > Assist'e gider; Gelen Kutusu'nda gorunmez, yeni posta uyarisi cikmaz. Kural Exchange'de
# (sunucu tarafi) calisir: yalniz "klasore tasi" + "durdur" eylemleri. Kullanici Kurallar'dan silebilir;
# sildiyse her sunucu acilisinda yeniden kurulur (izin bir kez alindi). Sunucu basina bir kez denenir.
$script:ortakKuralBakildi = $false; $script:ortakKural = ""   # "" bakilmadi | var | kuruldu | hata
function Ortak-Kural {
  if ($script:ortakKuralBakildi -or $env:ISTAKIP_PORT -or -not (Outlook-Uygun)) { return }
  $script:ortakKuralBakildi = $true
  try {
    $ad = U "Assist ortak i{015F}ler"
    $sozcukler = [string[]]@(("[Assist] " + (U "Ortak i{015F}:")), "[Assist] Talep:")   # 15.10.2026: talep iletileri de
    $kurallar = (Ns).DefaultStore.GetRules()
    foreach ($r in $kurallar) {
      if ($r.Name -ne $ad) { continue }
      $var = @($r.Conditions.Subject.Text)
      if ($var -contains "[Assist] Talep:") { $script:ortakKural = "var"; return }
      try { $r.Conditions.Subject.Text = [string[]]@($var + "[Assist] Talep:"); $kurallar.Save($false); $script:ortakKural = "var"; Write-Host "  Outlook kuralina talep sozcugu eklendi" }
      catch { $script:ortakKural = "eksik"; Write-Host ("  Outlook kuralina talep sozcugu eklenemedi: " + $_.Exception.Message) -ForegroundColor Yellow }
      return
    }
    $r = $kurallar.Create($ad, 0)   # olRuleReceive
    $r.Conditions.Subject.Text = $sozcukler
    $r.Conditions.Subject.Enabled = $true
    $r.Actions.MoveToFolder.Folder = (Assist-Klasoru)
    $r.Actions.MoveToFolder.Enabled = $true
    $r.Actions.Stop.Enabled = $true
    $r.Enabled = $true
    $kurallar.Save($false)
    Write-Host "  Outlook kurali kuruldu: Assist ortak isler"; $script:ortakKural = "kuruldu"
  } catch { $script:ortakKural = "hata"; Write-Host ("  Outlook kurali kurulamadi: " + $_.Exception.Message) -ForegroundColor Yellow }
}
# Tek bir gelen paketi isler. Doner: eklenen giris sayisi.
function Ortak-Isle($o, $paket, [string]$gonderen) {
  $gonderen = ([string]$gonderen).ToLowerInvariant()
  if (-not $paket.v -or [int]$paket.v -lt 2) { return (Ortak-Al $o $paket $gonderen) }   # eski surum (sifresiz)
  if (([string]$paket.gonderen).ToLowerInvariant() -ne $gonderen) { Write-Host "  Ortak is: gonderen tutmadi, alinmadi" -ForegroundColor Yellow; return 0 }
  $durum = Anahtar-Ogren $o $gonderen ([string]$paket.acik)
  if ($durum -eq "gecersiz") { return 0 }
  if ($durum -eq "degisti") {   # onaylanana kadar bekler
    $o.bekleyenGelen = @($o.bekleyenGelen) + @([pscustomobject]@{ gonderen = $gonderen; paket = $paket; t = (Get-Date).ToString("o") })
    Write-Host ("  Ortak is: " + $gonderen + " anahtari degisti - onay bekliyor") -ForegroundColor Yellow
    return 0
  }
  if ($paket.tur -eq "anahtar-iste") { [void](Ortak-Teslim $o @($gonderen) ([ordered]@{ v = 2; tur = "anahtar"; gonderen = (Ben-Eposta); acik = (Benim-Acik) })) }
  if ($paket.tur -eq "anahtar-iste" -or $paket.tur -eq "anahtar") { Bekleyen-Gonder $o $gonderen; return 0 }
  if ($paket.tur -ne "veri") { return 0 }
  # dogrula: imza (gonderenin bilinen anahtari) -> konu anahtarini ac -> HMAC -> coz
  $acikR = Rsa-Acik ([string]$o.anahtarlar[$gonderen].acik)
  if (-not $acikR.VerifyData((Imza-Verisi $paket), [Convert]::FromBase64String([string]$paket.imza), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)) {
    Write-Host ("  Ortak is: imza gecersiz, alinmadi (" + $gonderen + ")") -ForegroundColor Yellow; return 0 }
  $kilit = $null; foreach ($p in $paket.anahtarlar.PSObject.Properties) { if ($p.Name.ToLowerInvariant() -eq (Ben-Eposta)) { $kilit = [string]$p.Value } }
  if (-not $kilit) { return 0 }   # bana kilitlenmemis
  $ak = (Benim-Rsa).Decrypt([Convert]::FromBase64String($kilit), $true)
  $ic = $UTF8.GetString((Aes-Coz $ak ([string]$paket.iv) ([string]$paket.sif) ([string]$paket.mac))) | ConvertFrom-Json
  if ([string]$ic.konu.id -ne [string]$paket.konuId) { return 0 }
  $n = Ortak-Al $o $ic $gonderen
  $k = $o.konular | Where-Object { $_.id -eq $paket.konuId } | Select-Object -First 1
  if ($k -and -not $k.anahtar) { $k | Add-Member -NotePropertyName anahtar -NotePropertyValue (B64 $ak) -Force }   # sonraki gonderimlerim ayni anahtarla
  return $n
}
# Ortak is iletilerini alir: kural kurulmadan Gelen Kutusu'na dusenler Assist'e tasinir; kuralla Assist
# klasorune gelenler (okunmamis) islenir ve okundu yapilir. Gonderim kuyrugu da bosaltilir.
# Test sunucusu (ISTAKIP_TEST_POSTA): Outlook yerine klasordeki bana gelen dosyalar.
function Ortak-Kontrol {
  $test = [bool]$env:ISTAKIP_TEST_POSTA
  if (-not $test -and -not (Outlook-Uygun)) { return }
  $o = Ortak-Oku; $degisti = $false
  if (@($o.giden).Count) {
    $kalan = @()
    foreach ($p in $o.giden) { try { Ortak-Ileti @($p.kime) ([string]$p.json) ([string]$p.baslik) } catch { $kalan += $p } }
    $o.giden = $kalan; $degisti = $true
  }
  # Anahtari beklenen kisilerden (Assist'i kapali / eski surum) 10 dk'da bir yeniden iste
  $bek = @($o.bekleyenGiden | ForEach-Object { [string]$_.kime } | Select-Object -Unique)
  foreach ($e in $bek) { $once = [string]$o.istenen[$e]; Anahtar-Iste $o $e; if ([string]$o.istenen[$e] -ne $once) { $degisti = $true } }
  $gelen = @()
  if ($test) {
    foreach ($d in @(Get-ChildItem -LiteralPath $env:ISTAKIP_TEST_POSTA -Filter "*.json" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime)) {
      try { $x = [System.IO.File]::ReadAllText($d.FullName, $UTF8) | ConvertFrom-Json } catch { continue }
      if (([string]$x.kime).ToLowerInvariant() -ne (Ben-Eposta)) { continue }
      $gelen += @{ govde = [string]$x.govde; gonderen = [string]$x.gonderen; dosya = $d.FullName }
    }
  } else {
    $gk = (Ns).GetDefaultFolder(6); $hedef = Assist-Klasoru
    foreach ($it in $gk.Items.Restrict("[MessageClass] = '$ORTAK_SINIF'")) { $gelen += @{ it = $it } }
    foreach ($it in $hedef.Items.Restrict("[MessageClass] = '$ORTAK_SINIF' AND [UnRead] = True")) { $gelen += @{ it = $it } }
  }
  $yeni = 0
  foreach ($g in $gelen) {
    try {
      if ($g.it) { $g.govde = [string]$g.it.Body; $g.gonderen = Gonderen-Eposta $g.it }
      $govde = $g.govde
      $i = $govde.IndexOf($ORTAK_ISARET); $j = if ($i -ge 0) { $govde.IndexOf(":SON", $i) } else { -1 }
      if ($i -ge 0 -and $j -gt $i) {
        $b64 = ($govde.Substring($i + $ORTAK_ISARET.Length, $j - $i - $ORTAK_ISARET.Length)) -replace '\s', ''
        $paket = $UTF8.GetString([Convert]::FromBase64String($b64)) | ConvertFrom-Json
        $yeni += Ortak-Isle $o $paket $g.gonderen
        $degisti = $true
      }
    } catch { Write-Host ("  Ortak is iletisi islenemedi: " + $_.Exception.Message) -ForegroundColor Yellow }
    try {
      if ($g.dosya) { [System.IO.File]::Delete($g.dosya) }
      elseif ($g.it) { $g.it.UnRead = $false; $g.it.Save(); if ($g.it.Parent.EntryID -ne $hedef.EntryID) { [void]$g.it.Move($hedef) } }   # once kaydet, sonra tasi
    } catch {}
  }
  # Kuralla Assist klasorune gelen teslim/okundu bildirimleri ("Delivered: [Assist] ...") okundu yapilir; birikmesin
  if (-not $test) { try { foreach ($it in @((Assist-Klasoru).Items.Restrict("[UnRead] = True"))) { if ([string]$it.MessageClass -like "REPORT.*") { $it.UnRead = $false; $it.Save() } } } catch {} }
  if ($degisti) { Ortak-Yaz $o }
  if ($yeni) { try { Bildirim-Goster $BLD_AD ($(if ($yeni -eq 1) { U "Ortak i{015F}lerde 1 yeni giri{015F}" } else { (U "Ortak i{015F}lerde ") + $yeni + (U " yeni giri{015F}") })) "ortak" } catch {} }
}
# Bu Windows oturumunun kurumsal e-postasi (kucuk harf, onbellekli): once etki alani, olmazsa Outlook.
$script:benEposta = $null
function Ben-Eposta {
  if ($null -ne $script:benEposta) { return $script:benEposta }
  if ($env:ISTAKIP_BEN_EPOSTA) { $script:benEposta = $env:ISTAKIP_BEN_EPOSTA.ToLowerInvariant(); return $script:benEposta }   # test sunucusu
  $e = ""
  try { Add-Type -AssemblyName System.DirectoryServices.AccountManagement; $e = [string][System.DirectoryServices.AccountManagement.UserPrincipal]::Current.EmailAddress } catch {}
  if (-not $e) { try { $e = [string](Ns).CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {} }
  $script:benEposta = $e.Trim().ToLowerInvariant()
  return $script:benEposta
}


# --- TALEP UCTAN UCA SIFRELEME (G6/G7, 16.10.2026) -----------------------------------
# Talep iletileri de ortak istekiyle AYNI altyapiyla sifreli + imzali gider: icerik (baslik,
# aciklama, cevap, kapatma/silme) Exchange'de OKUNAMAZ ve imzasiz ileti kabul edilmez; boylece
# "Send As / Send on Behalf" ile baskasi adina cevap/kapat/sil URETILEMEZ. Alicinin acik anahtari
# bilinmiyorsa istenir, talep kuyrukta (bekleyenTalep) bekler, anahtar gelince yollanir.
# Tek fark: talep tek aliciya gider ve her iletinin kendi tek kullanimlik anahtari vardir (saklanmaz).
function Talep-Zarf($o, [string]$alici, $ic) {
  $ak = New-Object byte[] 64; [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($ak)
  $s = Aes-Sifrele $ak ($UTF8.GetBytes(($ic | ConvertTo-Json -Depth 8 -Compress)))
  $kilit = @{}; $kilit[$alici] = B64 ((Rsa-Acik ([string]$o.anahtarlar[$alici].acik)).Encrypt($ak, $true))
  $z = [ordered]@{ v = 2; tur = "talep"; konuId = [guid]::NewGuid().ToString("N"); gonderen = (Ben-Eposta); acik = (Benim-Acik); anahtarlar = $kilit; iv = $s.iv; sif = $s.sif; mac = $s.mac }
  $z.imza = B64 ((Benim-Rsa).SignData((Imza-Verisi $z), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1))
  return $z
}
# Gelen sifreli talep zarfini dogrular ve acar; basarisizsa HATA firlatir. Doner: talep nesnesi.
function Talep-Ic-Coz($paket, [string]$gonderen) {
  $gonderen = ([string]$gonderen).ToLowerInvariant()
  if (([string]$paket.gonderen).ToLowerInvariant() -ne $gonderen) { throw "gonderen tutmadi" }
  if ([string]$paket.tur -ne "talep") { throw "tur beklenmeyen" }
  $o = Ortak-Oku
  $durum = Anahtar-Ogren $o $gonderen ([string]$paket.acik)
  if ($durum -eq "gecersiz") { throw "anahtar gecersiz" }
  if ($durum -eq "degisti") { Ortak-Yaz $o; throw ("anahtar degisti - onay gerekli: " + $gonderen) }
  if ($durum -eq "yeni") { Ortak-Yaz $o }
  $acikR = Rsa-Acik ([string]$o.anahtarlar[$gonderen].acik)
  if (-not $acikR.VerifyData((Imza-Verisi $paket), [Convert]::FromBase64String([string]$paket.imza), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)) { throw "imza gecersiz" }
  $kilit = $null; foreach ($p in $paket.anahtarlar.PSObject.Properties) { if ($p.Name.ToLowerInvariant() -eq (Ben-Eposta)) { $kilit = [string]$p.Value } }
  if (-not $kilit) { throw "bana kilitlenmemis" }
  $ak = (Benim-Rsa).Decrypt([Convert]::FromBase64String($kilit), $true)
  $ic = $UTF8.GetString((Aes-Coz $ak ([string]$paket.iv) ([string]$paket.sif) ([string]$paket.mac))) | ConvertFrom-Json
  return $ic.talep
}
# Tek talebi aliciya yollar (test modunda ISTAKIP_TEST_POSTA klasorune).
function Talep-Yolla($o, [string]$alici, $talep, [string]$konu, $posta) {
  $zarf = Talep-Zarf $o $alici ([ordered]@{ talep = $talep; konu = $konu })
  $govde = (U "Bu talep Assist ile g{00F6}nderildi (sifreli; yaln{0131}z Assist kullan{0131}c{0131}lar{0131} a{00E7}ar).") + "`r`n" + $VERI_ISARET + ([Convert]::ToBase64String($UTF8.GetBytes(($zarf | ConvertTo-Json -Depth 8 -Compress)))) + ":SON"
  if ($env:ISTAKIP_TEST_POSTA) {
    [System.IO.File]::WriteAllText((Join-Path $env:ISTAKIP_TEST_POSTA ([guid]::NewGuid().ToString("N") + ".json")), (@{ kime = $alici; gonderen = (Ben-Eposta); govde = $govde; sinif = "talep" } | ConvertTo-Json -Compress), $UTF8)
    return
  }
  $m = (Outlook).CreateItem(0); $m.MessageClass = $GELEN_SINIF; $m.To = $alici
  $m.OriginatorDeliveryReportRequested = $false; $m.ReadReceiptRequested = $false
  $m.Subject = "[Assist] " + (U "Talep: g{00FC}venli ileti"); $m.Body = $govde
  if ($posta -and $posta.id) { try { $k = Posta-Bul ([string]$posta.id) ([string]$posta.store) ([string]$posta.msgId); if ($k) { [void]$m.Attachments.Add($k, 5) } } catch {} }
  [void]$m.Recipients.ResolveAll(); $m.Send()
}
# Anahtari beklenen taleplerin kuyrugu: anahtar gelmisse yollar, yoksa yeniden ister.
function Talep-Giden-Kontrol {
  if (-not $env:ISTAKIP_TEST_POSTA -and -not (Outlook-Uygun)) { return }
  $o = Ortak-Oku; if (-not @($o.bekleyenTalep).Count) { return }
  $kalan = @(); $degisti = $false
  foreach ($b in @($o.bekleyenTalep)) {
    $alici = ([string]$b.kime).ToLowerInvariant()
    if (-not $o.anahtarlar[$alici]) { $once = [string]$o.istenen[$alici]; Anahtar-Iste $o $alici; if ([string]$o.istenen[$alici] -ne $once) { $degisti = $true }; $kalan += $b; continue }
    try { Talep-Yolla $o $alici $b.talep ([string]$b.konu) $b.posta; $degisti = $true; Write-Host ("  Kuyruktaki talep yollandi: " + $alici) }
    catch { $kalan += $b }
  }
  $o.bekleyenTalep = $kalan; if ($degisti) { Ortak-Yaz $o }
}
function Gelen-Talep-Kontrol {
  if (-not (Outlook-Uygun)) { return }
  $gk = (Ns).GetDefaultFolder(6)
  $bulunan = @(); foreach ($it in $gk.Items.Restrict("[MessageClass] = '$GELEN_SINIF'")) { $bulunan += $it }   # tasirken koleksiyon degismesin
  # Outlook kurali (konu "[Assist] Talep:") talebi dogrudan Assist klasorune koyar: orada okunmamis olanlar da alinir
  try { foreach ($it in (Assist-Klasoru).Items.Restrict("[MessageClass] = '$GELEN_SINIF' AND [UnRead] = True")) { $bulunan += $it } } catch {}
  if (-not $bulunan.Count) { return }
  $L = New-Object System.Collections.ArrayList
  foreach ($x in @(Gelen-Oku)) { [void]$L.Add($x) }
  $hedef = Assist-Klasoru; $yeni = 0
  # Bulgu 5 (15.10.2026): yalniz kendi kurum alan adindan, makul boyutta ve dakikada en cok 20 talep alinir.
  # Kurala uymayan ileti Gelen Kutusu'nda normal posta olarak KALIR (tasinmaz, kayit olusmaz).
  $benAlan = ((Ben-Eposta) -split '@')[-1]
  if (-not $script:gelenRed) { $script:gelenRed = @{} }   # reddedilen iletiler her dakika yeniden yazilmasin
  foreach ($it in $bulunan) {
    if ($yeni -ge 20) { break }
    try {
      if ($script:gelenRed.ContainsKey([string]$it.EntryID)) { continue }
      $gonderen = [string]$it.SenderName; $ep = Gonderen-Eposta $it; $alinma = Tarih-Iso $it.ReceivedTime
      if (-not $benAlan -or ([string]$ep).ToLowerInvariant() -notlike ("*@" + $benAlan)) { $script:gelenRed[[string]$it.EntryID] = $true; Write-Host "  Gelen talep reddedildi: kurum disi gonderen" -ForegroundColor Yellow; continue }
      $govde = [string]$it.Body
      $i = $govde.IndexOf($VERI_ISARET); if ($i -lt 0) { continue }
      $j = $govde.IndexOf(":SON", $i); if ($j -lt 0) { continue }
      if ($j - $i -gt 90000) { $script:gelenRed[[string]$it.EntryID] = $true; Write-Host "  Gelen talep reddedildi: veri cok buyuk" -ForegroundColor Yellow; continue }
      $b64 = ($govde.Substring($i + $VERI_ISARET.Length, $j - $i - $VERI_ISARET.Length)) -replace '\s', ''
      $paket = $UTF8.GetString([Convert]::FromBase64String($b64)) | ConvertFrom-Json
      # G6/G7 (16.10.2026): v2 ise sifreli+imzali; dogrula ve coz. v1 (eski, sifresiz) gecis icin alinir ama dogrulanmamis isaretlenir.
      $dogrulandi = $false
      if ($paket -and $paket.v -and [int]$paket.v -ge 2) {
        $t = $null; $hata = ""
        try { $t = Talep-Ic-Coz $paket $ep } catch { $hata = $_.Exception.Message }
        if (-not $t) { $script:gelenRed[[string]$it.EntryID] = $true; Write-Host ("  Gelen talep reddedildi: " + $hata) -ForegroundColor Yellow; continue }
        $dogrulandi = $true
      } else { $t = $paket }
      if (-not $t -or $t -isnot [System.Management.Automation.PSCustomObject] -or ([string]$t.baslik).Length -gt 500 -or @($t.aksiyonlar).Count -gt 50) { $script:gelenRed[[string]$it.EntryID] = $true; Write-Host "  Gelen talep reddedildi: beklenmeyen bicim" -ForegroundColor Yellow; continue }
      $msgId = ""; try { $msgId = [string]$it.PropertyAccessor.GetProperty($PR_MSGID) } catch {}
      # Okundu isareti TASIMADAN ONCE kaydedilir: kaydedilmeden tasinan ileti Assist klasorunde okunmamis kaliyor,
      # bir sonraki taramada yeniden alinip talep IKI KEZ dusuyordu (15.10.2026)
      $it.UnRead = $false; $it.Save()
      if ($it.Parent.EntryID -eq $hedef.EntryID) { $tasinan = $it } else { $tasinan = $it.Move($hedef) }
      if ($msgId -and @($L | Where-Object { $_.posta -and [string]$_.posta.msgId -eq $msgId }).Count) { continue }   # ayni ileti zaten kuyrukta
      # Gonderen talebe bilerek posta ekladiyse (posta penceresi > + > Talep, 16.10.2026) tasiyici iletide ekli
      # posta (olEmbeddeditem) olur: adi "ek" alanina yazilir, sayfa yalniz o zaman talebe posta olarak baglar.
      $ekAd = ""; try { foreach ($e in $tasinan.Attachments) { if ([int]$e.Type -eq 5) { $ekAd = [string]$e.DisplayName; break } } } catch {}
      [void]$L.Add([ordered]@{ id = [guid]::NewGuid().ToString("N"); gonderen = $gonderen; gonderenEposta = $ep; alinma = $alinma; talep = $t; dogrulanmis = $dogrulandi;
        posta = @{ id = [string]$tasinan.EntryID; store = [string]$hedef.StoreID; msgId = $msgId; konu = [string]$tasinan.Subject; kimden = $gonderen; tarih = $alinma; ek = $ekAd } })
      $yeni++
    } catch { Write-Host ("  Gelen talep okunamadi: " + $_.Exception.Message) -ForegroundColor Yellow }
  }
  if ($yeni) {
    Gelen-Yaz $L
    $script:gelenRev++   # sayfa /api/posta/rev'den hemen gorsun (talep kuralla Assist klasorune gidince Gelen Kutusu degismez)
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
  # HTML govde (16.10.2026): Outlook'taki gibi baglantilar, tablolar ve resimler gorunsun.
  # Govdedeki "cid:" resimler (ekli satir ici resimler) data: adresine cevrilir; satir ici
  # ekler ek listesinde gosterilmez (Outlook da gostermez). Sayfa bu HTML'i betiksiz, ayri
  # kokenli bir iframe'de (sandbox) gosterir; uzak resimler varsayilan olarak engelli.
  $html = Posta-Html $it
  $ekler = @(); $ekListe = @(); try { foreach ($e in $it.Attachments) { if ($html.satirIci -notcontains [string]$e.Index) { $ekler += [string]$e.FileName; $ekListe += @{ ad = [string]$e.FileName; i = [int]$e.Index } } } } catch {}
  return @{ ok = $true; id = [string]$it.EntryID; store = [string]$it.Parent.StoreID; msgId = $msgId
    konu = [string]$it.Subject; kimden = [string]$it.SenderName; eposta = (Gonderen-Eposta $it)
    kime = [string]$it.To; bilgi = [string]$it.CC; tarih = (Tarih-Iso $it.ReceivedTime)
    bayrakTermin = (Bayrak-Termin $it); govde = $govde; html = $html.metin; ekler = @($ekler); ekListe = @($ekListe) }
}
$PR_CID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
$RESIM_TUR = @{ png = "image/png"; jpg = "image/jpeg"; jpeg = "image/jpeg"; gif = "image/gif"; bmp = "image/bmp"; webp = "image/webp" }
function Posta-Html($it) {
  $sonuc = @{ metin = ""; satirIci = @() }
  try { if ([int]$it.BodyFormat -eq 1) { return $sonuc } } catch {}   # 1 = duz metin
  $h = ""; try { $h = [string]$it.HTMLBody } catch { return $sonuc }
  if (-not $h -or $h.Length -gt 2000000) { return $sonuc }
  $toplam = 0
  try {
    foreach ($e in $it.Attachments) {
      $cid = ""; try { $cid = [string]$e.PropertyAccessor.GetProperty($PR_CID) } catch {}
      if (-not $cid) { continue }
      $cid = $cid.Trim('<', '>')
      if ($h.IndexOf("cid:" + $cid, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
      $uz = ([IO.Path]::GetExtension([string]$e.FileName)).TrimStart('.').ToLowerInvariant()
      $mime = $RESIM_TUR[$uz]; if (-not $mime) { continue }
      if ([int]$e.Size -gt 1500000 -or ($toplam + [int]$e.Size) -gt 6000000) { continue }
      $gecici = Join-Path ([IO.Path]::GetTempPath()) ("assist_" + [guid]::NewGuid().ToString("N") + "." + $uz)
      try {
        $e.SaveAsFile($gecici)
        $b = [IO.File]::ReadAllBytes($gecici); $toplam += $b.Length
        $h = [regex]::Replace($h, ("cid:" + [regex]::Escape($cid)), ("data:" + $mime + ";base64," + [Convert]::ToBase64String($b)), "IgnoreCase")
        $sonuc.satirIci += [string]$e.Index
      } catch {} finally { Remove-Item -LiteralPath $gecici -Force -ErrorAction SilentlyContinue }
    }
  } catch {}
  $sonuc.metin = $h
  return $sonuc
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
# Posta listesi isaretleri (16.10.2026). Ek-Say: gercek ek sayisi; gizli ekler ve govdedeki satir ici
# resimler (imza logosu gibi, content-id'li resim) sayilmaz. Yazisma-Sira: ayni yazismada bu iletiden
# onceki ileti sayisi; Outlook ConversationIndex'i 22 bayt baslik + her yanit icin 5 bayt (hex: 44 + 10n).
$PR_EK_GIZLI = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
function Ek-Adlari($it) {
  $L = @()
  try {
    if ([int]$it.Attachments.Count -eq 0) { return @() }
    # Hiz: ozellik yalniz resim eklerinde okunur (satir ici / gizli ekler pratikte hep resim); digerleri dogrudan sayilir
    foreach ($e in $it.Attachments) {
      $ad = [string]$e.FileName
      if ($ad -notmatch '\.(png|jpe?g|gif|bmp|webp)$') { $L += $ad; continue }
      $pa = $e.PropertyAccessor; $gizli = $false; $cid = ""
      try { $gizli = [bool]$pa.GetProperty($PR_EK_GIZLI) } catch {}
      if (-not $gizli) { try { $cid = [string]$pa.GetProperty($PR_CID) } catch {} }
      if (-not $gizli -and -not $cid) { $L += $ad }
    }
  } catch {}
  return ,$L
}
function Ek-Say($it) { return @(Ek-Adlari $it).Count }
# Hiz icin: klasor tablosundan (tek toplu okuma) EntryID -> "gercek eki var" (PR_HASATTACH). Tablodaki deger
# Exchange'in hesapladigi degerdir ve satir ici resimleri saymaz (16.10.2026: 134 postada Ek-Say ile birebir);
# iletinin kendisinden okunan PR_HASATTACH ise satir ici resim varken de true doner, o yuzden kullanilmaz.
$PR_HASATTACH = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"
function Ek-Haritasi($f, [int]$enCok) {
  $h = @{}
  try {
    $tb = $f.GetTable(); $tb.Sort("[ReceivedTime]", $true); [void]$tb.Columns.Add($PR_HASATTACH)
    $n = 0
    while (-not $tb.EndOfTable -and $n -lt $enCok) { $r = $tb.GetNextRow(); $n++; try { $h[[string]$r.Item("EntryID")] = [bool]$r.Item($PR_HASATTACH) } catch {} }
  } catch {}
  return $h
}
function Yazisma-Sira($it) {
  try { $ci = [string]$it.ConversationIndex; if ($ci.Length -gt 44) { return [int][math]::Floor(($ci.Length - 44) / 10) } } catch {}
  return 0
}
# Dosya turu simgesi: SHGetFileInfo (USEFILEATTRIBUTES: dosya gerekmez, yalniz "x.<uzanti>") -> 32 px PNG.
# Uzanti basina bellekte saklanir. Tanimsiz turde Windows'un genel dosya simgesi doner.
$script:simgeOnbellek = @{}
function Dosya-Simge([string]$uz) {
  if ($uz -ne "__genel" -and $script:simgeOnbellek.ContainsKey($uz)) { return ,$script:simgeOnbellek[$uz] }
  if (-not ('AssistDosyaSimge' -as [type])) {
    Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System; using System.Runtime.InteropServices; using System.Drawing;
public static class AssistDosyaSimge {
  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
  struct SHFILEINFO { public IntPtr hIcon; public int iIcon; public uint dwAttributes;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName; }
  [DllImport("shell32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SHGetFileInfo(string p, uint a, ref SHFILEINFO s, uint cb, uint f);
  [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr h);
  public static byte[] Png(string uzanti) {
    SHFILEINFO s = new SHFILEINFO();
    SHGetFileInfo("x." + uzanti, 0x80, ref s, (uint)Marshal.SizeOf(s), 0x100 | 0x10);   // SHGFI_ICON | SHGFI_USEFILEATTRIBUTES (buyuk simge)
    if (s.hIcon == IntPtr.Zero) return null;
    try { using (Icon ic = Icon.FromHandle(s.hIcon)) using (Bitmap bm = ic.ToBitmap()) using (System.IO.MemoryStream ms = new System.IO.MemoryStream()) {
      bm.Save(ms, System.Drawing.Imaging.ImageFormat.Png); return ms.ToArray(); } }
    finally { DestroyIcon(s.hIcon); }
  }
}
'@
  }
  # Windows'ta tanimsiz turler genel bos sayfa simgesi alir: o zaman simge verilmez (sayfa kendi isaretini koyar)
  if (-not $script:simgeOnbellek.ContainsKey("__genel")) { $g = $null; try { $g = [AssistDosyaSimge]::Png("assistbilinmeyentur") } catch {}; $script:simgeOnbellek["__genel"] = $g }
  $b = $null; try { $b = [AssistDosyaSimge]::Png($uz) } catch {}
  $g = $script:simgeOnbellek["__genel"]
  if ($b -and $g -and [Convert]::ToBase64String($b) -eq [Convert]::ToBase64String($g)) { $b = $null }
  $script:simgeOnbellek[$uz] = $b
  return ,$b
}
# Ayni yazismadaki (Outlook "konusma") iletiler, eskiden yeniye; en cok 30. Konusma kapaliysa
# Gelen + Gonderilmis'te ayni konu basligi (ConversationTopic). Govdeden yalniz ilk 1500 karakter
# (sayfa alintilanan eski metni keser). Assist'in tasiyici iletileri alinmaz. (16.10.2026)
function Posta-Yazisma($it) {
  $L = New-Object System.Collections.ArrayList; $gor = @{}
  $ekle = { param($x)
    try {
      if ([int]$x.Class -ne 43) { return }
      if (([string]$x.MessageClass) -like "IPM.Note.Assist*" -or ([string]$x.Subject).StartsWith("[Assist] ")) { return }
      $eid = [string]$x.EntryID; if ($gor.ContainsKey($eid)) { return }; $gor[$eid] = 1
      $t = [datetime]$x.ReceivedTime; if ($t.Year -gt 4000) { $t = [datetime]$x.SentOn }
      $g = [string]$x.Body; if ($g.Length -gt 1500) { $g = $g.Substring(0, 1500) }
      [void]$L.Add([ordered]@{ id = $eid; store = [string]$x.Parent.StoreID; kimden = ([string]$x.SenderName).Trim().Trim("'", '"').Trim()
        tarih = (Tarih-Iso $t); govde = $g; ekler = @(Ek-Adlari $x); t = $t })   # gercek ekler (imza resimleri haric)
    } catch {}
  }
  $ns = Ns; $konusma = $null
  try { $konusma = $it.GetConversation() } catch {}
  if ($konusma) {
    try {
      $tb = $konusma.GetTable(); $say = 0
      while (-not $tb.EndOfTable -and $say -lt 60) {
        $r = $tb.GetNextRow(); $say++
        try { & $ekle ($ns.GetItemFromID([string]$r.Item("EntryID"))) } catch {}
      }
    } catch {}
  }
  if ($L.Count -le 1) {
    $konu = ""; try { $konu = [string]$it.ConversationTopic } catch {}
    if ($konu) {
      $filtre = "[ConversationTopic] = '" + $konu.Replace("'", "''") + "'"
      foreach ($no in 6, 5) { try { foreach ($x in $ns.GetDefaultFolder($no).Items.Restrict($filtre)) { & $ekle $x; if ($L.Count -ge 60) { break } } } catch {} }
    }
  }
  $S = @($L | Sort-Object { $_.t } | Select-Object -Last 30)
  foreach ($o in $S) { $o.Remove("t") }
  return $S
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
$script:testPostaSon = [datetime]::MinValue; $script:gelenSon = [datetime]::MinValue; $script:gelenRev = 0; $script:gelenOkunmamis = -1; $script:gelenToplam = -1   # Gelen Kutusu izleme
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
  $komut = 'powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy ' + (Assist-Politika) + ' -File "' + (Join-Path $betikKok "pencere.ps1") + '" -Port ' + $port + ' -Url "%1"'
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
    if (((Get-Date) - $script:postaSon).TotalMinutes -ge $dk -and (Outlook-Uygun)) {
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
  $k.TargetPath = Assist-Baslatici $portalKok; $k.Arguments = ""   # imza zorunluysa yerel baslatici
  $k.WorkingDirectory = $portalKok; $k.WindowStyle = 7; $k.Description = "STFA Assist"
  $k.IconLocation = (Join-Path $portalKok "AssistA.ico") + ",0"
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
# Logo degisti (15.10.2026, AssistLogo.ico / Assist.ico -> AssistA.ico; yeni ad = Windows simge onbellegi yenilenir):
# masaustu ve Baslangic klasorundeki Assist kisayollari yeni simgeye gecirilir.
try {
  $yeniSimge = Join-Path $portalKok "AssistA.ico"
  if (Test-Path -LiteralPath $yeniSimge) {
    $ws = New-Object -ComObject WScript.Shell
    foreach ($kl in @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Startup"))) {
      foreach ($l in @(Get-ChildItem -LiteralPath $kl -Filter "*.lnk" -ErrorAction SilentlyContinue)) {
        $k = $ws.CreateShortcut($l.FullName)
        if ($k.IconLocation -match 'AssistLogo\.ico|AssistSimge\.ico|\\Assist\.ico') { $k.IconLocation = $yeniSimge + ",0"; $k.Save() }
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
      if (-not $env:ISTAKIP_PORT) { try { Talep-Giden-Kontrol } catch { Write-Host ("  Talep giden hatasi: " + $_.Exception.Message) -ForegroundColor Yellow } }
      if (-not $env:ISTAKIP_PORT) { try { Ortak-Kural } catch {} }
      if (-not $env:ISTAKIP_PORT) { try { Ortak-Kontrol } catch { Write-Host ("  Ortak is hatasi: " + $_.Exception.Message) -ForegroundColor Yellow } }
      try { Gecici-Temizle } catch {}
      try { Mobil-Disa-Aktar } catch { Write-Host ("  Mobil disa aktarma hatasi: " + $_.Exception.Message) -ForegroundColor Yellow }
    }
    if ($env:ISTAKIP_TEST_POSTA -and ((Get-Date) - $script:testPostaSon).TotalSeconds -ge 2) { $script:testPostaSon = Get-Date; try { Ortak-Kontrol } catch { Write-Host ("  [test] ortak hata: " + $_.Exception.Message) } }
    # Gelen Kutusu izleme (5 sn): yalniz okunmamis + toplam sayi okunur (hafif, korumali alan yok). Degisince
    # gelenRev artar; sayfa ve tepsi /api/posta/rev'i yoklayip hemen tazelenir (yeni posta aninda dusar).
    if (((Get-Date) - $script:gelenSon).TotalSeconds -ge 5) {
      $script:gelenSon = Get-Date
      if (Outlook-Uygun) {
        try {
          $gk = (Ns).GetDefaultFolder(6); $u = [int]$gk.UnReadItemCount; $c = [int]$gk.Items.Count
          $ak = 0; try { $ak = [int]$gk.Folders.Item("Assist").UnReadItemCount } catch {}
          if ($ak -gt 0 -and -not $env:ISTAKIP_PORT) { try { Ortak-Kontrol } catch {}; try { Gelen-Talep-Kontrol } catch {} }   # kuralla Assist klasorune ileti geldi
          if ($u -ne $script:gelenOkunmamis -or $c -ne $script:gelenToplam) {
            # Ilk okumada da rev artar (16.10.2026): acilista sayfa ve tepsi okunmamis sayiyi hemen alsin
            $script:gelenRev++
            if ($script:gelenToplam -ge 0) {
              if (-not $env:ISTAKIP_PORT) { try { Ortak-Kontrol } catch {}; try { Gelen-Talep-Kontrol } catch {} }   # ortak is / talep iletisi geldiyse hemen islensin
            }
            $script:gelenOkunmamis = $u; $script:gelenToplam = $c
          }
        } catch {}
      }
    }
    try { $z = (Get-Item -LiteralPath $PSCommandPath).LastWriteTimeUtc.Ticks } catch { $z = $BETIK_ZAMAN }
    if ($z -ne $BETIK_ZAMAN) {
      # Imza zorunlu kipte yeni kod, sabitlenen sertifikayla gecerli imzali degilse YUKLENMEZ (guvenlik bulgu 1).
      if ($GUV.imzaZorunlu) {
        Start-Sleep -Milliseconds 500   # dosya yazimi bitsin
        $yeniOk = (Assist-Imza-Gecerli $PSCommandPath $GUV.parmakIzi) -and (Assist-Imza-Gecerli (Join-Path $betikKok "guvenlik.ps1") $GUV.parmakIzi)
        if (-not $yeniOk) {
          Write-Host "  GUVENLIK: sunucu dosyasi degisti ama imzasi gecersiz - yeniden YUKLENMEDI, mevcut surum calisiyor." -ForegroundColor Red
          $BETIK_ZAMAN = $z; continue
        }
      }
      $yenidenYukle = $true; break
    }
  }
  if ($yenidenYukle) {
    Write-Host "  Sunucu dosyasi guncellendi - yeniden yukleniyor..." -ForegroundColor Cyan
    Start-Sleep -Milliseconds 500   # dosya yazimi bitsin
    try { $listener.Stop(); $listener.Close() } catch {}
    # Bulgu 10 (15.10.2026): ic ice "& $PSCommandPath" her guncellemede cagri yiginini buyutuyordu.
    # Yeni sunucu_calistir.ps1 dongusu varsa ona birakilir (yigin buyumez); eski baslatici icin eski yol.
    if ($global:AssistSunucuDongusu) { $global:AssistYenidenYukle = $true; return }
    & $PSCommandPath
    return
  }
  try { $ctx = $listener.EndGetContext($bekle) } catch { break }
  try {
    $yol   = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath)
    $metod = $ctx.Request.HttpMethod

    if ($yol -eq "/__saglik") {
      # Kanit: istemci (baslat/pencere/tepsi) rastgele nonce yollar; sunucu oturum anahtariyla HMAC uretir.
      # Anahtari bilmeyen bir program portu kapsa bile Assist sunucusu taklidi yapamaz (guvenlik bulgu 2).
      $sg = [ordered]@{ ok = $true; urun = "istakip"; surum = $SURUM; portalKok = $portalKok; pid = $PID; betikZaman = [string]$BETIK_ZAMAN }
      $nonce = Sorgu $ctx "nonce"
      if ($nonce -match '^[0-9a-f]{32}$') { $sg.kanit = Assist-Hmac $ANAHTAR (Assist-Kanit-Metni ([pscustomobject]$sg) $nonce) }
      Yaz-Json $ctx $sg
      $ctx.Response.OutputStream.Close(); continue
    }

    # Dosya turu simgesi (16.10.2026): posta ek listesinde Windows'un o tur icin gosterdigi ozgun simge
    # (Excel, Word, PDF... kurulu programin simgesi; Gezgin'deki gibi). Yalniz uzanti alir, dosya okumaz,
    # disari istek yapmaz; <img> baslik gonderemedigi icin anahtarsiz, ama yalniz yerel istege cevap verir.
    if ($yol -match '^/simge/([a-z0-9]{1,8})\.png$') {
      $b = $null; if ($ctx.Request.IsLocal) { $b = Dosya-Simge $Matches[1] }
      if ($b) { $ctx.Response.StatusCode = 200; $ctx.Response.ContentType = "image/png"; $ctx.Response.Headers.Add("Cache-Control", "max-age=86400"); $ctx.Response.ContentLength64 = $b.Length; $ctx.Response.OutputStream.Write($b, 0, $b.Length) }
      else { $ctx.Response.StatusCode = 404 }
      $ctx.Response.OutputStream.Close(); continue
    }

    if ($yol.StartsWith("/api/")) {
      # Istek kaydi: sunucu bir istekte takilirsa kayittaki SON satir hangisi oldugunu gosterir.
      if (($yol -ne "/api/pencere/hedef" -or $metod -ne "GET") -and $yol -ne "/api/acilis/durum" -and $yol -ne "/api/posta/rev") { Write-Host ("  > " + (Get-Date).ToString("HH:mm:ss") + " " + $metod + " " + $yol) }   # Bulgu 7 (15.10.2026): sorgu dizisi (posta kimligi, kategori, kisi adi) kayda YAZILMAZ; pencerenin 2 sn yoklamasi yazilmaz
      # Baska bir sayfadan gelen istek ozel basligi gonderemez (CORS izni yok).
      $koken = [string]$ctx.Request.Headers["Origin"]
      $kokenOk = (-not $koken) -or ($koken -eq "http://localhost:$port") -or ($koken -eq "http://127.0.0.1:$port")
      if (-not $ctx.Request.IsLocal -or $ctx.Request.Headers["X-IsTakip"] -ne "1" -or -not $kokenOk) {
        Yaz-Hata $ctx "Erisim reddedildi." 403
        $ctx.Response.OutputStream.Close(); continue
      }
      # Oturum anahtari (guvenlik bulgu 2): sabit baslik ayni bilgisayardaki her programca gonderilebilir;
      # anahtar yalniz bu Windows kullanicisinin dosyasinda durur.
      if (-not (Assist-Esit ([string]$ctx.Request.Headers["X-Assist-Anahtar"]) $ANAHTAR)) {
        Yaz-Hata $ctx "Oturum anahtari gecersiz." 403
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
        elseif ($ctx.Request.ContentLength64 -lt 0 -or $ctx.Request.ContentLength64 -gt 26214400) { Yaz-Hata $ctx "Belge 25 MB'tan buyuk." 413 }   # Bulgu 10: once boyut, sonra okuma
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
          Internet-Isareti $hedef   # Bulgu 3: Office Korumali Gorunum / makro engeli devrede kalsin
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
        if (Outlook-Uygun) { try { $okunmamis = [int](Ns).GetDefaultFolder(6).UnReadItemCount } catch {} }
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
        if (Outlook-Uygun) { try {
          $gk = (Ns).GetDefaultFolder(6); $postaSay = [int]$gk.UnReadItemCount
          $its = $gk.Items.Restrict("[UnRead] = True"); $its.Sort("[ReceivedTime]", $true)
          $it = $its.GetFirst(); $n = 0
          while ($it -and $n -lt 8) {
            $rt = $null; try { $rt = [datetime]$it.ReceivedTime } catch {}
            if ($rt -and $rt -lt $sinir) { break }
            if ($it.Class -eq 43 -or ($it.Class -ge 53 -and $it.Class -le 57)) {   # posta + toplanti iletileri (Outlook sayaciyla ayni)
              # Gonderen (SenderName) OKUNMAZ: Outlook onu korumali sayar, antivirus durumu "guncel degil" gorunurse
              # her 30 sn "A program is trying to access email address information" uyarisi cikiyordu (15.10.2026).
              $posta += @{ saat = $(if ($rt) { & $saatYaz $rt } else { "" }); konu = [string]$it.Subject
                           id = [string]$it.EntryID; store = [string]$it.Parent.StoreID; yeni = [bool]($rt -and $rt -gt $gor) }
              $n++
            }
            $it = $its.GetNext()
          }
        } catch {} }
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

      # ---------- Gelen Kutusu degisti mi? (sayfa ve tepsi birkac sn'de bir yoklar; sunucu 5 sn'de bir bakar) ----------
      # sayfa: is_takip.html'in surumu de doner; sayfa kodu guncellenince acik pencere kendini yeniler
      elseif ($yol -eq "/api/posta/rev") {
        $sz = ""; try { $sz = [string](Get-Item -LiteralPath (Join-Path $webKok "is_takip.html")).LastWriteTimeUtc.Ticks } catch {}
        Yaz-Json $ctx @{ ok = $true; rev = $script:gelenRev; okunmamis = $script:gelenOkunmamis; sayfa = $sz; ortak = $script:ortakRev }
      }

      # ---------- ORTAK IS TAKIBI ----------
      elseif ($yol -eq "/api/ortak" -and $metod -eq "GET") {
        $o = Ortak-Oku; $b = Ben-Kimlik
        $deg = @(foreach ($e in $o.degisen.Keys) { @{ eposta = $e; parmak = $o.degisen[$e].parmak; eski = $(if ($o.anahtarlar[$e]) { $o.anahtarlar[$e].parmak } else { "" }) } })
        $kon = @($o.konular | Select-Object * -ExcludeProperty anahtar)   # konu anahtari sayfaya verilmez
        Yaz-Json $ctx @{ ok = $true; ben = $b; konular = $kon; bekleyen = @($o.giden).Count; kural = $script:ortakKural
          parmak = (Parmak (Benim-Acik)); degisen = $deg; anahtarBekleyen = @($o.bekleyenGiden | ForEach-Object { $_.kime } | Select-Object -Unique).Count }
      }
      # Bir kisinin anahtari degistiyse (yeni bilgisayar) kullanici onaylar; bekleyen girisleri islenir
      elseif ($yol -eq "/api/ortak/anahtar-kabul" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $o = Ortak-Oku; $e = ([string]$g.eposta).ToLowerInvariant()
        if (-not $o.degisen[$e]) { Yaz-Hata $ctx (U "Bu ki{015F}i i{00E7}in bekleyen anahtar de{011F}i{015F}ikli{011F}i yok.") 404 }
        else {
          $o.anahtarlar[$e] = $o.degisen[$e]; [void]$o.degisen.Remove($e)
          $bek = @($o.bekleyenGelen | Where-Object { $_.gonderen -eq $e }); $o.bekleyenGelen = @($o.bekleyenGelen | Where-Object { $_.gonderen -ne $e })
          $n = 0; foreach ($x in $bek) { try { $n += Ortak-Isle $o $x.paket $e } catch {} }
          Bekleyen-Gonder $o $e
          Ortak-Yaz $o
          Yaz-Json $ctx @{ ok = $true; eklenen = $n }
        }
      }
      # Kural elle eklendiyse ya da yeniden denenmek isteniyorsa: tekrar bak / kur
      elseif ($yol -eq "/api/ortak/kural" -and $metod -eq "POST") {
        $script:ortakKuralBakildi = $false
        if (Outlook-Uygun) { Ortak-Kural } else { $script:ortakKural = "" }
        Yaz-Json $ctx @{ ok = $true; kural = $script:ortakKural }
      }
      elseif ($yol -eq "/api/ortak/yeni" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $b = Ben-Kimlik
        $baslik = ([string]$g.baslik).Trim()
        $kat = @(@($g.katilimcilar) | Where-Object { $_.eposta -match '^[^@\s;,]+@[^@\s;,]+\.[A-Za-z]{2,}$' } | ForEach-Object { [pscustomobject]@{ ad = [string]$_.ad; eposta = ([string]$_.eposta).ToLowerInvariant() } })
        $kat = @($kat | Where-Object { $_.eposta -ne $b.eposta } | Sort-Object eposta -Unique)
        if (-not $baslik) { Yaz-Hata $ctx (U "Ba{015F}l{0131}k gerekli.") 400 }
        elseif (-not $kat.Count) { Yaz-Hata $ctx (U "En az bir kat{0131}l{0131}mc{0131} se{00E7}in.") 400 }
        else {
          $o = Ortak-Oku
          $simdi = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
          $acan = [pscustomobject]@{ ad = $b.ad; eposta = $b.eposta }
          $k = [pscustomobject]@{ id = [guid]::NewGuid().ToString("N"); baslik = $baslik; termin = [string]$g.termin; acan = $acan
            katilimcilar = @(@($acan) + $kat); durum = "acik"; olusturma = $simdi; katilma = $simdi; okunan = $simdi; girisler = @() }
          $gl = @(Ortak-Giris "olusturma" ((U "Konuyu a{00E7}t{0131}. Kat{0131}l{0131}mc{0131}lar: ") + (($kat | ForEach-Object { $_.ad }) -join ", ")))
          if (([string]$g.aciklama).Trim()) { Start-Sleep -Milliseconds 5; $gl += Ortak-Giris "not" ([string]$g.aciklama).Trim() }   # ayni anda yazilan iki giris sirasini korusun (ms)
          $k.girisler = $gl
          $o.konular = @($o.konular) + @($k)
          $gitti = Ortak-Gonder $o @($kat | ForEach-Object { $_.eposta }) @{ v = 1; konu = (Ortak-Baslik $k); girisler = @($gl) }
          Ortak-Yaz $o
          Yaz-Json $ctx @{ ok = $true; id = $k.id; gonderildi = $gitti.ok; anahtar = $gitti.anahtar }
        }
      }
      elseif ($yol -eq "/api/ortak/giris" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $o = Ortak-Oku
        $k = $o.konular | Where-Object { $_.id -eq [string]$g.id } | Select-Object -First 1
        $metin = ([string]$g.metin).Trim()
        if (-not $k) { Yaz-Hata $ctx (U "Konu bulunamad{0131}.") 404 }
        elseif ($k.durum -ne "acik") { Yaz-Hata $ctx (U "Konu kapal{0131}; giri{015F} eklenemez.") 400 }
        elseif (-not $metin) { Yaz-Hata $ctx (U "Bo{015F} giri{015F} eklenmez.") 400 }
        else {
          $gr = Ortak-Giris "not" $metin
          $k.girisler = @($k.girisler) + @($gr); $k.okunan = $gr.t
          $gitti = Ortak-Gonder $o @($k.katilimcilar | ForEach-Object { $_.eposta }) @{ v = 1; konu = (Ortak-Baslik $k); girisler = @($gr) }
          Ortak-Yaz $o
          Yaz-Json $ctx @{ ok = $true; gonderildi = $gitti.ok; anahtar = $gitti.anahtar }
        }
      }
      # Yalniz konuyu acan: katilimci ekle / kapat / yeniden ac / gecmisi paylas
      elseif ($yol -eq "/api/ortak/yonet" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $o = Ortak-Oku; $b = Ben-Kimlik
        $k = $o.konular | Where-Object { $_.id -eq [string]$g.id } | Select-Object -First 1
        if (-not $k) { Yaz-Hata $ctx (U "Konu bulunamad{0131}.") 404 }
        elseif (([string]$k.acan.eposta).ToLowerInvariant() -ne $b.eposta) { Yaz-Hata $ctx (U "Bu i{015F}lemi yaln{0131}z konuyu a{00E7}an yapabilir.") 403 }
        else {
          $islem = [string]$g.islem; $gitti = @{ ok = $true; anahtar = 0 }
          if ($islem -eq "katilimci") {
            $mevcut = @{}; foreach ($x in $k.katilimcilar) { $mevcut[([string]$x.eposta).ToLowerInvariant()] = $true }
            $yeni = @(@($g.kisiler) | Where-Object { $_.eposta -match '^[^@\s;,]+@[^@\s;,]+\.[A-Za-z]{2,}$' -and -not $mevcut[([string]$_.eposta).ToLowerInvariant()] } | ForEach-Object { [pscustomobject]@{ ad = [string]$_.ad; eposta = ([string]$_.eposta).ToLowerInvariant() } })
            if (-not $yeni.Count) { Yaz-Hata $ctx (U "Eklenecek yeni ki{015F}i yok.") 400; $ctx.Response.OutputStream.Close(); continue }
            $k.katilimcilar = @($k.katilimcilar) + $yeni
            $gr = Ortak-Giris "katilimci" ((U "Kat{0131}l{0131}mc{0131} ekledi: ") + (($yeni | ForEach-Object { $_.ad }) -join ", "))
            $k.girisler = @($k.girisler) + @($gr)
            # eskiler yalniz bu girisi, yeniler basligi + bu girisi alir (gecmis istenirse ayrica paylasilir)
            $gitti = Ortak-Gonder $o @($k.katilimcilar | ForEach-Object { $_.eposta }) @{ v = 1; konu = (Ortak-Baslik $k); girisler = @($gr) }
          }
          elseif ($islem -eq "kapat" -or $islem -eq "ac") {
            $k.durum = $(if ($islem -eq "kapat") { "kapali" } else { "acik" })
            $gr = Ortak-Giris "durum" $(if ($islem -eq "kapat") { U "Konuyu kapatt{0131}." } else { U "Konuyu yeniden a{00E7}t{0131}." })
            $k.girisler = @($k.girisler) + @($gr)
            $gitti = Ortak-Gonder $o @($k.katilimcilar | ForEach-Object { $_.eposta }) @{ v = 1; konu = (Ortak-Baslik $k); girisler = @($gr) }
          }
          elseif ($islem -eq "gecmis") {
            $kime = ([string]$g.eposta).ToLowerInvariant()
            if (-not (@($k.katilimcilar) | Where-Object { ([string]$_.eposta).ToLowerInvariant() -eq $kime })) { Yaz-Hata $ctx (U "Ki{015F}i bu konunun kat{0131}l{0131}mc{0131}s{0131} de{011F}il.") 400; $ctx.Response.OutputStream.Close(); continue }
            $ad = (@($k.katilimcilar) | Where-Object { ([string]$_.eposta).ToLowerInvariant() -eq $kime } | Select-Object -First 1).ad
            $gr = Ortak-Giris "gecmis" ((U "Ge{00E7}mi{015F}i payla{015F}t{0131}: ") + $ad)
            $k.girisler = @($k.girisler) + @($gr)
            $gitti = Ortak-Gonder $o @($kime) @{ v = 1; konu = (Ortak-Baslik $k); girisler = @($k.girisler) }
            [void](Ortak-Gonder $o @($k.katilimcilar | ForEach-Object { $_.eposta } | Where-Object { $_ -ne $kime }) @{ v = 1; konu = (Ortak-Baslik $k); girisler = @($gr) })
          }
          else { Yaz-Hata $ctx "Bilinmeyen islem." 400; $ctx.Response.OutputStream.Close(); continue }
          $k.okunan = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
          Ortak-Yaz $o
          Yaz-Json $ctx @{ ok = $true; gonderildi = $gitti.ok; anahtar = $gitti.anahtar }
        }
      }
      elseif ($yol -eq "/api/ortak/okundu" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $o = Ortak-Oku
        $k = $o.konular | Where-Object { $_.id -eq [string]$g.id } | Select-Object -First 1
        if ($k) { $son = @($k.girisler | ForEach-Object { [string]$_.t } | Sort-Object)[-1]; if ($son -and $son -gt [string]$k.okunan) { $k.okunan = $son; Ortak-Yaz $o } }
        Yaz-Json $ctx @{ ok = $true }
      }
      # Test sunucusu: baska bir Assist'ten paket gelmis gibi (posta yollamadan birlestirmeyi denemek icin)
      elseif ($yol -eq "/api/ortak/test-al" -and $metod -eq "POST" -and $env:ISTAKIP_PORT) {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $o = Ortak-Oku
        $n = Ortak-Al $o $g.paket ([string]$g.gonderen); Ortak-Yaz $o
        Yaz-Json $ctx @{ ok = $true; eklenen = $n }
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
      elseif ($yol -eq "/api/pencere/ust-kaldir" -and $metod -eq "POST") {   # acilis ekrani bitti: pencere en ustte kalmasin
        & (Join-Path $betikKok "pencere.ps1") -UstKaldir
        Yaz-Json $ctx @{ ok = $true }
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
      # ---------- klasor agaci: Outlook'taki gibi (tum posta kutulari + veri dosyalari / Kisisel Klasorler) ----------
      elseif ($yol -eq "/api/posta/klasorler") {
        $agac = Posta-Agac   # ",@(...)" ile doner: tek depoda da dizi kalir
        Yaz-Cevap $ctx ('{"ok":true,"depolar":' + $(if ($agac.Count) { ConvertTo-Json -InputObject $agac -Depth 30 -Compress } else { '[]' }) + '}') "application/json; charset=utf-8" 200
      }
      # ---------- ceviri ve ses (Azure; anahtar sayfaya geri verilmez) ----------
      elseif ($yol -eq "/api/ceviri/ayar") {
        $gecersiz = $false
        if ($metod -eq "POST") {
          $g = (Govde-Oku $ctx) | ConvertFrom-Json; $a = Ceviri-Ayar
          # Bos alan = degistirme; "-" = sil. Bolge yalniz kucuk harf/rakam (ornek: westeurope)
          foreach ($al in "bolge", "konusmaBolge") { $v = ([string]$g.$al).Trim(); if ($v -and $v.ToLowerInvariant() -notmatch '^[a-z0-9]+$') { $gecersiz = $true } }
          if (-not $gecersiz) {
            foreach ($al in "anahtar", "konusmaAnahtar") { $v = [string]$g.$al; if ($v -eq "-") { $a.$al = "" } elseif ($v.Trim()) { $a.$al = $v.Trim() } }
            foreach ($al in "bolge", "konusmaBolge") { if ($null -ne $g.$al) { $a.$al = ([string]$g.$al).Trim().ToLowerInvariant() } }
            Sifreli-Yaz $ceviriDos ($a | ConvertTo-Json -Compress)
          }
        }
        $a = Ceviri-Ayar
        if ($gecersiz) { Yaz-Hata $ctx "Gecersiz bolge adi (ornek: westeurope)." 400 }
        else { Yaz-Json $ctx @{ ok = $true; ceviri = [bool][string]$a.anahtar; konusma = [bool](([string]$a.konusmaAnahtar -or [string]$a.anahtar) -and ([string]$a.konusmaBolge -or [string]$a.bolge)); bolge = [string]$a.bolge; konusmaBolge = [string]$a.konusmaBolge } }
      }
      elseif ($yol -eq "/api/ceviri" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $m = [string]$g.metin
        if (-not $m.Trim()) { Yaz-Hata $ctx "Cevrilecek metin yok." 400 }
        elseif ($m.Length -gt 60000) { Yaz-Hata $ctx "Metin cok uzun (en cok 60.000 karakter)." 400 }
        else { try { $c = Ceviri-Yap $m ([string]$g.hedef); Yaz-Json $ctx @{ ok = $true; ceviri = $c.ceviri; dil = $c.dil } } catch { Yaz-Hata $ctx ("Ceviri yapilamadi: " + $_.Exception.Message) 502 } }
      }
      # Cevrimdisi ceviri motoru (Bergamot) durumu. ?kontrol=1: Mozilla katalogundan YALNIZ surum listesi okunur
      # (posta/kisisel veri gitmez), kurulu surumle karsilastirilir; sonuc 7 gun saklanir. Indirme/guncelleme YAPILMAZ:
      # bakim sorumlusu Sunucu\bergamot_indir.ps1 + imzala.ps1 ile yapar (kullanici karari 16.10.2026).
      elseif ($yol -eq "/api/bergamot/durum") {
        $sj = Join-Path $webKok "lib\bergamot\surum.json"; $kur = $null
        if (Test-Path -LiteralPath $sj) { try { $kur = Get-Content -LiteralPath $sj -Raw | ConvertFrom-Json } catch {} }
        $kd = Join-Path $veriKok "bergamot_kontrol.json"; $son = $null
        if (Test-Path -LiteralPath $kd) { try { $son = Get-Content -LiteralPath $kd -Raw | ConvertFrom-Json } catch {} }
        $eski = -not $son -or ((Get-Date) - [datetime]::ParseExact([string]$son.tarih, "yyyy-MM-ddTHH:mm", [Globalization.CultureInfo]::InvariantCulture)).TotalDays -ge 7
        if ($kur -and (Sorgu $ctx "kontrol") -eq "1" -and $eski) {
          try {
            Tls12; $K = "https://firefox.settings.services.mozilla.com/v1/buckets/main/collections/"
            $sv = { param($v) try { [version]([string]$v -replace '[^\d.]', '') } catch { [version]"0.0" } }
            $wm = @((Invoke-RestMethod ($K + "translations-wasm/records") -TimeoutSec 30).data) | Where-Object { $_.name -eq "bergamot-translator" -and -not $_.filter_expression } | Sort-Object { & $sv $_.version } -Descending | Select-Object -First 1
            $mm = @((Invoke-RestMethod ($K + "translations-models/records") -TimeoutSec 30).data) | Where-Object { -not $_.filter_expression }
            $yeni = @()
            if ((& $sv $wm.version) -gt (& $sv $kur.motor.surum)) { $yeni += ("motor " + $kur.motor.surum + " > " + $wm.version) }
            foreach ($c in $kur.ciftler.PSObject.Properties.Name) {
              $kv = (@($mm | Where-Object { $_.fromLang -eq $c.Substring(0, 2) -and $_.toLang -eq $c.Substring(2, 2) }) | Sort-Object { & $sv $_.version } -Descending | Select-Object -First 1).version
              if ($kv -and (& $sv $kv) -gt (& $sv $kur.ciftler.$c.surum)) { $yeni += ($c + " " + $kur.ciftler.$c.surum + " > " + $kv) }
            }
            $son = [pscustomobject]@{ tarih = (Get-Date).ToString("yyyy-MM-ddTHH:mm"); yeni = @($yeni) }
            [System.IO.File]::WriteAllText($kd, ($son | ConvertTo-Json -Compress), $UTF8)
          } catch { Write-Host ("  Bergamot surum kontrolu yapilamadi: " + $_.Exception.Message) -ForegroundColor Yellow }
        }
        Yaz-Json $ctx @{ ok = $true; kurulu = [bool]$kur; motor = $(if ($kur) { [string]$kur.motor.surum } else { "" }); ciftler = $(if ($kur) { @{ entr = [string]$kur.ciftler.entr.surum; tren = [string]$kur.ciftler.tren.surum } } else { @{} })
          dosyalar = $(if ($kur) { $kur.ciftler } else { $null }); kontrolTarih = $(if ($son) { [string]$son.tarih } else { "" }); yeni = $(if ($son) { @($son.yeni) } else { @() }) }
      }
      # Outlook'un kendi cevirisi (Microsoft Translator, sirketin M365 lisansi) ile gelen postayi Outlook'un ceviri
      # diline (Turkce) cevirir ve metni Assist'e getirir (16.10.2026, kullanici istegi). COM'da ceviri komutu yok:
      # posta KUCULTULMUS pencerede acilir, "TranslateMessage" komutu calistirilir, cevrilen govde UI Automation ile
      # okunur, pencere KAYDEDILMEDEN kapatilir (Close olDiscard). Outlook ceviriyi hatirlarsa posta zaten cevrili
      # acilir: o zaman komut hata verir, okunan metin dogrudan kullanilir.
      elseif ($yol -eq "/api/posta/outlook-cevir" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = Posta-Bul ([string]$g.id) ([string]$g.store) ([string]$g.msgId)
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else {
          Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
          $AE = [System.Windows.Automation.AutomationElement]; $SC = [System.Windows.Automation.TreeScope]
          $govdeOku = { param($p) foreach ($e in $p.FindAll($SC::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Document)))) { try { return $e.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern).DocumentRange.GetText(-1) } catch {} }; return "" }
          $ins = $null; $sonuc = $null; $hata = ""; $kaynak = ""
          try {
            $ins = $it.GetInspector; $ins.Display($false); $ins.WindowState = 1   # olMinimized
            $kap = [string]$ins.Caption; $pen = $null
            for ($d = 0; $d -lt 20 -and -not $pen; $d++) {
              Start-Sleep -Milliseconds 250
              foreach ($w in $AE::RootElement.FindAll($SC::Children, [System.Windows.Automation.Condition]::TrueCondition)) { if ($kap -and $w.Current.Name.StartsWith($kap.Substring(0, [Math]::Min(20, $kap.Length)))) { $pen = $w; break } }
            }
            if (-not $pen) { throw "Outlook penceresi bulunamadi." }
            $once = ""; for ($d = 0; $d -lt 12 -and -not $once; $d++) { Start-Sleep -Milliseconds 250; $once = & $govdeOku $pen }
            $calisti = $true; try { $ins.CommandBars.ExecuteMso("TranslateMessage") } catch { $calisti = $false }
            $sonra = $once
            if ($calisti) { for ($d = 0; $d -lt 24; $d++) { Start-Sleep -Milliseconds 600; $sonra = & $govdeOku $pen; if ($sonra -and $sonra -ne $once) { break } } }
            # "Translated from: X" yazisindan kaynak dil
            foreach ($e in $pen.FindAll($SC::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text)))) { $m = [regex]::Match([string]$e.Current.Name, '(?i)(translated from|\u015Furadan \u00E7evrildi|cevrildi)\s*:\s*(.+)$'); if ($m.Success) { $kaynak = $m.Groups[2].Value.Trim(); break } }
            if ($sonra -and ($sonra -ne $once -or $kaynak)) { $sonuc = $sonra } else { $hata = "Outlook bu postayi cevirmedi (posta zaten Outlook'un ceviri dilinde olabilir)." }
          } catch { $hata = $_.Exception.Message }
          finally { if ($ins) { try { $ins.Close(1) } catch {} } }
          if ($sonuc) { Yaz-Json $ctx @{ ok = $true; ceviri = $sonuc.Trim(); kaynak = $kaynak } } else { Yaz-Hata $ctx ("Outlook cevirisi alinamadi: " + $hata) 502 }
        }
      }
      # Windows sesli yazma (Win+H): Outlook'taki dikte gibi; odakli metin kutusuna (Assist yanit kutusu) yazar.
      # Sayfa once kutuya odaklanir, sonra bu uc Win+H'ye basar. Anahtar gerekmez; ses Windows'un Microsoft
      # konusma hizmetine gider (Windows Ayarlar > Gizlilik > Konusma). 16.10.2026.
      elseif ($yol -eq "/api/sesli-yazma" -and $metod -eq "POST") {
        if (-not ("AssistTus" -as [type])) {
          Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class AssistTus {
            [DllImport("user32.dll")] static extern void keybd_event(byte k, byte s, uint f, UIntPtr e);
            public static void WinH() { keybd_event(0x5B, 0, 0, UIntPtr.Zero); keybd_event(0x48, 0, 0, UIntPtr.Zero);
              keybd_event(0x48, 0, 2, UIntPtr.Zero); keybd_event(0x5B, 0, 2, UIntPtr.Zero); } }'
        }
        Start-Sleep -Milliseconds 150   # sayfadaki odak kutuya yerlessin
        [AssistTus]::WinH()
        Yaz-Json $ctx @{ ok = $true }
      }
      elseif ($yol -eq "/api/ses-oku" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $m = [string]$g.metin
        if (-not $m.Trim()) { Yaz-Hata $ctx "Okunacak metin yok." 400 }
        elseif ($m.Length -gt 5000) { Yaz-Hata $ctx "Parca cok uzun (en cok 5.000 karakter)." 400 }
        else { try { $b = Ses-Sentez $m ([string]$g.dil); Yaz-Json $ctx @{ ok = $true; ses = [Convert]::ToBase64String($b) } } catch { Yaz-Hata $ctx ("Ses olusturulamadi: " + $_.Exception.Message) 502 } }
      }
      elseif ($yol -eq "/api/konusma" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json; $b64 = [string]$g.ses
        if (-not $b64 -or $b64.Length -gt 4000000) { Yaz-Hata $ctx "Ses kaydi yok ya da cok uzun (en cok ~60 sn)." 400 }
        else { try { $t = Konusma-Yaz ([Convert]::FromBase64String($b64)) ([string]$g.dil); Yaz-Json $ctx @{ ok = $true; metin = $t } } catch { Yaz-Hata $ctx ("Ses yaziya cevrilemedi: " + $_.Exception.Message) 502 } }
      }
      # Outlook yanit imzasi (onizleme icin metin hali). Outlook Reply() imzayi (resimleriyle) kendisi ekler (16.10.2026 denendi).
      elseif ($yol -eq "/api/imza") {
        $ad = ""
        try {
          foreach ($p in Get-ChildItem "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" -ErrorAction Stop) {
            $pk = Join-Path $p.PSPath "9375CFF0413111d3B88A00104B2A6676"
            if (-not (Test-Path $pk)) { continue }
            foreach ($h in Get-ChildItem $pk) {
              $q = Get-ItemProperty $h.PSPath; $r = $q."Reply-Forward Signature"; if (-not $r) { $r = $q."New Signature" }
              if ($r) { $ad = $(if ($r -is [byte[]]) { [Text.Encoding]::Unicode.GetString($r).TrimEnd([char]0) } else { [string]$r }); break }
            }
            if ($ad) { break }
          }
        } catch {}
        $sd = Join-Path $env:APPDATA "Microsoft\Signatures"; $metin = ""; $html = ""
        if ($ad -and -not ($ad.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0)) {
          $tx = Join-Path $sd ($ad + ".txt"); if (Test-Path -LiteralPath $tx) { try { $metin = [IO.File]::ReadAllText($tx).Trim() } catch {} }
          # Gorunum icin imzanin HTML hali (16.10.2026): meta charset'e gore cozulur, <body> ici alinir, imza klasorundeki
          # resimler data: adresine gomulur (sayfa dis/yerel dosyaya erismez). Sayfa betiksiz temizleyip golge DOM'da gosterir.
          $hm = Join-Path $sd ($ad + ".htm")
          if ((Test-Path -LiteralPath $hm) -and (Get-Item -LiteralPath $hm).Length -lt 2MB) {
            try {
              $bayt = [IO.File]::ReadAllBytes($hm); $kod = [Text.Encoding]::UTF8
              $cs = [regex]::Match([Text.Encoding]::ASCII.GetString($bayt, 0, [Math]::Min(4096, $bayt.Length)), 'charset\s*=\s*"?([\w-]+)', 'IgnoreCase').Groups[1].Value
              if ($cs) { try { $kod = [Text.Encoding]::GetEncoding($cs) } catch {} }
              $h = $kod.GetString($bayt)
              $mb = [regex]::Match($h, '(?is)<body[^>]*>(.*)</body>'); if ($mb.Success) { $h = $mb.Groups[1].Value }
              $klasor = Join-Path $sd ($ad + "_files")
              $h = [regex]::Replace($h, '(?i)(\bsrc\s*=\s*")([^"]+)"', {
                param($m)
                $yol = [Uri]::UnescapeDataString($m.Groups[2].Value) -replace '/', '\'
                $dosya = Join-Path $sd $yol
                $tam = [IO.Path]::GetFullPath($dosya)
                $uz = [IO.Path]::GetExtension($tam).TrimStart('.').ToLowerInvariant()
                if ($tam.StartsWith($klasor, [StringComparison]::OrdinalIgnoreCase) -and $RESIM_TUR[$uz] -and (Test-Path -LiteralPath $tam)) {
                  return $m.Groups[1].Value + "data:" + $RESIM_TUR[$uz] + ";base64," + [Convert]::ToBase64String([IO.File]::ReadAllBytes($tam)) + '"'
                }
                return $m.Groups[1].Value + '"'   # imza klasoru disi / bulunamayan kaynak bos birakilir
              }.GetNewClosure())
              if ($h.Length -lt 1500000) { $html = $h.Trim() }
            } catch {}
          }
        }
        Yaz-Json $ctx @{ ok = $true; ad = $ad; metin = $metin; html = $html }
      }
      # Yanit: gonder=true ise Outlook uzerinden GONDERILIR (sayfa once kullanicidan onay alir), yoksa TASLAK acilir.
      # Metin imzanin ve alintinin ustune eklenir; imzayi Outlook kendisi koyar.
      elseif ($yol -eq "/api/posta/yanitla" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = Posta-Bul ([string]$g.id) ([string]$g.store) ([string]$g.msgId)
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else {
          $y = $(if ($g.tumune) { $it.ReplyAll() } else { $it.Reply() })
          $null = $y.GetInspector   # imza ve alinti govdeye yerlessin
          $ek = "<div style=""font-family:Calibri,Arial,sans-serif;font-size:11pt"">" + ([System.Net.WebUtility]::HtmlEncode([string]$g.metin) -replace "\r?\n", "<br>") + "</div><br>"
          $hb = [string]$y.HTMLBody; $m = [regex]::Match($hb, '<body[^>]*>', 'IgnoreCase')
          $y.HTMLBody = $(if ($m.Success) { $hb.Insert($m.Index + $m.Length, $ek) } else { $ek + $hb })
          if ($g.gonder -eq $true) {
            $n = 0; try { $n = [int]$y.Recipients.Count } catch {}
            $y.Send(); Write-Host ("  Yanit gonderildi (" + $n + " alici)")
            Yaz-Json $ctx @{ ok = $true; gonderildi = $true }
          } else { $y.Display($false); Yaz-Json $ctx @{ ok = $true; gonderildi = $false } }
        }
      }
      # ---------- posta listesi ----------
      # kaynak: gelen | bayrakli | kategori (Gelen Kutusu tabanli, "son N gun") ya da klasorId+store (agactan secilen
      # herhangi bir klasor, tarih siniri yok, en yeni 300). kime=1: Gonderilmis/Taslak/Giden icin alicilar da okunur.
      elseif ($yol -eq "/api/posta/liste") {
        $kaynak = Sorgu $ctx "kaynak"; $kategori = Sorgu $ctx "kategori"; $klasorYol = Sorgu $ctx "klasor"
        $klasorId = Sorgu $ctx "klasorId"; $kimeOku = (Sorgu $ctx "kime") -eq "1"
        $gun = [int](Sorgu $ctx "gun"); if ($gun -le 0) { $gun = 14 }
        $enCok = 300
        try {
          if ($klasorId) { $f = (Ns).GetFolderFromID($klasorId, (Sorgu $ctx "store")); $kaynak = "klasor" }
          else { $f = Klasor $klasorYol }
        } catch { Yaz-Hata $ctx ("Klasor bulunamadi: " + $(if ($klasorId) { "secilen klasor" } else { $klasorYol })) 404; $ctx.Response.OutputStream.Close(); continue }
        $items = $f.Items
        # "[FlagStatus] = 2" bu posta kutusunda HIC sonuc dondurmuyor (13.10.2026,
        # Exchange); "<> 0" calisiyor. Tamamlanmis bayraklar (1) asagida elenir.
        if ($kaynak -eq "bayrakli") { $items = $items.Restrict("[FlagStatus] <> 0") }
        elseif ($kaynak -eq "kategori" -and $kategori) { $items = $items.Restrict("[Categories] = '" + $kategori.Replace("'", "''") + "'") }
        # Taslak / Giden'de alinma zamani yok: son degisiklige gore siralanir
        $taslakTur = $false; try { $taslakTur = $f.EntryID -eq $f.Store.GetDefaultFolder(16).EntryID -or $f.EntryID -eq $f.Store.GetDefaultFolder(4).EntryID } catch {}
        $items.Sort($(if ($taslakTur) { "[LastModificationTime]" } else { "[ReceivedTime]" }), $true)
        $sinir = (Get-Date).Date.AddDays(-$gun)
        $ekHarita = $(if ($taslakTur) { @{} } else { Ek-Haritasi $f 350 })   # gercek ek isareti (satir ici resimler haric); taslakta tabloda alinma zamani yok
        $liste = New-Object System.Collections.ArrayList
        $n = $items.Count
        for ($i = 1; $i -le $n -and $i -le 2000 -and $liste.Count -lt $enCok; $i++) {
          $it = $items.Item($i)
          # Posta (43) ve toplanti iletileri (53-57: davet, iptal, kabul/ret/belki yaniti). Outlook'un okunmamis
          # sayisi toplanti iletilerini de sayar; listede olmayinca sayilar tutmuyordu (15.10.2026).
          $sinif = [int]$it.Class
          if ($sinif -ne 43 -and ($sinif -lt 53 -or $sinif -gt 57)) { continue }
          if ($kaynak -eq "bayrakli" -and $it.FlagStatus -ne 2) { continue }
          # Assist'in tasiyici iletileri (talep / ortak is) listede gosterilmez (16.10.2026): Gonderilmis'i dolduruyordu
          if (([string]$it.MessageClass) -like "IPM.Note.Assist*" -or ([string]$it.Subject).StartsWith("[Assist] Talep:") -or ([string]$it.Subject).StartsWith("[Assist] Ortak i")) { continue }
          $alindi = [datetime]$it.ReceivedTime
          if ($taslakTur -or $alindi.Year -gt 4000) { try { $alindi = [datetime]$it.LastModificationTime } catch {} }
          # Bayrakli postalar da "son N gun" ile sinirli: posta kutusunda yillar once
          # bayrak konmus yuzlerce posta var, listeyi onlar doldurmasin (13.10.2026).
          # Kategori ve agactan secilen klasorde sinir yok (en yeni 300).
          if ($kaynak -ne "kategori" -and $kaynak -ne "klasor" -and $alindi -lt $sinir) { break }
          $on = ""
          try { $on = ([string]$it.Body) -replace '\s+', ' '; if ($on.Length -gt 220) { $on = $on.Substring(0, 220) } } catch {}
          $o = @{
            id = [string]$it.EntryID; store = [string]$f.StoreID
            konu = [string]$it.Subject; kimden = [string]$it.SenderName
            tarih = (Tarih-Iso $alindi); on = $on.Trim()
            bayrak = ($it.FlagStatus -eq 2); bayrakTermin = (Bayrak-Termin $it)
            okunmadi = [bool]$it.UnRead; ek = $(if ([int]$it.Attachments.Count -eq 0) { 0 } elseif ($ekHarita.ContainsKey([string]$it.EntryID)) { [int]$ekHarita[[string]$it.EntryID] } else { Ek-Say $it }); yz = (Yazisma-Sira $it)
            kategori = [string]$it.Categories; toplanti = ($sinif -ne 43)
          }
          if ($kimeOku) {
            # To bos olabilir (toplanti yaniti, yalniz BCC): alici listesinden doldurulur. Adlardaki tirnaklar temizlenir.
            $kime = ""; try { $kime = [string]$it.To } catch {}
            if (-not $kime.Trim()) { try { $kime = (@(foreach ($r in $it.Recipients) { [string]$r.Name }) -join "; ") } catch {} }
            $o.kime = (@($kime.Split(';') | ForEach-Object { $_.Trim().Trim("'", '"').Trim() } | Where-Object { $_ }) -join "; ")
          }
          $o.kimden = $o.kimden.Trim().Trim("'", '"').Trim()
          [void]$liste.Add($o)
        }
        Yaz-Json $ctx @{ ok = $true; klasor = [string]$f.FolderPath; okunmamis = [int]$f.UnReadItemCount; postalar = @($liste) }
      }

      # ---------- tek posta (tam metin) ----------
      elseif ($yol -eq "/api/posta/oku") {
        $it = Posta-Bul (Sorgu $ctx "id") (Sorgu $ctx "store") (Sorgu $ctx "msgId")
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else { Yaz-Json $ctx (Posta-Ayrinti $it) }
      }
      # ---------- ayni yazismadaki diger iletiler (posta penceresinin sag bolmesi) ----------
      elseif ($yol -eq "/api/posta/yazisma") {
        $it = Posta-Bul (Sorgu $ctx "id") (Sorgu $ctx "store") ""
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else { Yaz-Json $ctx @{ ok = $true; iletiler = @(Posta-Yazisma $it) } }
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
      # Assist'te acilan posta penceresi kapatilinca o posta Outlook'ta okundu yapilir (kullanici istedi, 15.10.2026)
      elseif ($yol -eq "/api/posta/okundu-tek" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = Posta-Bul ([string]$g.id) ([string]$g.store) ([string]$g.msgId)
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else { if ($it.UnRead) { $it.UnRead = $false; $it.Save() }; Yaz-Json $ctx @{ ok = $true; kalan = [int]((Ns).GetDefaultFolder(6).UnReadItemCount) } }
      }
      elseif ($yol -eq "/api/posta/okundu" -and $metod -eq "POST") {
        # Govdede klasorId+store varsa o klasor (agactan secilen), yoksa Gelen Kutusu
        $g = $null; try { $g = (Govde-Oku $ctx) | ConvertFrom-Json } catch {}
        $gk = (Ns).GetDefaultFolder(6); $n = 0
        if ($g -and [string]$g.klasorId) { try { $gk = (Ns).GetFolderFromID([string]$g.klasorId, [string]$g.store) } catch { Yaz-Hata $ctx "Klasor bulunamadi." 404; $ctx.Response.OutputStream.Close(); continue } }
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
      # ---------- postayi sil (16.10.2026, kullanici istedi) ----------
      # Delete() postayi Outlook'ta "Silinmis Ogeler"e TASIR (geri alinabilir).
      # Posta zaten Silinmis Ogeler'deyse Delete() KALICI siler: o durumda reddedilir.
      elseif ($yol -eq "/api/posta/sil" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = Posta-Bul ([string]$g.id) ([string]$g.store) ([string]$g.msgId)
        if (-not $it) { Yaz-Hata $ctx "Posta bulunamadi." 404 }
        else {
          $copte = $false
          try { $copte = ($it.Parent.EntryID -eq $it.Parent.Store.GetDefaultFolder(3).EntryID) } catch {}
          if ($copte) { Yaz-Hata $ctx "Posta zaten Silinmis Ogeler'de; kalici silme Assist'ten yapilmaz." 400 }
          else {
            $it.Delete()
            Write-Host "  Posta: Silinmis Ogeler'e tasindi"
            Yaz-Json $ctx @{ ok = $true }
          }
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
          # G6/G7 (16.10.2026): talep uctan uca sifreli + imzali gider (ortak isteki altyapi).
          # Alicinin acik anahtari yoksa istenir ve talep kuyruga alinir; dakikalik kontrolde yollanir.
          $o = Ortak-Oku; $alici = $kime.ToLowerInvariant()
          if (-not $o.anahtarlar[$alici]) {
            Anahtar-Iste $o $alici
            $o.bekleyenTalep = @($o.bekleyenTalep) + @([pscustomobject]@{ kime = $alici; talep = $g.talep; konu = [string]$g.konu; posta = $g.posta; t = (Get-Date).ToString("o") })
            Ortak-Yaz $o
            Write-Host ("  Talep kuyrukta: " + $alici + " guvenlik anahtari bekleniyor")
            Yaz-Json $ctx @{ ok = $true; beklemede = $true }
          } else {
            try { Talep-Yolla $o $alici $g.talep ([string]$g.konu) $g.posta; Write-Host "  Talep gonderildi (sifreli Assist iletisi)"; Yaz-Json $ctx @{ ok = $true } }
            catch { Write-Host ("  Talep gonderilemedi: " + $_.Exception.Message) -ForegroundColor Yellow; Yaz-Hata $ctx ("Talep gonderilemedi: " + $_.Exception.Message) 500 }
          }
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

      # ---------- postanin ekini ac (16.10.2026) ----------
      # Ek Outlook'taki postadan alinir; Outlook'un kendi gecici ek klasorune yazilip varsayilan
      # programla acilir (Outlook'ta eke cift tiklamakla ayni). Assist'te ek saklanmaz.
      # Belge turu disindaki ekler (exe, betik vb.) acilmaz: posta Outlook'ta acilir, Outlook'un
      # kendi guvenlik denetimi gecerli olur.
      elseif ($yol -eq "/api/posta/ek-ac" -and $metod -eq "POST") {
        $g = (Govde-Oku $ctx) | ConvertFrom-Json
        $it = $null; try { $it = Posta-Bul ([string]$g.id) ([string]$g.store) ([string]$g.msgId) } catch {}
        $e = $null; if ($it) { try { $e = $it.Attachments.Item([int]$g.i) } catch {} }
        if (-not $e) { Yaz-Hata $ctx "Ek Outlook'ta bulunamadi (posta tasinmis ya da silinmis olabilir)." 404 }
        else {
          $ad = ([IO.Path]::GetFileName([string]$e.FileName)) -replace '[\\/:*?"<>|]', '_'
          $uz = ([IO.Path]::GetExtension($ad)).ToLowerInvariant()
          if (@(1, 5) -notcontains [int]$e.Type -or $BELGE_UZ -notcontains $uz) {
            $it.Display($false); try { $it.GetInspector.Activate() } catch {}
            Yaz-Json $ctx @{ ok = $true; kaynak = "outlook" }
          } else {
            $kok = Outlook-Ek-Klasoru
            # Outlook, kendi gecici klasorunde daha once kullandigi adi (dosya silinmis olsa bile) ayirmis tutabiliyor:
            # SaveAsFile "This file name is in use by Outlook for a temporary file" hatasi verir (16.10.2026). O zaman ad
            # sonuna (1), (2)... eklenip yeniden denenir.
            $tem = [IO.Path]::GetFileNameWithoutExtension($ad); $hedef = $null
            for ($n = 0; $n -lt 30 -and -not $hedef; $n++) {
              $aday = Join-Path $kok ($(if ($n) { $tem + " (" + $n + ")" } else { $tem }) + $uz)
              if (Test-Path -LiteralPath $aday) { continue }
              try { $e.SaveAsFile($aday); $hedef = $aday } catch { if ($n -ge 29) { throw } }
            }
            if (-not $hedef) { throw "Ek kaydedilemedi." }
            Internet-Isareti $hedef
            Start-Process -FilePath $hedef
            Yaz-Json $ctx @{ ok = $true; kaynak = "ek" }
          }
        }
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
            Internet-Isareti $yolMsg   # Bulgu 3
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
    if ($yol -eq "/favicon.ico") { $tam = Join-Path $portalKok "AssistA.ico" }   # uygulama penceresinin gorev cubugu simgesi
    if (-not ($tam.StartsWith($webKok, [StringComparison]::OrdinalIgnoreCase) -or $yol -eq "/favicon.ico") -or -not (Test-Path -LiteralPath $tam -PathType Leaf)) {
      Yaz-Cevap $ctx "Bulunamadi" "text/plain; charset=utf-8" 404
    } elseif ($GUV.imzaZorunlu -and $yol -ne "/favicon.ico" -and -not (& {
        # Sayfa imzali katalogla (Sunucu\assist.cat) dogrulanir; sonuc dosya zamanlari degismedikce onbellekte (guvenlik bulgu 1)
        $cat = Join-Path $betikKok "assist.cat"
        $an = (Get-Item -LiteralPath $tam).LastWriteTimeUtc.Ticks.ToString() + "|" + $(if (Test-Path -LiteralPath $cat) { (Get-Item -LiteralPath $cat).LastWriteTimeUtc.Ticks } else { "" })
        if ($an -ne $script:katalogAnahtar) { $script:katalogOk = Assist-Katalog-Gecerli $portalKok $GUV.parmakIzi; $script:katalogAnahtar = $an }
        $script:katalogOk })) {
      Write-Host ("  GUVENLIK: " + $yol + " imzali katalogla dogrulanamadi - sunulmadi.") -ForegroundColor Red
      Yaz-Cevap $ctx "Assist sayfasi dogrulanamadi (imza/katalog gecersiz). Yoneticiye bildirin." "text/plain; charset=utf-8" 503
    } else {
      $b = [System.IO.File]::ReadAllBytes($tam)
      $ctx.Response.ContentType = Icerik-Turu ([System.IO.Path]::GetExtension($tam))
      $ctx.Response.Headers.Add("Cache-Control", "no-store")
      # Cevrimdisi ceviri iscisi (Bergamot): AG IZNI YOK. connect-src tanimsiz -> default-src 'none': isci hicbir adrese
      # (Assist dahil) istek atamaz; yalniz kendi kokunden betik yukler ve WebAssembly derler (16.10.2026).
      if ($yol -eq "/ceviri-isci.js") { $ctx.Response.Headers.Add("Content-Security-Policy", "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'") }
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

# SIG # Begin signature block
# MIIdkQYJKoZIhvcNAQcCoIIdgjCCHX4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDvvvZk4EnPD9UP
# 7fUrbktAUFatNAQtqJqYao4FdHhgAaCCF1QwggQWMIICfqADAgECAhA2m8l6HI/0
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEICparbn77gUL
# s3jjMJLIf6tOXeOmM8VdvYbHSHnoLKTeMA0GCSqGSIb3DQEBAQUABIIBgE+lyujd
# Nt3djzqRvXAmSwV0lqu2YgDHS0jF+iMkMWz7msFyxB3mRQ/1mN1US4AS5AsZkfVw
# OrPS/9cN9ugBOe3R/5UC9/tQpNF+mAHXHm89w1TUJJm57TR3Y51WvoimRjE+Px0+
# kTy3Yxhs53MVQ4ELdvfS7cgtIwZ05rlLCwy14WWtTEjz7SLxPHGgF5kQw3qc+2R/
# pE0ybl15wChqxUMARIyIARuZK1A3/3d1WUjMPSfE9cgncYl9YmPyoooJ3iIfwzGo
# 4DNuVFNTYY3DReRrOBp04UdgYf22fw3VMgpvVsH8GHxlQxRl1gpfeTQB3cJhQ/qf
# EW0twMO330tGmSTZIDM1LMpVoDsKOLjOyIogTxX2yum64zZY3Uco8pfZnwfjt0J6
# 0FrpbnRYd5JuNpwgGjGEPeBh4V6/14WfdVSPNgApKEYVAaRHTIzy2e1PlyJV2Vc0
# qnboOZuwcpRgW5JtfDMKwN4b4bnpve6x0xhMoqA8lR8sxAsRlK8bFvCIYKGCAyYw
# ggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYD
# VQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBH
# NCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO
# 28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTYxMTI4NDZaMC8GCSqGSIb3DQEJBDEi
# BCCZ8gBhHMg/VEZeYCQrwkgTtcpwzDsNy+6uMHkBA0brlDANBgkqhkiG9w0BAQEF
# AASCAgBJoGBSBrA8ba9NIsF5d/4GK5l8RB/dGaGeGZKUXm+hRYu6bwaU8n/n18FI
# PdpqLHyPn/WAoQDndsZmzLwPaB+tQAUyIl9MzDzJ3/tcirTdJWtWbuy8HtU9bOCF
# QP1Tr3/Sc/dqrFyepAEkl8vqzTUwhx8ZLgtMQT7IBT9UxUDD+/y8GU9M6nU/q3uW
# HQHiGWca5hGOYZThbVt0Z7j8DNzZ/bpStT9eon9FKuk39w2rlevvict1lwpMt6Ea
# OF7X/A47JarTlqbK5JP2pA270k2jfKcip0yP4Aaycoz6cL92c8IH/8Uee74mqTzJ
# XTZ9b3D8GR6Dz862jKHGL3wpkP9m5pi8pyfJHmQ7gF8vgbiO6bIs6tr79xEN27pk
# r/fDG9aSs4GvL/trayNienResW/08eJy+fBKcLQ2f5PvDd4vkgXOc5PcH54QfBua
# 0E2rvhT/ughgqsQ2z4f8o2ul92eZCPofG/0JwykWvFoxT4Icsznv0fIfXE1L5HQK
# 9ZIfJrBvWOFqulPXurFeyunBtqOw43HsNVgy3U/zaMfEkdKqYVkee/GD1JBTzI8c
# QXlh/Naw9JFjIFAhrBwXmaaMDQsC+1KuDTbFvNXNN0tk1XxYLxUBrt/dCkpETKRQ
# h/WBYrYaW9gJGaLPrS0cT2+PWt2eeA37T8QcTGMfMXEg9BTXOA==
# SIG # End signature block
