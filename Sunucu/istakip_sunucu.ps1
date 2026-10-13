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
$ayarDos = Join-Path $betikKok ("istakip_ayar." + $env:COMPUTERNAME + ".json")
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

function Xml-Kacis([string]$s) { return [System.Security.SecurityElement]::Escape($s) }
function Bildirim-Goster([string]$baslik, [string]$metin, [string]$hedef) {
  if (-not $script:bildirimHazir) { return }
  $url = "http://localhost:$port/?b=" + [DateTime]::Now.Ticks + "#" + $hedef
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
# Ad degisikligi (Is Takip -> Assist, 13.10.2026): eski otomatik baslat kisayolu yeni bat'a tasinir.
try {
  $eskiLnk = Join-Path ([Environment]::GetFolderPath("Startup")) "Is Takip.lnk"
  if (Test-Path -LiteralPath $eskiLnk) {
    $k = (New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path ([Environment]::GetFolderPath("Startup")) "Assist.lnk"))
    $k.TargetPath = Join-Path $portalKok "Assist.bat"; $k.WorkingDirectory = $portalKok; $k.WindowStyle = 7; $k.Description = "Assist"; $k.Save()
    Remove-Item -LiteralPath $eskiLnk -Force
  }
} catch {}
# Outlook/MiCollab/Teams'i bat baslatir (Sunucu\uygulamalar.ps1): sunucu zaten aciksa da calissin diye.

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
      try { Gecici-Temizle } catch {}
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
      Write-Host ("  > " + (Get-Date).ToString("HH:mm:ss") + " " + $metod + " " + $yol + $(if ($yol -ne "/api/kisi/coz") { $ctx.Request.Url.Query }))   # kisi adi kayda yazilmaz
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
        } else { Yaz-Hata $ctx "Gecersiz istek." 400 }
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

      # ---------- Windows acilisinda otomatik baslat ----------
      # Kullanicinin Baslangic klasorune bat dosyasinin kisayolu konur / kaldirilir.
      # Bat sunucuyu ve tarayiciyi acar; sunucu tarayicidan bagimsiz calisir.
      elseif ($yol -eq "/api/otobaslat") {
        $lnk = Join-Path ([Environment]::GetFolderPath("Startup")) "Assist.lnk"
        if ($metod -eq "POST") {
          $g = (Govde-Oku $ctx) | ConvertFrom-Json
          if ($g.ac) {
            $ws = New-Object -ComObject WScript.Shell
            $k = $ws.CreateShortcut($lnk)
            $k.TargetPath = Join-Path $portalKok "Assist.bat"
            $k.WorkingDirectory = $portalKok
            $k.WindowStyle = 7   # simge durumunda
            $k.Description = "Assist"
            $k.Save()
          } elseif (Test-Path -LiteralPath $lnk) { Remove-Item -LiteralPath $lnk -Force }
        }
        Yaz-Json $ctx @{ ok = $true; acik = (Test-Path -LiteralPath $lnk) }
      }

      # ---------- toplanti odalari: Outlook "All Rooms" listesi + doluluk (30 dk dilim) ----------
      # Doluluk her odanin Outlook serbest/mesgul bilgisinden okunur: 0 bos, 1 belirsiz, 2 dolu, 3 disarida.
      elseif ($yol -eq "/api/oda" -and $metod -eq "GET") {
        $gun = [datetime]::Today
        try { $gun = [datetime]::ParseExact((Sorgu $ctx "gun"), "yyyy-MM-dd", $null) } catch {}
        $L = New-Object System.Collections.ArrayList
        foreach ($ae in (Oda-Listesi)) {
          $dilim = ""; try { $dilim = ([string]$ae.GetFreeBusy($gun, 30, $true)).Substring(0, 48) } catch {}
          $ep = ""; try { $ep = [string]$ae.GetExchangeUser().PrimarySmtpAddress } catch {}
          [void]$L.Add(@{ ad = [string]$ae.Name; eposta = $ep; dilim = $dilim })
        }
        Yaz-Json $ctx @{ ok = $true; gun = $gun.ToString("yyyy-MM-dd"); odalar = @($L) }
      }
      # Rezervasyon: odanin KAYNAK olarak eklendigi toplanti daveti Outlook'ta HAZIR acilir.
      # Gonder'e kullanici basar; Assist daveti asla kendisi gondermez (talep iletme ile ayni kural).
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
          [void]$t.Recipients.ResolveAll()
          $t.Display()
          Yaz-Json $ctx @{ ok = $true }
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
    if (-not $tam.StartsWith($webKok, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $tam -PathType Leaf)) {
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
