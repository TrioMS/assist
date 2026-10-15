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
#  Dosya ASCII'dir.
# =====================================================================
param([string]$Hedef = "", [string]$Url = "", [int]$Port = 8096, [switch]$Kapat, [switch]$Ilk)

$BASLIK = "STFA Assist"   # Portal\is_takip.html <title> ile ayni olmali (pencere bu adla bulunur)

if (-not ('AssistPencere2' -as [type])) {
  Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public static class AssistPencere2 {
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
  public static void Kapat(IntPtr h) { PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); }   // WM_CLOSE
}
"@
}

$h = [AssistPencere2]::Bul($BASLIK)
if ($Kapat) { if ($h -ne [IntPtr]::Zero) { [AssistPencere2]::Kapat($h) }; return }

if ($Url -match '^assist:(.*)$') { $Hedef = [Uri]::UnescapeDataString($Matches[1].TrimStart('/')) }
$adres = "http://localhost:$Port/"

if ($h -ne [IntPtr]::Zero) {
  if ($Hedef) {
    try {
      $r = [System.Net.WebRequest]::Create($adres + "api/pencere/hedef")
      $r.Method = "POST"; $r.ContentType = "application/json"; $r.Timeout = 5000; $r.Headers.Add("X-IsTakip", "1")
      $b = [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json @{ hedef = $Hedef } -Compress))
      $r.ContentLength = $b.Length; $s = $r.GetRequestStream(); $s.Write($b, 0, $b.Length); $s.Close(); $r.GetResponse().Close()
    } catch {}
  }
  [AssistPencere2]::OneGetir($h)
  return
}

$u = $adres + "?app=1" + $(if ($Ilk) { "&ilk=1" } else { "" }) + "&v=" + [DateTime]::Now.Ticks + $(if ($Hedef) { "#" + [Uri]::EscapeDataString($Hedef) } else { "" })
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
    $h = [AssistPencere2]::Bul($BASLIK)
    if ($h -ne [IntPtr]::Zero) { [AssistPencere2]::Buyut($h); [AssistPencere2]::OneGetir($h); break }
  }
}
else { try { Start-Process $u } catch { Start-Process explorer.exe -ArgumentList $u } }   # Edge/Chrome yoksa normal tarayici sekmesi
