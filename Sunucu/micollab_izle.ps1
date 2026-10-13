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
$log = Join-Path $PSScriptRoot ("_tepsi_" + $env:COMPUTERNAME + ".log")
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
$sayac = 0; $gizleBitis = [DateTime]::MinValue
while ($true) {
  Start-Sleep -Seconds 2
  $sayac++
  $yas = 1e9
  try { $yas = ((Get-Date) - (Get-Process -Name UCA -ErrorAction Stop | Select-Object -First 1).StartTime).TotalSeconds } catch {}
  if ($yas -lt 300 -or ($sayac % 30 -eq 0)) { if (Acil-Pencere-Kapat) { $gizleBitis = (Get-Date).AddSeconds(30) } }
  if ($yas -lt 90 -or (Get-Date) -lt $gizleBitis) { Liste-Gizle }
}
