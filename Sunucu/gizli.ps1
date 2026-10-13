# =====================================================================
#  IS TAKIP - PENCERESIZ BASLATICI
#  ---------------------------------------------------------------
#  Bir PowerShell betigini HIC PENCERE ACMADAN calistirir.
#  Neden: Windows 11'de varsayilan konsol "Windows Terminal" olunca
#  "-WindowStyle Hidden" dikkate alinmiyor ve sunucu / tepsi simgesi
#  bos birer pencere olarak gorunuyordu (13.10.2026). Islem burada
#  CreateNoWindow ile, konsol penceresi olusturulmadan baslatilir.
# =====================================================================
param([Parameter(Mandatory = $true)][string]$Betik, [int]$Port = 0)
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "powershell.exe"
$psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $Betik + '"' + $(if ($Port) { " -Port $Port" } else { "" })
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
$psi.WorkingDirectory = Split-Path -Parent $Betik
[void][System.Diagnostics.Process]::Start($psi)
