<#
  ============================================================================
   KARDEX FIMAS - Desinstalador (Windows Server)
  ============================================================================
   Detiene y elimina el servicio (tarea programada), quita la regla de firewall
   y, opcionalmente, borra los archivos.

   Uso (PowerShell como Administrador):
     .\uninstall.ps1
     .\uninstall.ps1 -BorrarTodo
  ============================================================================
#>
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string] $Destino = "C:\Kardex-ERP",
    [int]    $Puerto  = 8765,
    [string] $Tarea   = "KardexERP",
    [switch] $BorrarTodo
)

$ErrorActionPreference = "Stop"

function Linea { Write-Host "  -----------------------------------------------------------" -ForegroundColor DarkGray }
function Ok($m)   { Write-Host "  [OK] $m" -ForegroundColor Green }
function Warn($m) { Write-Host "  [!]  $m" -ForegroundColor Yellow }

Write-Host ""
Write-Host "  ===========================================================" -ForegroundColor Cyan
Write-Host "        KARDEX FIMAS - Desinstalacion" -ForegroundColor Cyan
Write-Host "  ===========================================================" -ForegroundColor Cyan
Write-Host ""
Linea

# Detener y eliminar la tarea programada
if (Get-ScheduledTask -TaskName $Tarea -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $Tarea -ErrorAction SilentlyContinue
    schtasks /End /TN "$Tarea" 2>$null | Out-Null
    Start-Sleep -Seconds 1
    schtasks /Delete /TN "$Tarea" /F | Out-Null
    Ok "Tarea programada '$Tarea' eliminada."
} else {
    Warn "No existe la tarea programada '$Tarea'."
}

# Matar procesos de la app que sigan vivos
Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like "*kardex_server.py*" } |
    ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force; Ok "Proceso detenido (PID $($_.ProcessId))." } catch {} }

# Quitar regla de firewall
$regla = "Kardex Fimas $Puerto"
if (Get-NetFirewallRule -DisplayName $regla -ErrorAction SilentlyContinue) {
    try { Remove-NetFirewallRule -DisplayName $regla -ErrorAction Stop; Ok "Regla de firewall eliminada." }
    catch { netsh advfirewall firewall delete rule name="$regla" | Out-Null; Ok "Regla de firewall eliminada (netsh)." }
}

# Archivos
if ($BorrarTodo) {
    if (Test-Path $Destino) {
        Remove-Item -LiteralPath $Destino -Recurse -Force
        Ok "Carpeta eliminada: $Destino"
    }
} else {
    Warn "Los archivos NO se borraron: $Destino  (use -BorrarTodo para eliminarlos)."
    Warn "Se conservaron las bases de datos (.db) y la configuracion (.env)."
}

Write-Host ""
Ok "Desinstalacion finalizada."
Write-Host ""
