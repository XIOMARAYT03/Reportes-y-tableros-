<#
  ============================================================================
   KARDEX FIMAS - Instalador para Windows Server
  ============================================================================
   Copia la aplicacion, instala dependencias y la registra como TAREA
   PROGRAMADA que arranca al iniciar el servidor (sin ventanas, con reinicio
   automatico ante fallos). No requiere software de terceros.

   Uso (PowerShell como Administrador):
     .\install.ps1
     .\install.ps1 -Destino "D:\Apps\Kardex" -Puerto 8765
     .\install.ps1 -EnvPath "D:\Config\kardex.env" -SinFirewall

   Ejecute .\install.ps1 -? para ver todos los parametros.
  ============================================================================
#>
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string] $Destino   = "C:\Kardex-ERP",
    [int]    $Puerto    = 8765,
    [string] $EnvPath   = "",
    [string] $Tarea     = "KardexERP",
    [switch] $SinFirewall,
    [switch] $SinIniciar
)

$ErrorActionPreference = "Stop"
$script:Fallos = 0

function Linea { Write-Host "  -----------------------------------------------------------" -ForegroundColor DarkGray }
function Info($m) { Write-Host "  $m" -ForegroundColor Gray }
function Ok($m)   { Write-Host "  [OK] $m" -ForegroundColor Green }
function Warn($m) { Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Fail($m) { Write-Host "  [ERROR] $m" -ForegroundColor Red; exit 1 }
function Titulo($m) {
    Write-Host ""
    Write-Host "  $m" -ForegroundColor Cyan
    Linea
}

function PuertoActivo([int]$p) {
    try {
        $c = New-Object System.Net.Sockets.TcpClient
        $c.Connect("127.0.0.1", $p)
        $c.Close()
        return $true
    } catch { return $false }
}

Write-Host ""
Write-Host "  ===========================================================" -ForegroundColor Cyan
Write-Host "        KARDEX FIMAS - Instalacion en Windows Server" -ForegroundColor Cyan
Write-Host "  ===========================================================" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
Titulo "1/8  Verificando privilegios"
$esAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $esAdmin) {
    Fail "Ejecute este instalador como Administrador (clic derecho > 'Ejecutar como administrador')."
}
Ok "Permisos de administrador confirmados."

# ---------------------------------------------------------------------------
Titulo "2/8  Localizando la aplicacion"
$appDir = Join-Path (Split-Path -Parent $PSScriptRoot) "app"
if (-not (Test-Path (Join-Path $appDir "kardex_server.py"))) {
    if (Test-Path (Join-Path $PSScriptRoot "kardex_server.py")) {
        $appDir = $PSScriptRoot
    } else {
        Fail "No se encontro 'kardex_server.py'. Ejecute el instalador desde la carpeta 'deploy' del paquete o desde la carpeta de la aplicacion."
    }
}
Ok "Aplicacion: $appDir"

# ---------------------------------------------------------------------------
Titulo "3/8  Verificando Python"
$py = Get-Command python -ErrorAction SilentlyContinue
$pyArgs = ""
if ($py) {
    $pyExe = $py.Source
} else {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { $pyExe = $py.Source; $pyArgs = "-3" }
    else {
        Fail "No se encontro Python. Instale Python 3.10 o superior desde https://www.python.org/downloads/windows/ y marque 'Add python.exe to PATH'."
    }
}
try {
    $verTxt = (& $pyExe $pyArgs --version 2>&1) -join " "
} catch {
    Fail "No se pudo ejecutar Python ($pyExe). Revise la instalacion."
}
Info "Version detectada: $verTxt"
Ok "Interprete: `"$pyExe`" $pyArgs"

# ---------------------------------------------------------------------------
Titulo "4/8  Deteniendo una instalacion previa (si existe)"
if (Get-ScheduledTask -TaskName $Tarea -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $Tarea -ErrorAction SilentlyContinue
    schtasks /End /TN "$Tarea" 2>$null | Out-Null
    Ok "Instancia existente detenida."
} else {
    Info "No habia instalacion previa registrada."
}
Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like "*kardex_server.py*" -and $_.CommandLine -like "*$Destino*" } |
    ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force } catch {} }

# ---------------------------------------------------------------------------
Titulo "5/8  Copiando archivos a $Destino"
New-Item -ItemType Directory -Force -Path $Destino | Out-Null
# Nunca sobrescribimos bases de datos ni logs del servidor (se excluyen *.db).
$rc = Start-Process robocopy -Wait -NoNewWindow -PassThru -ArgumentList @(
    "`"$appDir`"", "`"$Destino`"", "/E",
    "/XF", "*.db", "*.log", "*.log.err", "*.pyc", "*.pyo",
    "/XD", "__pycache__", ".git", "exports",
    "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:2", "/W:2"
)
if ($rc.ExitCode -ge 8) { Fail "Fallo la copia de archivos (robocopy codigo $($rc.ExitCode))." }
Ok "Archivos de la aplicacion copiados."

# ---------------------------------------------------------------------------
Titulo "6/8  Instalando dependencias de Python"
& $pyExe $pyArgs -m pip install --upgrade pip --quiet 2>$null
& $pyExe $pyArgs -m pip install -r (Join-Path $Destino "requirements.txt") --quiet
if ($LASTEXITCODE -ne 0) { Warn "Revise la instalacion de dependencias (pip). La app podria no iniciar." }
else { Ok "Dependencias instaladas (openpyxl)." }

# ---------------------------------------------------------------------------
Titulo "7/8  Configuracion de credenciales (.env)"
if ($EnvPath) {
    $envDest = $EnvPath
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $envDest) | Out-Null
} else {
    $envDest = Join-Path $Destino ".env"
}
if (Test-Path $envDest) {
    Ok "Configuracion existente conservada: $envDest"
} else {
    Copy-Item (Join-Path $Destino ".env.example") $envDest -Force
    Warn "Se creo $envDest con valores de EJEMPLO."
    Warn "Edite ese archivo con las credenciales reales de Odoo ANTES de usar el sistema."
}

# ---------------------------------------------------------------------------
Titulo "8/8  Registrando el servicio (tarea programada) y firewall"

# Servicio de arranque: genera el .cmd que la tarea ejecutara.
$servicioCmd = @"
@echo off
setlocal
cd /d "%~dp0"
set "KARDEX_HOST=0.0.0.0"
set "KARDEX_PORT=$Puerto"
set "KARDEX_ENV_PATH=$envDest"
if not exist "logs" mkdir "logs"
"$pyExe" $pyArgs kardex_server.py >> "logs\kardex_server.log" 2>&1
endlocal
"@
$servicioPath = Join-Path $Destino "servicio.cmd"
[System.IO.File]::WriteAllText($servicioPath, $servicioCmd, [System.Text.Encoding]::Default)
Ok "Script de arranque: $servicioPath"

# Registro de la tarea programada (arranca al iniciar el servidor, como SYSTEM).
$xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Kardex Fimas - servidor de reportes de inventario (Odoo).</Description>
  </RegistrationInfo>
  <Triggers>
    <BootTrigger><Enabled>true</Enabled></BootTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RestartOnFailure>
      <Interval>PT1M</Interval>
      <Count>999</Count>
    </RestartOnFailure>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>cmd.exe</Command>
      <Arguments>/c "$servicioPath"</Arguments>
      <WorkingDirectory>$Destino</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@
$xmlPath = Join-Path $env:TEMP "KardexERP-task.xml"
[System.IO.File]::WriteAllText($xmlPath, $xml, [System.Text.Encoding]::Unicode)
schtasks /Create /TN "$Tarea" /XML "$xmlPath" /F | Out-Null
if ($LASTEXITCODE -ne 0) { Fail "No se pudo registrar la tarea programada '$Tarea'." }
Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
Ok "Tarea programada '$Tarea' registrada (arranque automatico)."

# Firewall
if (-not $SinFirewall) {
    $regla = "Kardex Fimas $Puerto"
    if (Get-NetFirewallRule -DisplayName $regla -ErrorAction SilentlyContinue) {
        Ok "Regla de firewall ya existente ($regla)."
    } else {
        try {
            New-NetFirewallRule -DisplayName $regla -Direction Inbound -Action Allow `
                -Protocol TCP -LocalPort $Puerto -Profile Any -ErrorAction Stop | Out-Null
            Ok "Regla de firewall creada para el puerto $Puerto."
        } catch {
            netsh advfirewall firewall add rule name="$regla" dir=in action=allow protocol=TCP localport=$Puerto | Out-Null
            Ok "Regla de firewall creada (netsh) para el puerto $Puerto."
        }
    }
}

# ---------------------------------------------------------------------------
if (-not $SinIniciar) {
    Start-ScheduledTask -TaskName $Tarea -ErrorAction SilentlyContinue
    Info "Iniciando el servicio..."
    $activo = $false
    for ($i = 0; $i -lt 15; $i++) {
        Start-Sleep -Seconds 1
        if (PuertoActivo $Puerto) { $activo = $true; break }
    }
    if ($activo) { Ok "El servicio responde en el puerto $Puerto." }
    else { Warn "El servicio aun no responde. Revise $Destino\logs\kardex_server.log" }
}

Write-Host ""
Write-Host "  ===========================================================" -ForegroundColor Green
Write-Host "   INSTALACION COMPLETADA" -ForegroundColor Green
Write-Host "  ===========================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   Carpeta ......... $Destino"
Write-Host "   Credenciales .... $envDest"
Write-Host "   Puerto .......... $Puerto"
Write-Host "   Tarea ........... $Tarea  (arranque automatico al iniciar el servidor)"
Write-Host "   Log ............. $Destino\logs\kardex_server.log"
Write-Host ""
Write-Host "   Acceso local .... http://localhost:$Puerto/" -ForegroundColor Cyan
try {
    $ip = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.PrefixOrigin -ne "WellKnown" -and $_.IPAddress -notlike "169.254.*" } |
        Select-Object -First 1).IPAddress
    if ($ip) { Write-Host "   Acceso en red ... http://${ip}:$Puerto/" -ForegroundColor Cyan }
} catch {}
Write-Host ""
if (-not (Test-Path $envDest) -or (Select-String -Path $envDest -Pattern "suempresa" -Quiet)) {
    Write-Host "   RECUERDE: edite '$envDest' con las credenciales reales de Odoo." -ForegroundColor Yellow
}
Write-Host ""
