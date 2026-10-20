<#
  ============================================================================
   KARDEX FIMAS - Actualizacion (Windows Server)
  ============================================================================
   Actualiza una instalacion existente SIN perder datos: detiene el servicio,
   copia los archivos nuevos (nunca sobrescribe bases de datos ni el .env) y
   vuelve a iniciar el servicio.

   Uso (PowerShell como Administrador):
     .\actualizar.ps1
     .\actualizar.ps1 -Destino "D:\Apps\Kardex" -Puerto 8765
  ============================================================================
#>
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string] $Destino = "C:\Kardex-ERP",
    [int]    $Puerto  = 8765,
    [string] $EnvPath = "",
    [string] $Tarea   = "KardexERP"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "  ===========================================================" -ForegroundColor Cyan
Write-Host "        KARDEX FIMAS - Actualizacion" -ForegroundColor Cyan
Write-Host "  ===========================================================" -ForegroundColor Cyan

$parametros = @{ Destino = $Destino; Puerto = $Puerto; Tarea = $Tarea }
if ($EnvPath) { $parametros.EnvPath = $EnvPath }

& (Join-Path $PSScriptRoot "install.ps1") @parametros
