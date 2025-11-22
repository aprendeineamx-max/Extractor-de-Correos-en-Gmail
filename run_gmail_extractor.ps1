<#
.SYNOPSIS
  Lanza el webhook de tu Apps Script Gmail→Sheets desde PowerShell.
.DESCRIPTION
  Usa el token y la URL publicados de tu implementación web. Acepta acciones:
  fullrescan, backfill, incremental, ping. Si quieres sobreescribir valores
  sin editar el archivo, define las variables de entorno:
    $env:GMAIL_EXTRACTOR_URL
    $env:RUN_TOKEN
  Ejemplos:
    .\run_gmail_extractor.ps1 -Action fullrescan
    .\run_gmail_extractor.ps1 -Action backfill
#>

[CmdletBinding()]
param(
  [ValidateSet('fullrescan','backfill','incremental','ping')]
  [string]$Action = 'fullrescan',

  # Permite reemplazar la URL vía CLI o $env:GMAIL_EXTRACTOR_URL
  [string]$BaseUrl = $env:GMAIL_EXTRACTOR_URL,

  # Permite reemplazar el token vía CLI o $env:RUN_TOKEN
  [string]$Token = $env:RUN_TOKEN
)

$ProgressPreference = 'SilentlyContinue'

if (-not $BaseUrl) {
  $BaseUrl = 'https://script.google.com/macros/s/AKfycbycJbwT6kSasjRSvy9T39lxEKtdNSQVc8o9gmTTDDYSbxl0ugxYURM4Lyrn0aqE7DVJDQ/exec'
}
if (-not $Token) {
  $Token = '4b9f1e5e6f4f46b9a730b4f88a3d9e25c2a1c0a5afd349e0b6c7f9123a5c0d21'
}

Write-Host "Lanzando acción '$Action' contra $BaseUrl" -ForegroundColor Cyan

$body = @{
  action = $Action
  token  = $Token  # el endpoint acepta header y/o parámetro
}
$headers = @{
  'X-Run-Token' = $Token
}

try {
  $resp = Invoke-RestMethod -Uri $BaseUrl -Method Post -Headers $headers `
    -Body $body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 60 `
    -ErrorAction Stop

  Write-Host "Respuesta:" -ForegroundColor Green
  $resp | ConvertTo-Json -Depth 4
} catch {
  Write-Error "Falló la llamada: $($_.Exception.Message)"
  if ($_.ErrorDetails) { Write-Warning $_.ErrorDetails }
  exit 1
}
