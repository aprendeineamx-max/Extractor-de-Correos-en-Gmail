[CmdletBinding()]
param(
  [string]$BaseUrl = $env:GMAIL_EXTRACTOR_URL,
  [string]$Token   = $env:RUN_TOKEN
)

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$here\run_gmail_extractor.ps1" -Action ping -BaseUrl $BaseUrl -Token $Token
