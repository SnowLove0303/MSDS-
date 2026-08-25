param(
  [int]$Port = 8765
)
$webRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -LiteralPath $webRoot
Write-Host "MSDS Web formal: http://0.0.0.0:$Port" -ForegroundColor Cyan
python .\server.py --host 0.0.0.0 --port $Port
