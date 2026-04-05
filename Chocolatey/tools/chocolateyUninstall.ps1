$ErrorActionPreference = 'Stop'

$toolsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$exePath = Join-Path $toolsDir 'net10.0-windows\Mahou.exe'

if (Test-Path $exePath) {
  Uninstall-BinFile -Name 'mahou3' -Path $exePath
}

$installDir = Join-Path $toolsDir 'net10.0-windows'
if (Test-Path $installDir) {
  Remove-Item $installDir -Recurse -Force
}
