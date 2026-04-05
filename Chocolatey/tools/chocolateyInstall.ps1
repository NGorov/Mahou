$ErrorActionPreference = 'Stop'

$toolsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$packageArgs = @{
  packageName    = 'mahou3'
  url            = 'https://github.com/NGorov/Mahou3/releases/download/v3.0/Release_x86_x64.zip'
  unzipLocation  = $toolsDir
  checksum       = 'C1E108FFC6CD50A9ADDDE14D85FB7F1B5F51C6338C24E2D96F9E601A1F87E118'
  checksumType   = 'sha256'
}

Install-ChocolateyZipPackage @packageArgs

$exePath = Join-Path $toolsDir 'net10.0-windows\Mahou.exe'
Install-BinFile -Name 'mahou3' -Path $exePath
