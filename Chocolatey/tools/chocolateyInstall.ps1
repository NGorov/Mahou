$ErrorActionPreference = 'Stop'

$packageArgs = @{
  packageName    = 'mahou3'
  url            = 'https://github.com/NGorov/Mahou/releases/download/latest-commit/Release_x86_x64.zip'
  unzipLocation  = "$(Split-Path -parent $MyInvocation.MyCommand.Definition)"
  checksum       = '4BA9D2579C1A34C67C65862C52419594F4E9412D65996911D7260606CED0DB1C'
  checksumType   = 'sha256'
}

Install-ChocolateyZipPackage @packageArgs
