$ErrorActionPreference = 'Stop'
$version = (Get-Content package.json | ConvertFrom-Json).version
$setup = Get-Item "release/$version/win-x64/*-Setup.exe"
$process = Start-Process -FilePath $setup.FullName -ArgumentList '/S' -Wait -PassThru
if ($process.ExitCode -ne 0) { throw "Installer failed: $($process.ExitCode)" }
$link = Join-Path ([Environment]::GetFolderPath('Programs')) 'Sheet Studio.lnk'
if (!(Test-Path $link)) { throw 'Start menu shortcut missing' }
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($link)
$target = $shortcut.TargetPath
if (!(Test-Path $target) -or $target.StartsWith($env:TEMP, [StringComparison]::OrdinalIgnoreCase)) { throw 'Shortcut must target a persistent installed executable' }
$explorer = New-Object -ComObject Shell.Application
$item = $explorer.NameSpace((Split-Path $link)).ParseName((Split-Path $link -Leaf))
if ($item.ExtendedProperty('System.AppUserModel.ID') -ne 'local.lattice.sheetstudio') { throw 'Shortcut AppUserModelID mismatch' }
Get-Process 'Sheet Studio' -ErrorAction SilentlyContinue | Stop-Process
$env:SHEET_STUDIO_EXECUTABLE = $target
node scripts/electron-smoke.cjs --packaged
if ($LASTEXITCODE -ne 0) { throw 'Installed application smoke test failed' }
Write-Output "Installed app and shortcut verified: $target"
