$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ($env:CI -ne 'true' -or $env:RUNNER_OS -ne 'Windows') {
  throw 'Run this install/reinstall/uninstall test only on an isolated Windows CI runner.'
}

$root = Split-Path $PSScriptRoot -Parent
$version = (Get-Content (Join-Path $root 'package.json') | ConvertFrom-Json).version
$setup = Join-Path $root "release/$version/win-x64/Sheet-Studio-$version-Windows-x64-Setup.exe"
$installDir = Join-Path $env:LOCALAPPDATA 'Programs/Sheet Studio'
$executable = Join-Path $installDir 'Sheet Studio.exe'
$profile = Join-Path $env:APPDATA 'Sheet Studio'
$recovery = Join-Path $profile 'recovery/installer-preservation-test.sheet.json'
$document = Join-Path $env:USERPROFILE 'Documents/installer-preservation-test.sheet.json'
$desktopLink = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Sheet Studio.lnk'
$startLink = Join-Path ([Environment]::GetFolderPath('StartMenu')) 'Programs/Sheet Studio.lnk'
$reportPath = Join-Path $root 'test-results/windows-installer.json'
$report = [ordered]@{ version = $version; passed = $false }

function Run-Installer([string] $file, [string[]] $arguments) {
  $process = Start-Process -FilePath $file -ArgumentList $arguments -PassThru
  if (-not $process.WaitForExit(180000)) {
    $process.Kill($true)
    throw "Installer timed out: $file"
  }
  if ($process.ExitCode -ne 0) { throw "Installer exit code: $($process.ExitCode)" }
}

function Assert-Installed {
  foreach ($file in @($executable, (Join-Path $installDir 'resources/app.asar'), $desktopLink, $startLink)) {
    if (-not (Test-Path $file)) { throw "Missing installed file: $file" }
  }
  $actual = (Get-Item $executable).VersionInfo.ProductVersion
  if ($actual -ne $version) { throw "Installed version $actual does not match $version" }
  $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($startLink)
  if ($shortcut.TargetPath -ne $executable) { throw 'Start menu shortcut points to another executable.' }
  $explorer = New-Object -ComObject Shell.Application
  $item = $explorer.NameSpace((Split-Path $startLink)).ParseName((Split-Path $startLink -Leaf))
  if ($item.ExtendedProperty('System.AppUserModel.ID') -ne 'local.lattice.sheetstudio') { throw 'Shortcut AppUserModelID mismatch' }
  $packed = Join-Path $root "release/$version/win-x64/win-unpacked/resources/app.asar"
  if ((Get-FileHash $packed -Algorithm SHA256).Hash -ne (Get-FileHash (Join-Path $installDir 'resources/app.asar') -Algorithm SHA256).Hash) { throw 'Installed app archive differs from tested package' }
  $entry = @(Get-ItemProperty 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/*' |
    Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -eq 'Sheet Studio' })
  if ($entry.Count -ne 1 -or $entry[0].DisplayVersion -ne $version) {
    throw 'Expected one current-user installed-app entry with the release version.'
  }
}

function Assert-Preserved {
  foreach ($file in @($recovery, $document)) {
    if (-not (Test-Path $file) -or (Get-Content $file -Raw).Trim() -ne 'PRESERVE SHEET STUDIO USER DATA') {
      throw "User data was removed or changed: $file"
    }
  }
}

if (Test-Path $installDir) { throw 'This runner already has Sheet Studio installed.' }
try {
  New-Item -ItemType Directory -Path (Split-Path $recovery -Parent), (Split-Path $document -Parent) -Force | Out-Null
  Set-Content $recovery 'PRESERVE SHEET STUDIO USER DATA'
  Set-Content $document 'PRESERVE SHEET STUDIO USER DATA'
  Run-Installer $setup @('/S', '/currentuser')
  Assert-Installed
  Assert-Preserved
  $report.freshInstall = $true
  $report.shortcuts = $true
  $report.installedAppsEntry = $true

  $env:SHEET_STUDIO_EXECUTABLE = $executable
  & node (Join-Path $root 'scripts/electron-smoke.cjs') --packaged
  if ($LASTEXITCODE -ne 0) { throw 'Installed Windows application smoke test failed.' }
  $report.installedApplication = $true

  Run-Installer $setup @('/S', '/currentuser')
  Assert-Installed
  Assert-Preserved
  $report.reinstall = $true

  $uninstaller = @(Get-ChildItem $installDir -Filter '*Uninstall*.exe')
  if ($uninstaller.Count -ne 1) { throw 'Expected exactly one uninstaller.' }
  Run-Installer $uninstaller[0].FullName @('/S', '/currentuser')
  # NSIS runs its uninstaller from a temporary child process. The original
  # process can exit before shortcut and registry cleanup has finished.
  $deadline = (Get-Date).AddSeconds(30)
  do {
    $leftoverFiles = @(@($executable, $desktopLink, $startLink) | Where-Object { Test-Path $_ })
    $remaining = @(Get-ItemProperty 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/*' |
      Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -eq 'Sheet Studio' })
    if ($leftoverFiles.Count -eq 0 -and $remaining.Count -eq 0) { break }
    Start-Sleep -Milliseconds 250
  } while ((Get-Date) -lt $deadline)
  foreach ($file in @($executable, $desktopLink, $startLink)) {
    if (Test-Path $file) { throw "Uninstall left an application file or shortcut: $file" }
  }
  $remaining = @(Get-ItemProperty 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/*' |
    Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -eq 'Sheet Studio' })
  if ($remaining.Count -ne 0) { throw 'Uninstall left the installed-app registration.' }
  Assert-Preserved
  $report.uninstall = $true
  $report.userDataPreserved = $true
  $report.passed = $true
} finally {
  New-Item -ItemType Directory -Path (Split-Path $reportPath -Parent) -Force | Out-Null
  $report | ConvertTo-Json | Set-Content $reportPath
  $report | ConvertTo-Json | Write-Output
}
