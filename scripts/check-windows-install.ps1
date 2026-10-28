$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ($env:CI -ne 'true' -or $env:RUNNER_OS -ne 'Windows') {
  throw 'Run this install/reinstall/uninstall test only on an isolated Windows CI runner.'
}

$root = Split-Path $PSScriptRoot -Parent
$version = (Get-Content (Join-Path $root 'package.json') | ConvertFrom-Json).version
$setup = Join-Path $root "release/$version/win-x64/Sheet-Studio-$version-Windows-x64-Setup.exe"
$installDir = Join-Path $env:LOCALAPPDATA 'Programs\한글 설치 경로\Sheet Studio'
$executable = Join-Path $installDir 'Sheet Studio.exe'
$profile = Join-Path $env:APPDATA 'Sheet Studio'
$recovery = Join-Path $profile 'recovery/installer-preservation-test.sheet.json'
$document = Join-Path $env:USERPROFILE 'Documents/installer-preservation-test.sheet.json'
$desktopLink = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Sheet Studio.lnk'
$startLink = Join-Path ([Environment]::GetFolderPath('StartMenu')) 'Programs/Sheet Studio.lnk'
$reportPath = Join-Path $root 'test-results/windows-installer.json'
$report = [ordered]@{ version = $version; passed = $false }
$extensions = @('xlsx', 'xlsm', 'csv', 'tsv')
$defaults = @{}
$choices = @{}
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'windows-file-support.cs') -Raw)

function Get-UserChoice([string] $extension) {
  $key = "HKCU:/Software/Microsoft/Windows/CurrentVersion/Explorer/FileExts/.$extension/UserChoice"
  if (-not (Test-Path $key)) { return '' }
  $value = Get-Item $key
  return (@($value.GetValueNames() | Sort-Object | ForEach-Object { @($_, $value.GetValue($_)) }) | ConvertTo-Json -Compress)
}

function Assert-AssociationsPreserved {
  foreach ($extension in $extensions) {
    $key = Get-Item "HKCU:/Software/Classes/.$extension"
    if ($key.GetValue('') -ne $defaults[$extension]) { throw "Changed existing .$extension default" }
    if ((Get-UserChoice $extension) -ne $choices[$extension]) { throw "Changed .$extension UserChoice" }
    $other = Get-Item "HKCU:/Software/Classes/.$extension/OpenWithProgids"
    if ($other.GetValueNames() -notcontains 'SheetStudio.QA.OtherEditor') { throw "Removed another .$extension supporting app" }
  }
}

function Assert-FileSupport {
  $registered = Get-Item 'HKCU:/Software/RegisteredApplications'
  if ($registered.GetValue('Sheet Studio') -ne 'Software\Sheet Studio\Capabilities') { throw 'Missing Default apps registration' }
  $supported = Get-Item 'HKCU:/Software/Classes/Applications/Sheet Studio.exe/SupportedTypes'
  $capabilities = Get-Item 'HKCU:/Software/Sheet Studio/Capabilities/FileAssociations'
  $command = '"' + $executable + '" "%1"'
  $appCommand = Get-Item 'HKCU:/Software/Classes/Applications/Sheet Studio.exe/shell/open/command'
  if ($appCommand.GetValue('') -ne $command) { throw 'Application command must quote executable and filename' }
  foreach ($extension in $extensions) {
    $progId = "SheetStudio.$extension"
    if ($supported.GetValueNames() -notcontains ".$extension") { throw "Missing .$extension SupportedTypes" }
    if ($capabilities.GetValue(".$extension") -ne $progId) { throw "Missing .$extension capability" }
    $openWith = Get-Item "HKCU:/Software/Classes/.$extension/OpenWithProgids"
    if ($openWith.GetValueNames() -notcontains $progId) { throw "Missing .$extension OpenWithProgids" }
    $open = Get-Item "HKCU:/Software/Classes/$progId/shell/open/command"
    if ($open.GetValue('') -ne $command) { throw "Incorrect .$extension open command" }
    $handlers = [SheetFileSupport]::Applications(".$extension")
    if ($handlers -notcontains 'Sheet Studio') { throw "Explorer does not offer Sheet Studio for .$extension : $handlers" }
  }
  $legacy = Get-Item 'HKCU:/Software/Classes/Excel Workbook/shell/open/command'
  if ($legacy.GetValue('') -ne $command) { throw 'Previous Sheet Studio default ProgID no longer opens this app' }
  $otherLegacy = Get-Item 'HKCU:/Software/Classes/CSV Spreadsheet/shell/open/command'
  if ($otherLegacy.GetValue('') -ne 'OTHER EDITOR COMMAND') { throw 'Overwrote another app legacy handler' }
  if ($supported.GetValueNames() -contains '.json' -or $supported.GetValueNames() -contains '.xls') { throw 'Registered an unsupported broad file type' }
  Assert-AssociationsPreserved
}

function Assert-FileSupportRemoved {
  foreach ($key in @('HKCU:/Software/Classes/Applications/Sheet Studio.exe', 'HKCU:/Software/Sheet Studio/Capabilities')) {
    if (Test-Path $key) { throw "Left registration after uninstall: $key" }
  }
  $registered = Get-Item 'HKCU:/Software/RegisteredApplications'
  if ($registered.GetValueNames() -contains 'Sheet Studio') { throw 'Left Default apps registration' }
  foreach ($extension in $extensions) {
    $progId = "SheetStudio.$extension"
    if (Test-Path "HKCU:/Software/Classes/$progId") { throw "Left .$extension ProgID" }
    $openWith = Get-Item "HKCU:/Software/Classes/.$extension/OpenWithProgids"
    if ($openWith.GetValueNames() -contains $progId) { throw "Left .$extension Open With entry" }
  }
  Assert-AssociationsPreserved
  if (Test-Path 'HKCU:/Software/Classes/Excel Workbook') { throw 'Left app-owned legacy alias' }
  $otherLegacy = Get-Item 'HKCU:/Software/Classes/CSV Spreadsheet/shell/open/command'
  if ($otherLegacy.GetValue('') -ne 'OTHER EDITOR COMMAND') { throw 'Removed another app legacy handler' }
}

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
  $unicodeTarget = [SheetFileSupport]::ShortcutTarget($startLink)
  $report.shortcutTargetWSH = $shortcut.TargetPath
  $report.shortcutTargetUnicode = $unicodeTarget
  New-Item -ItemType Directory -Path (Split-Path $reportPath -Parent) -Force | Out-Null
  Copy-Item $startLink (Join-Path (Split-Path $reportPath -Parent) 'installed-start-menu.lnk') -Force
  if ([IO.Path]::GetFullPath($unicodeTarget) -ne [IO.Path]::GetFullPath($executable)) {
    throw "Unicode shell shortcut target '$unicodeTarget' differs from '$executable'."
  }
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
  # Seed independent defaults and another supporting app on this CI runner.
  # TSV exercises the case where there is no extension default to begin with.
  foreach ($extension in $extensions) {
    $key = New-Item "HKCU:/Software/Classes/.$extension" -Force
    if ($extension -ne 'tsv') { $key.SetValue('', "SheetStudio.QA.Previous.$extension") }
    $defaults[$extension] = $key.GetValue('')
    $choices[$extension] = Get-UserChoice $extension
    New-Item "HKCU:/Software/Classes/.$extension/OpenWithProgids" -Force | Out-Null
    New-ItemProperty "HKCU:/Software/Classes/.$extension/OpenWithProgids" -Name 'SheetStudio.QA.OtherEditor' -Value '' -PropertyType String | Out-Null
  }
  # Simulate an older Sheet Studio choice whose ProgID was removed by its
  # uninstaller, plus a generic legacy name now owned by a different editor.
  (Get-Item 'HKCU:/Software/Classes/.xlsx').SetValue('', 'Excel Workbook')
  $defaults['xlsx'] = 'Excel Workbook'
  New-Item 'HKCU:/Software/Classes/CSV Spreadsheet/shell/open/command' -Force | Out-Null
  (Get-Item 'HKCU:/Software/Classes/CSV Spreadsheet/shell/open/command').SetValue('', 'OTHER EDITOR COMMAND')
  New-Item -ItemType Directory -Path (Split-Path $recovery -Parent), (Split-Path $document -Parent) -Force | Out-Null
  Set-Content $recovery 'PRESERVE SHEET STUDIO USER DATA'
  Set-Content $document 'PRESERVE SHEET STUDIO USER DATA'
  Run-Installer $setup @('/S', '/currentuser', "/D=$installDir")
  Assert-Installed
  Assert-Preserved
  Assert-FileSupport
  $report.openWith = $extensions
  $report.defaultAppsRegistration = $true
  $report.previousDefaultsAndOtherAppsPreserved = $true
  $report.quotedUnicodeInstallPath = $true
  $report.previousSheetStudioHandlerPreserved = $true
  $report.freshInstall = $true
  $report.shortcuts = $true
  $report.installedAppsEntry = $true

  $env:SHEET_STUDIO_EXECUTABLE = $executable
  & node (Join-Path $root 'scripts/electron-smoke.cjs') --packaged
  if ($LASTEXITCODE -ne 0) { throw 'Installed Windows application smoke test failed.' }
  & node (Join-Path $root 'tests/ui-file-associations.mjs') --packaged
  if ($LASTEXITCODE -ne 0) { throw 'Installed Windows file-open regression failed.' }
  $report.installedApplication = $true

  Run-Installer $setup @('/S', '/currentuser', "/D=$installDir")
  Assert-Installed
  Assert-Preserved
  Assert-FileSupport
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
  Assert-FileSupportRemoved
  $report.fileSupportRemovedOnUninstall = $true
  $report.uninstall = $true
  $report.userDataPreserved = $true
  $report.passed = $true
} finally {
  New-Item -ItemType Directory -Path (Split-Path $reportPath -Parent) -Force | Out-Null
  $report | ConvertTo-Json | Set-Content $reportPath
  $report | ConvertTo-Json | Write-Output
}
