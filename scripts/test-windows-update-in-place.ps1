$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($env:CI -ne 'true' -or $env:RUNNER_OS -ne 'Windows') { throw 'Requires an isolated Windows CI runner.' }
$root = Split-Path $PSScriptRoot -Parent
$version = (Get-Content (Join-Path $root 'package.json') | ConvertFrom-Json).version
$setup = Join-Path $root "release/$version/win-x64/Sheet-Studio-$version-Windows-x64-Setup.exe"
$legacyRoot = Join-Path $env:RUNNER_TEMP 'sheet-update-baseline'
$legacySetup = Join-Path $legacyRoot 'release/0.7.36/win-x64/Sheet-Studio-0.7.36-Windows-x64-Setup.exe'
$evidence = Join-Path $root 'test-results/windows-update-in-place'
New-Item -ItemType Directory -Path $evidence -Force | Out-Null
$ps51 = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
$env:STUDIO_DISABLE_UPDATES = '1'
Add-Type -TypeDefinition (Get-Content (Join-Path $PSScriptRoot 'windows-file-support.cs') -Raw)
$report = [ordered]@{ passed=$false; scenarios=@() }
$work = Join-Path $env:RUNNER_TEMP "Sheet Studio update O'Brien; & 한글"
New-Item -ItemType Directory -Path $work -Force | Out-Null
$recovery = Join-Path $env:APPDATA 'Sheet Studio/recovery/in-place-preservation.txt'
$document = Join-Path $env:USERPROFILE 'Documents/sheet-update-preservation.txt'
New-Item -ItemType Directory -Path (Split-Path $recovery -Parent), (Split-Path $document -Parent) -Force | Out-Null
[IO.File]::WriteAllText($recovery,'PRESERVE RECOVERY')
[IO.File]::WriteAllText($document,'PRESERVE DOCUMENT')

function Run-Setup([string]$File, [string]$Scope, [string]$Directory) {
  $process = Start-Process -FilePath $File -ArgumentList @('/S',$Scope,('/D=' + $Directory)) -PassThru
  $null = $process.Handle
  if (-not $process.WaitForExit(180000)) { throw 'Seed Setup timed out.' }
  $process.Refresh()
  if ($process.ExitCode -ne 0) { throw "Seed Setup failed: $($process.ExitCode)" }
}
function Entries {
  foreach ($hive in @('HKCU','HKLM')) {
    foreach ($entry in @(Get-ItemProperty "$($hive):/Software/Microsoft/Windows/CurrentVersion/Uninstall/*" -ErrorAction SilentlyContinue)) {
      if ($entry.PSObject.Properties['DisplayName'] -and $entry.DisplayName -eq 'Sheet Studio') { $entry }
    }
  }
}
function Desktop-Links {
  foreach ($folder in @([Environment]::GetFolderPath('DesktopDirectory'),[Environment]::GetFolderPath('CommonDesktopDirectory'))) {
    $file = Join-Path $folder 'Sheet Studio.lnk'
    if (Test-Path $file) { $file }
  }
}
function Assert-Preserved {
  if ([IO.File]::ReadAllText($recovery) -ne 'PRESERVE RECOVERY' -or [IO.File]::ReadAllText($document) -ne 'PRESERVE DOCUMENT') { throw 'Document or recovery changed.' }
}
function Stop-TestApp {
  # These are disposable CI installations, never a user's desktop.
  Get-Process -Name 'Sheet Studio' -ErrorAction SilentlyContinue | Stop-Process -Force
}
function Remove-TestInstall([string]$Directory,[string]$Scope) {
  Stop-TestApp
  $uninstall = Join-Path $Directory 'Uninstall Sheet Studio.exe'
  $process = Start-Process -FilePath $uninstall -ArgumentList @('/S',$Scope) -PassThru
  if (-not $process.WaitForExit(120000)) { throw 'CI uninstall timed out.' }
  $deadline = [DateTime]::UtcNow.AddSeconds(30)
  while (Test-Path (Join-Path $Directory 'Sheet Studio.exe')) {
    if ([DateTime]::UtcNow -gt $deadline) { throw 'CI uninstall left its executable.' }
    Start-Sleep -Milliseconds 250
  }
}
if (@(Entries).Count) { throw 'An earlier test left a Studio installation on this runner.' }
try {
  & (Join-Path $PSScriptRoot 'test-windows-installer-ui.ps1') -Setup $legacySetup -Stage old-forced-language -ExpectLanguageSelector
  foreach ($scenario in @('current-user','all-users','legacy-helper-with-duplicates')) {
    $scope = if ($scenario -eq 'current-user') { '/currentuser' } else { '/allusers' }
    $directory = Join-Path $work "$scenario/사용자 지정 도구"
    Run-Setup $legacySetup $scope $directory
    $target = Join-Path $directory 'Sheet Studio.exe'
    if (-not (Test-Path $target)) { throw 'Legacy setup ignored the custom directory.' }
    if ($scenario -eq 'legacy-helper-with-duplicates') {
      # Reproduce the old helper's forced /currentuser against an all-user install.
      Run-Setup $legacySetup '/currentuser' (Join-Path $env:LOCALAPPDATA 'Programs/Sheet Studio')
      if (@(Entries).Count -ne 2 -or @(Desktop-Links).Count -ne 2) { throw 'The old installation-scope defect was not reproduced.' }
    }
    $jobDirectory = Join-Path $work "$scenario/job"
    New-Item -ItemType Directory -Path $jobDirectory -Force | Out-Null
    $installer = Join-Path $jobDirectory 'update.exe'
    Copy-Item -LiteralPath $setup -Destination $installer
    $log = Join-Path $jobDirectory 'install.log'
    # A short-lived disposable parent exercises the real PowerShell 5.1 wait.
    $sleep = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes('Start-Sleep -Seconds 4'))
    $parent = Start-Process -FilePath $ps51 -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$sleep) -WindowStyle Hidden -PassThru
    $jobFile = Join-Path $jobDirectory 'install.json'
    @{ schema=1; id='sheet'; parentPid=$parent.Id; installer=$installer; bytes=(Get-Item $installer).Length;
       sha256=(Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant(); executable=$target; version=$version; log=$log } |
      ConvertTo-Json | Set-Content -LiteralPath $jobFile -Encoding UTF8
    $helper = if ($scenario -eq 'legacy-helper-with-duplicates') { Join-Path $legacyRoot 'electron/studio-update-windows.ps1' } else { Join-Path $root 'electron/studio-update-windows.ps1' }
    # Match the product loader's explicit UTF-8 decoding in PowerShell 5.1.
    $workerScript = Join-Path $jobDirectory 'helper.ps1'
    [IO.File]::WriteAllText($workerScript, (Get-Content -LiteralPath $helper -Raw -Encoding UTF8), [Text.UTF8Encoding]::new($true))
    $command = "& ([ScriptBlock]::Create([IO.File]::ReadAllText('" + $workerScript.Replace("'","''") + "', [Text.Encoding]::UTF8))) -JobFile '" + $jobFile.Replace("'","''") + "'"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $worker = Start-Process -FilePath $ps51 -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $jobDirectory 'stdout.log') -RedirectStandardError (Join-Path $jobDirectory 'stderr.log') -PassThru
    $null = $worker.Handle
    $worker.Refresh()
    if ($worker.HasExited) { throw ("Helper exited before Setup: " + (Get-Content (Join-Path $jobDirectory 'stderr.log') -Raw)) }
    & (Join-Path $PSScriptRoot 'test-windows-installer-ui.ps1') -Setup $setup -Stage $scenario -ObserveProcessId $worker.Id
    if (-not $worker.WaitForExit(120000) -or $worker.ExitCode -ne 0) { throw "Actual helper failed: $(Get-Content $log -Raw)" }
    $text = Get-Content $log -Raw
    if ($text -match 'Update failed:' -or $text -notmatch "Installed sheet $([regex]::Escape($version))") { throw "Unexpected update result: $text" }
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
      $main = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'Sheet Studio.exe' -and $_.ExecutablePath -ieq $target -and $_.CommandLine -notmatch '--type=' })
      if ($main.Count -eq 1) { break }
      Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($main.Count -ne 1) { throw "Expected one relaunched editor, found $($main.Count)." }
    if ((Get-Item -LiteralPath $target).VersionInfo.ProductVersion -ne $version) { throw 'The original installation was not updated.' }
    $entries = @(Entries)
    if ($entries.Count -ne 1 -or $entries[0].UninstallString -ine ('"' + (Join-Path $directory 'Uninstall Sheet Studio.exe') + '" ' + $scope)) { throw 'Update created another installation or moved the existing one.' }
    $links = @(Desktop-Links)
    if ($links.Count -ne 1) { throw "Expected one desktop shortcut, found $($links.Count)." }
    if ([SheetFileSupport]::ShortcutTarget($links[0]) -ine $target) { throw 'Desktop shortcut points to the wrong editor.' }
    Assert-Preserved
    $report.scenarios += [ordered]@{ scenario=$scenario; passed=$true; originalDirectoryPreserved=$true; originalScopePreserved=$true;
      installedVersion=$version; desktopShortcuts=$links.Count; mainProcesses=$main.Count; noFailurePopup=$true; documentAndRecoveryPreserved=$true }
    Copy-Item -LiteralPath $log -Destination (Join-Path $evidence "$scenario.log")
    Remove-TestInstall $directory $scope
    Assert-Preserved
    if (@(Entries).Count) { throw 'Uninstall left an extra installed copy.' }
  }
  $report.passed = $true
} finally {
  foreach ($file in @(Get-ChildItem -LiteralPath $work -Filter '*.log' -Recurse -File)) {
    $name = $file.FullName.Substring($work.Length + 1).Replace('\','-')
    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $evidence $name)
  }
  Stop-TestApp
  $report | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $evidence 'results.json')
  $report | ConvertTo-Json -Depth 8 | Write-Output
}
