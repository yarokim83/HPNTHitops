$ErrorActionPreference = 'Stop'
$sourceExe = Join-Path $PSScriptRoot 'dist\release\PRMakerWidget.exe'
$installDir = Join-Path $env:LOCALAPPDATA 'PRMaker\app'
$targetExe = Join-Path $installDir 'PRMakerWidget.exe'
if (-not (Test-Path -LiteralPath $sourceExe -PathType Leaf)) { throw 'Build the release first.' }
$check = Start-Process -FilePath $sourceExe -ArgumentList '--self-check' -WindowStyle Hidden -PassThru
if (-not $check.WaitForExit(60000)) { $check.Kill(); throw 'Release smoke check timed out.' }
if ($check.ExitCode -ne 0) { throw 'Release smoke check failed; previous installation preserved.' }
$running = Get-CimInstance Win32_Process -Filter "Name='PRMakerWidget.exe'" |
    Where-Object { $_.ExecutablePath -eq $targetExe }
if ($running) { throw 'Exit the installed PR Maker from its tray menu before replacing it.' }
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
$stagedExe = Join-Path $installDir 'PRMakerWidget.new.exe'
Copy-Item -LiteralPath $sourceExe -Destination $stagedExe -Force
if ((Get-FileHash -LiteralPath $stagedExe).Hash -ne (Get-FileHash -LiteralPath $sourceExe).Hash) {
    throw 'Release copy verification failed.'
}
if (Test-Path -LiteralPath $targetExe) {
    $backupExe = Join-Path $installDir ('PRMakerWidget-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.bak')
    [IO.File]::Replace($stagedExe, $targetExe, $backupExe)
} else {
    Move-Item -LiteralPath $stagedExe -Destination $targetExe
}
$startupLink = Join-Path ([Environment]::GetFolderPath('Startup')) 'PRMakerWidget.lnk'
$shellObj = New-Object -ComObject WScript.Shell
if (Test-Path -LiteralPath $startupLink) {
    Copy-Item -LiteralPath $startupLink -Destination (Join-Path $installDir 'startup-previous.lnk') -Force
    $shortcut = $shellObj.CreateShortcut($startupLink)
    $shortcut.TargetPath = $targetExe
    $shortcut.WorkingDirectory = $installDir
    $shortcut.Arguments = ''
    $shortcut.IconLocation = $targetExe + ',0'
    $shortcut.Save()
    if ($shellObj.CreateShortcut($startupLink).TargetPath -ne $targetExe) { throw 'Startup verification failed.' }
}
Write-Output "Installed and verified: $targetExe"
