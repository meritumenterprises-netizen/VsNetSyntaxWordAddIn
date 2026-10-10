param([ValidateSet('Debug', 'Release')][string]$Configuration = 'Debug')
$ErrorActionPreference = 'Stop'

$wordPath = $null
foreach ($registryPath in @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
)) {
    if (Test-Path -LiteralPath $registryPath) {
        $candidate = (Get-Item -LiteralPath $registryPath).GetValue('')
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            $wordPath = $candidate
            break
        }
    }
}
if (-not $wordPath) { throw 'Desktop Microsoft Word is not installed or its executable is not registered.' }

& (Join-Path $PSScriptRoot 'Build.ps1') -Configuration $Configuration -Register
Start-Process -FilePath $wordPath -ArgumentList '/x'
