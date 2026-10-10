param([ValidateSet('Debug', 'Release')][string]$Configuration = 'Debug')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$installation = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.OfficeTools -property installationPath
if (-not $installation) { throw 'Install Visual Studio with the Office/SharePoint development workload.' }
& (Join-Path $PSScriptRoot 'Initialize-DevSigning.ps1')
$msbuild = Join-Path $installation 'MSBuild\Current\Bin\MSBuild.exe'
$ide = (Join-Path $installation 'Common7\IDE') + '\'
& $msbuild (Join-Path $root 'VsNetSyntaxWordAddIn.csproj') /restore "/p:Configuration=$Configuration" "/p:DevEnvDir=$ide" "/p:RestoreConfigFile=$(Join-Path $root 'NuGet.Config')" /p:PrepareForRunDependsOn=VisualStudioForApplicationsBuild /v:minimal
if ($LASTEXITCODE -ne 0) { throw 'VSTO build failed.' }
