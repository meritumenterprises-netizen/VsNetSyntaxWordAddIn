param(
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Debug',
    [switch]$Register
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$instances = @(& $vswhere -all -products '*' -format json | ConvertFrom-Json)
$instance = $instances | Where-Object {
    $major = ([version]$_.installationVersion).Major
    (Test-Path -LiteralPath (Join-Path $_.installationPath 'MSBuild\Current\Bin\MSBuild.exe')) -and
    (Test-Path -LiteralPath (Join-Path $_.installationPath "MSBuild\Microsoft\VisualStudio\v$major.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"))
} | Sort-Object { [version]$_.installationVersion } -Descending | Select-Object -First 1
if (-not $instance) { throw 'Install Visual Studio with the Office/SharePoint development workload.' }
$installation = $instance.installationPath
$visualStudioVersion = "$(([version]$instance.installationVersion).Major).0"
& (Join-Path $PSScriptRoot 'Initialize-DevSigning.ps1')
$msbuild = Join-Path $installation 'MSBuild\Current\Bin\MSBuild.exe'
$ide = (Join-Path $installation 'Common7\IDE') + '\'
$arguments = @(
    (Join-Path $root 'VsNetSyntaxWordAddIn.csproj'), '/restore',
    "/p:Configuration=$Configuration", "/p:DevEnvDir=$ide",
    "/p:VisualStudioVersion=$visualStudioVersion",
    "/p:RestoreConfigFile=$(Join-Path $root 'NuGet.Config')", '/v:minimal'
)
if (-not $Register) { $arguments += '/p:PrepareForRunDependsOn=VisualStudioForApplicationsBuild' }
& $msbuild @arguments
if ($LASTEXITCODE -ne 0) { throw 'VSTO build failed.' }
