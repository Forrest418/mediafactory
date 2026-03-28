param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repoRoot "MediaFactory.App\MediaFactory.App.vbproj"
$publishRoot = Join-Path $repoRoot "artifacts\publish\$Runtime"
$packageRoot = Join-Path $repoRoot "artifacts\release\MediaFactory-$Runtime"
$zipPath = Join-Path $repoRoot "artifacts\release\MediaFactory-$Runtime.zip"
$standaloneExePath = Join-Path $repoRoot "artifacts\release\MediaFactory-$Runtime.exe"
$templateConfigPath = Join-Path $repoRoot "ModelProviders.json"

Write-Host "Publishing MediaFactory for $Runtime..."

dotnet publish $projectPath `
    -c $Configuration `
    -r $Runtime `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -o $publishRoot

if (Test-Path $packageRoot) {
    Remove-Item $packageRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
Copy-Item (Join-Path $publishRoot "*") $packageRoot -Recurse -Force
Copy-Item $templateConfigPath (Join-Path $packageRoot "ModelProviders.json") -Force

$exe = Get-ChildItem $packageRoot -Filter "*.exe" | Select-Object -First 1
if ($null -eq $exe) {
    throw "No executable was produced in $packageRoot."
}

if ($exe.Name -ne "MediaFactory.exe") {
    Rename-Item $exe.FullName "MediaFactory.exe"
}

if (Test-Path $standaloneExePath) {
    Remove-Item $standaloneExePath -Force
}

Copy-Item (Join-Path $packageRoot "MediaFactory.exe") $standaloneExePath -Force

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath -Force

Write-Host "Standalone executable: $standaloneExePath"
Write-Host "Zip package: $zipPath"
