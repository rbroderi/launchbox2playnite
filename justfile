set shell := ["powershell.exe", "-NoLogo", "-NoProfile", "-Command"]

# Show available commands when running `just` with no arguments
default:
    just --list

# Bundle the compiled plugin files into a Playnite extension package (.pext)
bundle output="LaunchBoxYamlImporter.pext":
    dotnet restore
    dotnet build -c Release
    @$ErrorActionPreference = "Stop"; \
    $projectRoot = Get-Location; \
    $files = @( \
        (Join-Path $projectRoot "LaunchBoxYamlImporter/bin/Release/net48/LaunchBoxYamlImporter.dll"), \
        (Join-Path $projectRoot "LaunchBoxYamlImporter/bin/Release/net48/YamlDotNet.dll"), \
        (Join-Path $projectRoot "LaunchBoxYamlImporter/extension.yaml") \
    ); \
    $missing = $files | Where-Object { -not (Test-Path $_) }; \
    if ($missing.Count -gt 0) { Write-Error ("Missing files: {0}" -f ($missing -join ", ")) }; \
    $destination = Join-Path $projectRoot "{{ output }}"; \
    $zipPath = [System.IO.Path]::ChangeExtension($destination, ".zip"); \
    if (Test-Path $destination) { Remove-Item $destination -Force }; \
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }; \
    Compress-Archive -Path $files -DestinationPath $zipPath -Force; \
    Move-Item -Path $zipPath -Destination $destination -Force; \
    Write-Host "Created $destination"
