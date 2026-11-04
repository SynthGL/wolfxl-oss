# Install an official wolfxl CLI release. Set WOLFXL_VERSION to pin a version.
$ErrorActionPreference = 'Stop'
$releaseBase = if ($env:WOLFXL_RELEASE_BASE_URL) { $env:WOLFXL_RELEASE_BASE_URL.TrimEnd('/') } else { 'https://github.com/SynthGL/wolfxl-oss/releases/download' }
$contextUrl = $releaseBase
$tempDir = $null
try {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw 'This installer supports Windows only; use install.sh on Linux or macOS.'
    }
    # Use the OS architecture, not the architecture of a 32-bit PowerShell process.
    $architecture = if ($env:PROCESSOR_ARCHITEW6432) { $env:PROCESSOR_ARCHITEW6432 } else { $env:PROCESSOR_ARCHITECTURE }
    if ($architecture -ne 'AMD64') { throw "Unsupported Windows architecture: $architecture" }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $tempDir = Join-Path ([IO.Path]::GetTempPath()) ('wolfxl-install-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    $version = $env:WOLFXL_VERSION
    if (-not $version) {
        # /latest tracks the Python package, not the CLI. Scan all release pages.
        $candidates = @()
        $page = 1
        do {
            $contextUrl = "https://api.github.com/repos/SynthGL/wolfxl-oss/releases?per_page=100&page=$page"
            $releases = Invoke-RestMethod -Uri $contextUrl -Headers @{ Accept = 'application/vnd.github+json' }
            $candidates += @($releases | Where-Object {
                -not $_.draft -and -not $_.prerelease -and $_.tag_name -cmatch '^wolfxl-cli-v[0-9]+\.[0-9]+\.[0-9]+$'
            })
            $page++
        } while ($releases.Count -eq 100)
        $release = $candidates | Sort-Object published_at -Descending | Select-Object -First 1
        if (-not $release) { throw 'No stable wolfxl-cli release found.' }
        $version = $release.tag_name
    }
    $version = $version -creplace '^wolfxl-cli-v', '' -creplace '^v', ''
    if ($version -cnotmatch '^[0-9A-Za-z.+-]+$') { throw "Invalid WOLFXL_VERSION: $version" }
    $asset = "wolfxl-cli-$version-x86_64-pc-windows-msvc.zip"
    $contextUrl = "$releaseBase/wolfxl-cli-v$version/SHA256SUMS"
    $checksums = Join-Path $tempDir 'SHA256SUMS'
    Invoke-WebRequest -UseBasicParsing -Uri $contextUrl -OutFile $checksums
    $entries = @(Get-Content -LiteralPath $checksums | Where-Object { $_ -cmatch ('^[0-9a-fA-F]{64}\s+\*?' + [regex]::Escape($asset) + '$') })
    if ($entries.Count -ne 1) { throw "Expected exactly one SHA256 checksum for $asset" }
    $expected = ($entries[0] -split '\s+')[0]
    $contextUrl = "$releaseBase/wolfxl-cli-v$version/$asset"
    $archive = Join-Path $tempDir $asset
    Invoke-WebRequest -UseBasicParsing -Uri $contextUrl -OutFile $archive
    $actual = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash
    if ($actual -ine $expected) { throw "SHA256 mismatch for $asset; nothing was installed." }
    # Read only the executable entry instead of extracting arbitrary archive paths.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try {
        $executables = @($zip.Entries | Where-Object { $_.FullName -ceq 'wolfxl.exe' })
        if ($executables.Count -ne 1) { throw 'Archive must contain exactly one wolfxl.exe.' }
        $executable = Join-Path $tempDir 'wolfxl.exe'
        [IO.Compression.ZipFileExtensions]::ExtractToFile($executables[0], $executable)
    } finally { $zip.Dispose() }
    $installDir = if ($env:WOLFXL_INSTALL_DIR) { $env:WOLFXL_INSTALL_DIR } else { Join-Path $HOME '.local\bin' }
    New-Item -ItemType Directory -Force -Path $installDir | Out-Null
    Copy-Item -LiteralPath $executable -Destination (Join-Path $installDir 'wolfxl.exe') -Force
    Write-Output "Installed wolfxl $version to $installDir\wolfxl.exe"
    if (($env:PATH -split ';') -notcontains $installDir) {
        Write-Output "Add this directory to your user PATH, then open a new terminal: $installDir"
    }
} catch {
    throw "wolfxl installer: $($_.Exception.Message)`nURL: $contextUrl"
} finally {
    if ($tempDir -and (Test-Path -LiteralPath $tempDir)) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}
