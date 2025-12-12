param(
  [switch]$WithoutOcr,
  [switch]$NoZip
)

$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
  throw "PyInstaller не найден. Установите: pip install pyinstaller"
}

$name = "KGS_Reader"
$entry = "KGS_Reader v6.py"
$icon = "icon.ico"
$tess = Join-Path $PSScriptRoot "tesseract"
$hasTess = Test-Path (Join-Path $tess "tesseract.exe")
$distDir = Join-Path $PSScriptRoot ("dist\\{0}" -f $name)

if (-not $WithoutOcr -and -not $hasTess) {
  Write-Host "Portable Tesseract not found. Preparing portable OCR bundle..."
  & (Join-Path $PSScriptRoot "setup_tesseract_portable.ps1") -Destination $tess -Langs @("rus", "eng")
  $hasTess = Test-Path (Join-Path $tess "tesseract.exe")
}

Write-Host "Building $name from '$entry'..."

if (-not $WithoutOcr) {
  if ($hasTess) {
    Write-Host "Including portable Tesseract: $tess"
  } else {
    throw "OCR requested but portable Tesseract is missing: $tess\\tesseract.exe"
  }
} else {
  Write-Host "OCR bundle disabled (-WithoutOcr)."
}

pyinstaller `
  --noconsole `
  --clean `
  --onedir `
  -y `
  --name $name `
  --icon $icon `
  $entry

if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller failed with exit code $LASTEXITCODE"
}

if (!(Test-Path -LiteralPath $distDir)) {
  throw "Dist folder not found: $distDir"
}

# Copy editable runtime files next to exe (portable-friendly)
Copy-Item -Force -LiteralPath (Join-Path $PSScriptRoot "comm_types.json") -Destination (Join-Path $distDir "comm_types.json")
Copy-Item -Force -LiteralPath (Join-Path $PSScriptRoot "icon.ico") -Destination (Join-Path $distDir "icon.ico")
Copy-Item -Force -LiteralPath (Join-Path $PSScriptRoot "LICENSE") -Destination (Join-Path $distDir "LICENSE")
Copy-Item -Force -LiteralPath (Join-Path $PSScriptRoot "README.md") -Destination (Join-Path $distDir "README.md")

if ($hasTess) {
  Copy-Item -Recurse -Force -LiteralPath $tess -Destination (Join-Path $distDir "tesseract")
}

if (-not $NoZip) {
  $releaseDir = Join-Path $PSScriptRoot "release"
  New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
  $zipPath = Join-Path $releaseDir "KGS_Reader_portable.zip"
  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -Force -LiteralPath $zipPath
  }
  try {
    $maxAttempts = 6
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
      try {
        Compress-Archive -Path (Join-Path $distDir "*") -DestinationPath $zipPath
        Write-Host "ZIP created: $zipPath"
        break
      } catch {
        if ($attempt -eq $maxAttempts) { throw }
        Start-Sleep -Seconds 2
      }
    }
  } catch {
    Write-Warning ("ZIP creation failed (you can just copy dist\\{0}\\). Error: {1}" -f $name, $_.Exception.Message)
  }
}

Write-Host "Done. Output: dist\\$name\\"
