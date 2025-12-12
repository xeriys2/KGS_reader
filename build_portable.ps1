$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
  Write-Error "PyInstaller не найден. Установи: pip install pyinstaller"
}

$name = "KGS_Reader"
$entry = "KGS_Reader v6.py"
$icon = "icon.ico"
$tess = Join-Path $PSScriptRoot "tesseract"
$hasTess = Test-Path (Join-Path $tess "tesseract.exe")
$distDir = Join-Path $PSScriptRoot ("dist\\{0}" -f $name)

Write-Host "Building $name from '$entry'..."

if ($hasTess) {
  Write-Host "Including portable Tesseract: $tess"
} else {
  Write-Host "Portable Tesseract not found at .\\tesseract\\tesseract.exe (OCR will require installed Tesseract)."
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

if (!(Test-Path $distDir)) {
  Write-Error "Dist folder not found: $distDir"
}

# Copy editable runtime files next to exe (portable-friendly)
Copy-Item -Force -LiteralPath (Join-Path $PSScriptRoot "comm_types.json") -Destination (Join-Path $distDir "comm_types.json")
Copy-Item -Force -LiteralPath (Join-Path $PSScriptRoot "icon.ico") -Destination (Join-Path $distDir "icon.ico")
Copy-Item -Force -LiteralPath (Join-Path $PSScriptRoot "LICENSE") -Destination (Join-Path $distDir "LICENSE")
Copy-Item -Force -LiteralPath (Join-Path $PSScriptRoot "README.md") -Destination (Join-Path $distDir "README.md")

if ($hasTess) {
  Copy-Item -Recurse -Force -LiteralPath $tess -Destination (Join-Path $distDir "tesseract")
}

Write-Host "Done. Output: dist\\$name\\"
