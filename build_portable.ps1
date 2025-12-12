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
  --name $name `
  --icon $icon `
  --add-data "comm_types.json;." `
  $(if ($hasTess) { @("--add-data","tesseract;tesseract") } else { @() }) `
  $entry

Write-Host "Done. Output: dist\\$name\\"
