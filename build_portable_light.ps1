$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

$name = "KGS_Reader"
$entry = "KGS_Reader v6.py"
$icon = "icon.ico"

$venvDir = Join-Path $PSScriptRoot ".venv_build"
$venvPython = Join-Path $venvDir "Scripts\\python.exe"

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
  throw "Python не найден в PATH. Установите Python 3.10+ и повторите."
}

if (-not (Test-Path -LiteralPath $venvPython)) {
  Write-Host "Creating build venv: $venvDir"
  python -m venv $venvDir
}

Write-Host "Installing minimal build dependencies..."
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r (Join-Path $PSScriptRoot "requirements_build.txt")

$tess = Join-Path $PSScriptRoot "tesseract"
$hasTess = Test-Path (Join-Path $tess "tesseract.exe")
$distDir = Join-Path $PSScriptRoot ("dist\\{0}" -f $name)
$buildDir = Join-Path $PSScriptRoot "build"

if (Test-Path -LiteralPath $distDir) {
  Remove-Item -Recurse -Force -LiteralPath $distDir
}
if (Test-Path -LiteralPath $buildDir) {
  Remove-Item -Recurse -Force -LiteralPath $buildDir
}

Write-Host "Building $name from '$entry' (lightweight venv)..."

if ($hasTess) {
  Write-Host "Including portable Tesseract: $tess"
} else {
  Write-Host "Portable Tesseract not found at .\\tesseract\\tesseract.exe (OCR will require installed Tesseract)."
}

& $venvPython -m PyInstaller `
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

Write-Host "Done. Output: dist\\$name\\"
