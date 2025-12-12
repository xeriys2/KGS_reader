param(
  [string]$Destination = (Join-Path $PSScriptRoot "tesseract"),
  [string[]]$Langs = @("rus", "eng"),
  [string]$Url = ""
)

$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$tesseractExe = Join-Path $Destination "tesseract.exe"
$tessdataDir = Join-Path $Destination "tessdata"

if (Test-Path -LiteralPath $tesseractExe) {
  Write-Host "Tesseract already present: $tesseractExe"
} else {
  # If Tesseract is already installed on this PC, copy it to Destination (best portable path).
  $installedCandidates = @(
    (Join-Path $env:ProgramFiles "Tesseract-OCR\\tesseract.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Tesseract-OCR\\tesseract.exe"),
    (Join-Path $env:LOCALAPPDATA "Programs\\Tesseract-OCR\\tesseract.exe")
  ) | Where-Object { $_ -and $_.Trim() -ne "" }

  $installed = $installedCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
  if ($installed) {
    $srcRoot = Split-Path -Parent $installed
    Write-Host "Copying installed Tesseract from: $srcRoot"
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    Copy-Item -Recurse -Force -Path (Join-Path $srcRoot "*") -Destination $Destination
  }

  if (Test-Path -LiteralPath $tesseractExe) {
    Write-Host "Tesseract portable created from installed copy: $tesseractExe"
  } else {
  $cacheDir = Join-Path $PSScriptRoot ".cache"
  New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null

  $candidateUrls = @(
    $Url,
    "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.4.0.20240606.exe",
    "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.3.20231005.exe",
    "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe"
  ) | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -Unique

  $installerPath = Join-Path $cacheDir "tesseract-installer.exe"
  $downloaded = $false

  if (Test-Path -LiteralPath $installerPath) {
    try {
      $existing = Get-Item -LiteralPath $installerPath
      if ($existing.Length -gt 10MB) {
        Write-Host "Using cached installer: $installerPath"
        $downloaded = $true
      }
    } catch {
      $downloaded = $false
    }
  }

  foreach ($u in $candidateUrls) {
    if ($downloaded) { break }
    try {
      Write-Host "Downloading Tesseract installer: $u"
      $tmpInstaller = Join-Path $cacheDir ("tesseract-installer.{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
      Invoke-WebRequest -Uri $u -OutFile $tmpInstaller -Headers @{ "User-Agent" = "Mozilla/5.0" }
      Move-Item -Force -LiteralPath $tmpInstaller -Destination $installerPath
      $downloaded = $true
      break
    } catch {
      Write-Host "Failed: $u"
      Write-Host $_.Exception.Message
    }
  }

  if (-not $downloaded) {
    throw "Не удалось скачать Tesseract. Скачайте и распакуйте/установите вручную в папку: $Destination (ожидается tesseract.exe и tessdata\\)."
  }

  New-Item -ItemType Directory -Force -Path $Destination | Out-Null

  Write-Host "Installing Tesseract into: $Destination"

  # Most Windows installers for UB-Mannheim builds are Inno Setup (supports spaces in /DIR).
  $args = @(
    "/VERYSILENT",
    "/SUPPRESSMSGBOXES",
    "/NORESTART",
    "/SP-",
    ('/DIR="{0}"' -f $Destination)
  )

  & $installerPath @args

  if (-not (Test-Path -LiteralPath $tesseractExe)) {
    Write-Host "Silent install didn't create tesseract.exe; trying extraction (innoextract)..."

    $innoZip = Join-Path $cacheDir "innoextract-windows.zip"
    $innoDir = Join-Path $cacheDir "innoextract"
    $innoExe = Join-Path $innoDir "innoextract.exe"

    if (-not (Test-Path -LiteralPath $innoExe)) {
      Write-Host "Downloading innoextract..."
      Invoke-WebRequest -Uri "https://github.com/dscharrer/innoextract/releases/download/1.9/innoextract-1.9-windows.zip" `
        -OutFile $innoZip `
        -Headers @{ "User-Agent" = "Mozilla/5.0" }

      if (Test-Path -LiteralPath $innoDir) {
        Remove-Item -Recurse -Force -LiteralPath $innoDir
      }
      Expand-Archive -LiteralPath $innoZip -DestinationPath $innoDir -Force
    }

    if (-not (Test-Path -LiteralPath $innoExe)) {
      throw "Не удалось подготовить innoextract для распаковки Tesseract."
    }

    $extractDir = Join-Path $cacheDir "tesseract-extract"
    if (Test-Path -LiteralPath $extractDir) {
      Remove-Item -Recurse -Force -LiteralPath $extractDir
    }
    New-Item -ItemType Directory -Force -Path $extractDir | Out-Null

    & $innoExe -d $extractDir $installerPath

    $found = Get-ChildItem -LiteralPath $extractDir -Recurse -Filter "tesseract.exe" -ErrorAction SilentlyContinue | Select-Object -First 1

    if (-not $found) {
      Write-Host "Extraction failed/unsupported; trying NSIS-style silent install to a temp path..."

      $tempInstallDir = Join-Path $env:TEMP "tesseract_portable_install"
      if (Test-Path -LiteralPath $tempInstallDir) {
        Remove-Item -Recurse -Force -LiteralPath $tempInstallDir
      }
      New-Item -ItemType Directory -Force -Path $tempInstallDir | Out-Null

      # NSIS: /S (silent), /D=path (must be last, typically cannot contain quotes/spaces)
      $nsisArgs = @("/S", ("/D={0}" -f $tempInstallDir))
      Start-Process -FilePath $installerPath -ArgumentList $nsisArgs -Wait | Out-Null

      $found = Get-ChildItem -LiteralPath $tempInstallDir -Recurse -Filter "tesseract.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
      if (-not $found) {
        throw "Не удалось получить tesseract.exe автоматически. Запустите $installerPath вручную и установите в папку '$Destination'."
      }

      $root = Split-Path -Parent $found.FullName
      New-Item -ItemType Directory -Force -Path $Destination | Out-Null
      Copy-Item -Recurse -Force -Path (Join-Path $root "*") -Destination $Destination
    } else {
      $root = Split-Path -Parent $found.FullName
      New-Item -ItemType Directory -Force -Path $Destination | Out-Null
      Copy-Item -Recurse -Force -Path (Join-Path $root "*") -Destination $Destination
    }
  }
  }
}

New-Item -ItemType Directory -Force -Path $tessdataDir | Out-Null

# Ensure traineddata files exist (download minimal 'fast' models if missing)
foreach ($lang in $Langs) {
  $target = Join-Path $tessdataDir ("{0}.traineddata" -f $lang)
  if (Test-Path -LiteralPath $target) {
    continue
  }

  $tessdataUrl = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/{0}.traineddata" -f $lang
  Write-Host "Downloading tessdata: $tessdataUrl"
  Invoke-WebRequest -Uri $tessdataUrl -OutFile $target
}

Write-Host "Tesseract portable ready: $tesseractExe"
