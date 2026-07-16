param(
    [string]$InnoSetupCompiler = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$buildDirectory = Join-Path $projectRoot "Build"
$outputDirectory = Join-Path $projectRoot "Output"
$addinFile = Join-Path $buildDirectory "MR_Helper.xlam"
$installerFile = Join-Path $outputDirectory "MR_Helper_Setup.exe"
$securityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"

New-Item -ItemType Directory -Force $buildDirectory, $outputDirectory | Out-Null

$hadAccessVbom = $false
$oldAccessVbom = $null
if (!(Test-Path $securityKey)) {
    New-Item -Path $securityKey -Force | Out-Null
}
try {
    $oldAccessVbom = (Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction Stop).AccessVBOM
    $hadAccessVbom = $true
}
catch {}

try {
    if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
        throw "Excel je spuštěný. Zavřete všechny instance Excelu a spusťte skript znovu."
    }

    # Přístup je povolen pouze během sestavení a v finally se obnoví původní stav.
    Set-ItemProperty -Path $securityKey -Name AccessVBOM -Type DWord -Value 1

    Write-Host "[1/3] Sestavuji MR_Helper.xlam..."
    & (Join-Path $projectRoot "tools\build_xlam.ps1") -Output $addinFile
    if (!(Test-Path -LiteralPath $addinFile)) {
        throw "Soubor XLAM nebyl vytvořen."
    }

    if ([string]::IsNullOrWhiteSpace($InnoSetupCompiler)) {
        $candidates = @(
            "${env:ProgramFiles(x86)}\Inno Setup 7\ISCC.exe",
            "$env:ProgramFiles\Inno Setup 7\ISCC.exe",
            "$env:LOCALAPPDATA\Programs\Inno Setup 7\ISCC.exe",
            "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
            "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
        )
        $InnoSetupCompiler = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    }
    if ([string]::IsNullOrWhiteSpace($InnoSetupCompiler) -or !(Test-Path -LiteralPath $InnoSetupCompiler)) {
        throw "Nenalezen ISCC.exe. Použijte parametr -InnoSetupCompiler s cestou k Inno Setup 7."
    }

    Write-Host "[2/3] Kompiluji instalační EXE..."
    & $InnoSetupCompiler (Join-Path $projectRoot "installer.iss")
    if ($LASTEXITCODE -ne 0) {
        throw "Kompilace Inno Setup selhala s kódem $LASTEXITCODE."
    }
    if (!(Test-Path -LiteralPath $installerFile)) {
        throw "Instalační EXE nebylo vytvořeno: $installerFile"
    }

    Write-Host "[3/3] Vytvářím kontrolní součet SHA-256..."
    $hash = Get-FileHash -LiteralPath $installerFile -Algorithm SHA256
    "$($hash.Hash)  MR_Helper_Setup.exe" | Set-Content `
        -LiteralPath (Join-Path $outputDirectory "MR_Helper_Setup_SHA256.txt") `
        -Encoding ASCII

    Write-Host "Hotovo: $installerFile"
}
finally {
    if ($hadAccessVbom) {
        Set-ItemProperty -Path $securityKey -Name AccessVBOM -Type DWord -Value $oldAccessVbom -ErrorAction SilentlyContinue
    }
    else {
        Remove-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    }
}
