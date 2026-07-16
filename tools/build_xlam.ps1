param([string]$Output = "")

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$sourceRoot = Join-Path $projectRoot "src"
$vbaRoot = Join-Path $sourceRoot "vba"
$formsRoot = Join-Path $sourceRoot "forms"
$ribbonFile = Join-Path $sourceRoot "ribbon\customUI14.xml"

if ([string]::IsNullOrWhiteSpace($Output)) {
    $Output = Join-Path $projectRoot "Build\FormularProPredvyplneni.xlam"
}

$Output = [IO.Path]::GetFullPath($Output)
New-Item -ItemType Directory -Force (Split-Path -Parent $Output) | Out-Null

$required = @(
    (Join-Path $vbaRoot "modGlobals.bas"),
    (Join-Path $vbaRoot "modValueHistory.bas"),
    (Join-Path $vbaRoot "CChangeRecord.cls"),
    (Join-Path $vbaRoot "modPrefillEngine.bas"),
    (Join-Path $formsRoot "frmPrefill.frm"),
    (Join-Path $formsRoot "frmPhraseManager.frm"),
    (Join-Path $formsRoot "frmReview.frm"),
    $ribbonFile
)

foreach ($file in $required) {
    if (-not (Test-Path -LiteralPath $file)) {
        throw "Chybí povinný soubor: $file"
    }
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Excel je spuštěný. Zavřete všechny instance Excelu."
}

function Read-Utf8TextFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $bytes = [IO.File]::ReadAllBytes($resolvedPath)
    $utf8Strict = New-Object Text.UTF8Encoding($false, $true)

    try {
        $text = $utf8Strict.GetString($bytes)
    }
    catch {
        throw "Soubor není platné UTF-8: $resolvedPath. $($_.Exception.Message)"
    }

    # Odstranění správně dekódovaného UTF-8 BOM.
    if ($text.Length -gt 0 -and [int][char]$text[0] -eq 0xFEFF) {
        $text = $text.Substring(1)
    }

    # Ochrana proti BOM, který byl už dříve chybně převeden na viditelný text.
    $text = $text -replace '^(ï»¿|ď»ż)', ''

    if ($text.IndexOf([char]0xFFFD) -ge 0) {
        throw "Soubor obsahuje znak U+FFFD a je textově poškozen: $resolvedPath"
    }

    return $text
}

function Get-VbaCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $text = Read-Utf8TextFile -Path $Path

    # Exportované .bas, .cls a .frm mohou obsahovat hlavičky VERSION, BEGIN/END
    # a Attribute VB_*. Při programovém vytvoření komponenty přes Add(...) se
    # do CodeModule.AddFromString vkládá pouze vlastní kód od Option Explicit.
    $match = [regex]::Match(
        $text,
        '(?im)^\s*Option\s+Explicit\s*$'
    )

    if (-not $match.Success) {
        throw "Ve VBA zdroji chybí Option Explicit: $Path"
    }

    $code = $text.Substring($match.Index)

    # Normalizace konců řádků pro VBE.
    $code = $code.Replace("`r`n", "`n")
    $code = $code.Replace("`r", "`n")
    $code = $code.Replace("`n", "`r`n")

    if ($code -match '(?im)^\s*Attribute\s+VB_') {
        throw "Kód určený pro AddFromString stále obsahuje Attribute VB_*: $Path"
    }

    if ($code -match '^(ï»¿|ď»ż)' -or ($code.Length -gt 0 -and [int][char]$code[0] -eq 0xFEFF)) {
        throw "Kód určený pro AddFromString stále obsahuje BOM: $Path"
    }

    return $code
}

function Add-VbaComponent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [ValidateSet(1, 2)]
        [int]$ComponentType,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ComponentName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SourcePath
    )

    $component = $VBProject.VBComponents.Add($ComponentType)
    $component.Name = $ComponentName

    $code = Get-VbaCode -Path $SourcePath
    $component.CodeModule.AddFromString($code)

    return $component
}

function Add-Control {
    param(
        $Designer,
        [string]$ProgId,
        [string]$Name,
        $Caption,
        $Left,
        $Top,
        $Width,
        $Height
    )

    [void]$Designer.Controls.Add($ProgId, $Name, $true)
    $control = $Designer.Controls.Item($Name)

    if ($null -eq $control) {
        throw "Nepodařilo se vytvořit prvek: $Name"
    }

    if ($null -ne $Caption) {
        $control.Caption = $Caption
    }

    $control.Left = $Left
    $control.Top = $Top
    $control.Width = $Width
    $control.Height = $Height

    return $control
}

function Add-PrefillForm {
    param($VBProject)

    # 3 = vbext_ct_MSForm. Formulář se nevytváří importem .frm.
    $component = $VBProject.VBComponents.Add(3)
    $component.Name = "frmPrefill"
    $component.Properties.Item("Caption").Value = "Formulář pro předvyplnění"
    $component.Properties.Item("Width").Value = 620
    $component.Properties.Item("Height").Value = 540
    $designer = $component.Designer

    $header = Add-Control $designer "Forms.Label.1" "lblHeader" "FORMULÁŘ PRO PŘEDVYPLNĚNÍ" 0 0 612 44
    $header.BackColor = 5195062
    $header.ForeColor = 16777215
    $header.TextAlign = 2

    [void](Add-Control $designer "Forms.Label.1" "lblCount" "" 18 55 565 20)

    $sheets = Add-Control $designer "Forms.ListBox.1" "lstSheets" $null 18 82 136 308
    $sheets.MultiSelect = 1

    $frame = Add-Control $designer "Forms.Frame.1" "fraPhrases" "Nalezená slovní spojení a hodnoty" 160 82 437 308
    $frame.ScrollBars = 2
    $frame.KeepScrollBarsVisible = 2

    [void](Add-Control $designer "Forms.CommandButton.1" "cmdRefresh" "Obnovit výskyty" 18 408 136 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdManage" "Spravovat slovní spojení" 160 408 160 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdCancel" "Zrušit" 408 456 80 30)

    $save = Add-Control $designer "Forms.CommandButton.1" "cmdSave" "Použít změny" 496 456 96 30
    $save.Default = $true

    $code = Get-VbaCode -Path (Join-Path $formsRoot "frmPrefill.frm")
    $component.CodeModule.AddFromString($code)
}

function Add-PhraseManagerForm {
    param($VBProject)

    # 3 = vbext_ct_MSForm. Formulář se nevytváří importem .frm.
    $component = $VBProject.VBComponents.Add(3)
    $component.Name = "frmPhraseManager"
    $component.Properties.Item("Caption").Value = "Správa slovních spojení"
    $component.Properties.Item("Width").Value = 520
    $component.Properties.Item("Height").Value = 390
    $designer = $component.Designer

    $managerHeader = Add-Control $designer "Forms.Label.1" "lblInfo" "ULOŽENÁ SLOVNÍ SPOJENÍ" 0 0 512 44
    $managerHeader.BackColor = 5195062
    $managerHeader.ForeColor = 16777215
    $managerHeader.TextAlign = 2
    [void](Add-Control $designer "Forms.ListBox.1" "lstPhrases" $null 16 56 470 200)
    [void](Add-Control $designer "Forms.TextBox.1" "txtPhrase" $null 16 285 270 25)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdAdd" "Přidat" 295 284 60 27)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdRename" "Přejmenovat" 360 284 75 27)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdDelete" "Odstranit" 440 284 60 27)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdClose" "Zavřít" 420 325 80 28)

    $code = Get-VbaCode -Path (Join-Path $formsRoot "frmPhraseManager.frm")
    $component.CodeModule.AddFromString($code)
}

function Add-ReviewForm {
    param($VBProject)

    # 3 = vbext_ct_MSForm. Formulář se nevytváří importem .frm.
    $component = $VBProject.VBComponents.Add(3)
    $component.Name = "frmReview"
    $component.Properties.Item("Caption").Value = "Kontrola provedených změn"
    $component.Properties.Item("Width").Value = 760
    $component.Properties.Item("Height").Value = 440
    $designer = $component.Designer

    $reviewHeader = Add-Control $designer "Forms.Label.1" "lblHeader" "KONTROLA PROVEDENÝCH ZMĚN" 0 0 752 44
    $reviewHeader.BackColor = 5195062
    $reviewHeader.ForeColor = 16777215
    $reviewHeader.TextAlign = 2
    [void](Add-Control $designer "Forms.Label.1" "lblCount" "Provedené změny: 0" 16 55 710 20)
    [void](Add-Control $designer "Forms.ListBox.1" "lstChanges" $null 16 82 710 236)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdBack" "Zpět do formuláře" 16 345 125 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdGoTo" "Přejít na buňku" 250 345 105 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdDetail" "Detail" 363 345 75 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdUndo" "Zrušit změnu" 446 345 90 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdUndoAll" "Vrátit vše" 544 345 80 28)

    $finish = Add-Control $designer "Forms.CommandButton.1" "cmdFinish" "Dokončit" 632 345 90 28
    $finish.Default = $true

    $code = Get-VbaCode -Path (Join-Path $formsRoot "frmReview.frm")
    $component.CodeModule.AddFromString($code)
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $vbProject = $workbook.VBProject

    if ($null -eq $vbProject) {
        throw "Excel nevrátil projekt VBA. Ověřte AccessVBOM."
    }

    # Standardní moduly: 1 = vbext_ct_StdModule.
    [void](Add-VbaComponent $vbProject 1 "modGlobals" (Join-Path $vbaRoot "modGlobals.bas"))
    [void](Add-VbaComponent $vbProject 1 "modValueHistory" (Join-Path $vbaRoot "modValueHistory.bas"))

    # Třída: 2 = vbext_ct_ClassModule.
    [void](Add-VbaComponent $vbProject 2 "CChangeRecord" (Join-Path $vbaRoot "CChangeRecord.cls"))

    [void](Add-VbaComponent $vbProject 1 "modPrefillEngine" (Join-Path $vbaRoot "modPrefillEngine.bas"))

    Add-PrefillForm $vbProject
    Add-PhraseManagerForm $vbProject
    Add-ReviewForm $vbProject

    $workbook.IsAddin = $true

    if (Test-Path -LiteralPath $Output) {
        Remove-Item -LiteralPath $Output -Force
    }

    # 55 = xlOpenXMLAddIn (.xlam)
    $workbook.SaveAs($Output, 55)
    $workbook.Close($true)
    $workbook = $null
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }

    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
}

if (-not (Test-Path -LiteralPath $Output)) {
    throw "XLAM nebyl vytvořen: $Output"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$tempDirectory = Join-Path $env:TEMP ("FormularProPredvyplneni_" + [guid]::NewGuid())

try {
    [IO.Compression.ZipFile]::ExtractToDirectory($Output, $tempDirectory)

    $customUiDirectory = Join-Path $tempDirectory "customUI"
    New-Item -ItemType Directory -Force $customUiDirectory | Out-Null

    # Ribbon XML se kopíruje bajtově. Nedochází k textovému překódování.
    Copy-Item -LiteralPath $ribbonFile -Destination (Join-Path $customUiDirectory "customUI14.xml") -Force

    $relationshipsFile = Join-Path $tempDirectory "_rels\.rels"
    $relationships = New-Object Xml.XmlDocument
    $relationships.PreserveWhitespace = $true
    $relationships.Load($relationshipsFile)

    $namespace = $relationships.DocumentElement.NamespaceURI
    $existingRelationship = $relationships.SelectSingleNode(
        "/*[local-name()='Relationships']/*[local-name()='Relationship' and @Id='rIdCustomUI']"
    )

    if ($null -eq $existingRelationship) {
        $relationship = $relationships.CreateElement("Relationship", $namespace)
        $relationship.SetAttribute("Id", "rIdCustomUI")
        $relationship.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
        $relationship.SetAttribute("Target", "customUI/customUI14.xml")
        [void]$relationships.DocumentElement.AppendChild($relationship)
    }
    else {
        $existingRelationship.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
        $existingRelationship.SetAttribute("Target", "customUI/customUI14.xml")
    }

    $relationships.Save($relationshipsFile)

    Remove-Item -LiteralPath $Output -Force
    [IO.Compression.ZipFile]::CreateFromDirectory($tempDirectory, $Output)
}
finally {
    if (Test-Path -LiteralPath $tempDirectory) {
        Remove-Item -LiteralPath $tempDirectory -Recurse -Force
    }
}

Write-Host "Vytvořeno: $Output"
