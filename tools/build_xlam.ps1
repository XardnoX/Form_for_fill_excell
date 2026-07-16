param([string]$Output = "")

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$sourceRoot = Join-Path $projectRoot "src"
$vbaRoot = Join-Path $sourceRoot "vba"
$formsRoot = Join-Path $sourceRoot "forms"
$ribbonFile = Join-Path $sourceRoot "ribbon\customUI14.xml"
$logoFile = Join-Path $projectRoot "assets\MR_Helper_logo.jpg"
$ribbonLogoFile = Join-Path $projectRoot "assets\MR_Helper_ribbon.png"

if ([string]::IsNullOrWhiteSpace($Output)) {
    $Output = Join-Path $projectRoot "Build\MR_Helper.xlam"
}

$Output = [IO.Path]::GetFullPath($Output)
New-Item -ItemType Directory -Force (Split-Path -Parent $Output) | Out-Null

$required = @(
    (Join-Path $vbaRoot "modGlobals.bas"),
    (Join-Path $vbaRoot "modMouseWheel.bas"),
    (Join-Path $vbaRoot "modValueHistory.bas"),
    (Join-Path $vbaRoot "CChangeRecord.cls"),
    (Join-Path $vbaRoot "modPrefillEngine.bas"),
    (Join-Path $formsRoot "frmPrefill.frm"),
    (Join-Path $formsRoot "frmPhraseManager.frm"),
    (Join-Path $formsRoot "frmReview.frm"),
    $ribbonFile,
    $logoFile,
    $ribbonLogoFile
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
    $component.Properties.Item("Caption").Value = "MR_Helper"
    $component.Properties.Item("Width").Value = 620
    $component.Properties.Item("Height").Value = 540
    $designer = $component.Designer

    $header = Add-Control $designer "Forms.Label.1" "lblHeader" "MR_HELPER" 0 0 612 44
    $header.BackColor = 16777215
    $header.ForeColor = 3615007
    $header.TextAlign = 2
    $logo = Add-Control $designer "Forms.Image.1" "imgLogo" $null 8 2 64 36
    $logo.BackStyle = 0
    $logo.BorderStyle = 0
    $logo.PictureSizeMode = 3
    $logo.SpecialEffect = 0

    [void](Add-Control $designer "Forms.Label.1" "lblCount" "" 18 55 565 20)

    [void](Add-Control $designer "Forms.Label.1" "lblSheetSearch" "Hledat list:" 18 82 136 16)
    [void](Add-Control $designer "Forms.TextBox.1" "txtSheetSearch" $null 18 102 136 24)
    $allSheets = Add-Control $designer "Forms.CheckBox.1" "chkAllSheets" "" 18 132 14 20
    $allSheets.Value = $true
    [void](Add-Control $designer "Forms.Label.1" "lblAllSheets" "Vybrat vše / zrušit výběr" 36 134 116 16)
    $sheets = Add-Control $designer "Forms.ListBox.1" "lstSheets" $null 18 158 136 232
    $sheets.MultiSelect = 1

    [void](Add-Control $designer "Forms.Label.1" "lblPhrasesHeader" "Nalezená slovní spojení a hodnoty" 160 78 437 16)
    $frame = Add-Control $designer "Forms.Frame.1" "fraPhrases" "" 160 94 437 296
    $frame.ScrollBars = 2
    $frame.KeepScrollBarsVisible = 2

    [void](Add-Control $designer "Forms.CommandButton.1" "cmdManage" "Spravovat slovní spojení" 18 408 160 28)
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
    $component.Properties.Item("Caption").Value = "MR_Helper – Správa slovních spojení"
    $component.Properties.Item("Width").Value = 520
    $component.Properties.Item("Height").Value = 390
    $designer = $component.Designer

    $managerHeader = Add-Control $designer "Forms.Label.1" "lblInfo" "MR_HELPER – ULOŽENÁ SLOVNÍ SPOJENÍ" 0 0 512 44
    $managerHeader.BackColor = 16777215
    $managerHeader.ForeColor = 3615007
    $managerHeader.TextAlign = 2
    $managerLogo = Add-Control $designer "Forms.Image.1" "imgLogo" $null 8 2 64 36
    $managerLogo.BackStyle = 0
    $managerLogo.BorderStyle = 0
    $managerLogo.PictureSizeMode = 3
    $managerLogo.SpecialEffect = 0
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
    $component.Properties.Item("Caption").Value = "MR_Helper – Kontrola provedených změn"
    $component.Properties.Item("Width").Value = 480
    $component.Properties.Item("Height").Value = 480
    $component.Properties.Item("StartUpPosition").Value = 0
    $designer = $component.Designer

    $reviewHeader = Add-Control $designer "Forms.Label.1" "lblHeader" "MR_HELPER – KONTROLA PROVEDENÝCH ZMĚN" 0 0 472 44
    $reviewHeader.BackColor = 16777215
    $reviewHeader.ForeColor = 3615007
    $reviewHeader.TextAlign = 2
    $reviewLogo = Add-Control $designer "Forms.Image.1" "imgLogo" $null 8 2 64 36
    $reviewLogo.BackStyle = 0
    $reviewLogo.BorderStyle = 0
    $reviewLogo.PictureSizeMode = 3
    $reviewLogo.SpecialEffect = 0
    [void](Add-Control $designer "Forms.Label.1" "lblCount" "Provedené změny: 0" 16 55 416 20)
    [void](Add-Control $designer "Forms.ListBox.1" "lstChanges" $null 16 82 416 220)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdBack" "Zpět do formuláře" 16 320 96 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdPrevious" "▲" 120 320 36 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdNext" "▼" 120 352 36 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdDetail" "Detail" 164 320 60 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdUndo" "Zrušit změnu" 232 320 80 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdUndoAll" "Vrátit vše" 320 320 68 28)

    $finish = Add-Control $designer "Forms.CommandButton.1" "cmdFinish" "Dokončit" 368 384 96 28
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

    $vbProject.Name = "MR_Helper"

    # Standardní moduly: 1 = vbext_ct_StdModule.
    [void](Add-VbaComponent $vbProject 1 "modGlobals" (Join-Path $vbaRoot "modGlobals.bas"))
    [void](Add-VbaComponent $vbProject 1 "modMouseWheel" (Join-Path $vbaRoot "modMouseWheel.bas"))
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
$tempDirectory = Join-Path $env:TEMP ("MR_Helper_" + [guid]::NewGuid())

try {
    [IO.Compression.ZipFile]::ExtractToDirectory($Output, $tempDirectory)

    $customUiDirectory = Join-Path $tempDirectory "customUI"
    New-Item -ItemType Directory -Force $customUiDirectory | Out-Null

    # Ribbon XML se kopíruje bajtově. Nedochází k textovému překódování.
    Copy-Item -LiteralPath $ribbonFile -Destination (Join-Path $customUiDirectory "customUI14.xml") -Force

    $customUiImagesDirectory = Join-Path $customUiDirectory "images"
    New-Item -ItemType Directory -Force $customUiImagesDirectory | Out-Null
    Copy-Item -LiteralPath $ribbonLogoFile -Destination (Join-Path $customUiImagesDirectory "MR_Helper_ribbon.png") -Force

    $customUiRelsDirectory = Join-Path $customUiDirectory "_rels"
    New-Item -ItemType Directory -Force $customUiRelsDirectory | Out-Null
    $customUiRelationships = New-Object Xml.XmlDocument
    $customUiRelationships.LoadXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="MRHelperLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/MR_Helper_ribbon.png"/></Relationships>')
    $customUiRelationships.Save((Join-Path $customUiRelsDirectory "customUI14.xml.rels"))

    $contentTypesFile = Join-Path $tempDirectory "[Content_Types].xml"
    $contentTypes = New-Object Xml.XmlDocument
    $contentTypes.PreserveWhitespace = $true
    $contentTypes.Load($contentTypesFile)
    $pngType = $contentTypes.SelectSingleNode("/*[local-name()='Types']/*[local-name()='Default' and @Extension='png']")
    if ($null -eq $pngType) {
        $defaultPng = $contentTypes.CreateElement("Default", $contentTypes.DocumentElement.NamespaceURI)
        $defaultPng.SetAttribute("Extension", "png")
        $defaultPng.SetAttribute("ContentType", "image/png")
        [void]$contentTypes.DocumentElement.AppendChild($defaultPng)
        $contentTypes.Save($contentTypesFile)
    }

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

$assetOutputDirectory = Join-Path (Split-Path -Parent $Output) "MR_Helper_assets"
New-Item -ItemType Directory -Force $assetOutputDirectory | Out-Null
$logoOutput = Join-Path $assetOutputDirectory "MR_Helper_logo.jpg"
Copy-Item -LiteralPath $logoFile -Destination $logoOutput -Force

Write-Host "Vytvořeno: $Output"
Write-Host "Logo: $logoOutput"
