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
    if (!(Test-Path -LiteralPath $file)) { throw "Chybí povinný soubor: $file" }
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Excel je spuštěný. Zavřete všechny instance Excelu."
}

function Get-FormCode([string]$Path) {
    # .frm soubory jsou vedené v Gitu jako čitelné zdroje. Formulář se při buildu
    # vytvoří přes VBIDE a z .frm se načte pouze část kódu od Option Explicit.
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding(1250))
    $position = $text.IndexOf("Option Explicit")
    if ($position -lt 0) { throw "Ve formuláři chybí Option Explicit: $Path" }
    return $text.Substring($position)
}

function Add-Control($designer, [string]$progId, [string]$name, $caption, $left, $top, $width, $height) {
    [void]$designer.Controls.Add($progId, $name, $true)
    $control = $designer.Controls.Item($name)
    if ($null -eq $control) { throw "Nepodařilo se vytvořit prvek: $name" }
    if ($null -ne $caption) { $control.Caption = $caption }
    $control.Left = $left; $control.Top = $top
    $control.Width = $width; $control.Height = $height
    return $control
}

function Add-PrefillForm($vbProject) {
    $component = $vbProject.VBComponents.Add(3)
    $component.Name = "frmPrefill"
    $component.Properties.Item("Caption").Value = "Formulář pro předvyplnění"
    $component.Properties.Item("Width").Value = 760
    $component.Properties.Item("Height").Value = 560
    $designer = $component.Designer

    $header = Add-Control $designer "Forms.Label.1" "lblHeader" "FORMULÁŘ PRO PŘEDVYPLNĚNÍ" 0 0 752 44
    $header.BackColor = 7487778; $header.ForeColor = 16777215; $header.TextAlign = 2
    [void](Add-Control $designer "Forms.Label.1" "lblCount" "" 18 55 700 20)
    $sheets = Add-Control $designer "Forms.ListBox.1" "lstSheets" $null 18 82 180 330
    $sheets.MultiSelect = 1
    $frame = Add-Control $designer "Forms.Frame.1" "fraPhrases" "Nalezená slovní spojení a hodnoty" 215 82 515 330
    $frame.ScrollBars = 2; $frame.KeepScrollBarsVisible = 2
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdRefresh" "Obnovit výskyty" 18 425 180 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdManage" "Spravovat slovní spojení" 215 425 170 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdCancel" "Zrušit" 555 475 80 30)
    $save = Add-Control $designer "Forms.CommandButton.1" "cmdSave" "Použít změny" 645 475 85 30
    $save.Default = $true
    $component.CodeModule.AddFromString((Get-FormCode (Join-Path $formsRoot "frmPrefill.frm")))
}

function Add-PhraseManagerForm($vbProject) {
    $component = $vbProject.VBComponents.Add(3)
    $component.Name = "frmPhraseManager"
    $component.Properties.Item("Caption").Value = "Správa slovních spojení"
    $component.Properties.Item("Width").Value = 520
    $component.Properties.Item("Height").Value = 390
    $designer = $component.Designer

    [void](Add-Control $designer "Forms.Label.1" "lblInfo" "Uložená slovní spojení" 16 16 470 20)
    [void](Add-Control $designer "Forms.ListBox.1" "lstPhrases" $null 16 42 470 230)
    [void](Add-Control $designer "Forms.TextBox.1" "txtPhrase" $null 16 285 270 25)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdAdd" "Přidat" 295 284 60 27)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdRename" "Přejmenovat" 360 284 75 27)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdDelete" "Odstranit" 440 284 60 27)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdClose" "Zavřít" 420 325 80 28)
    $component.CodeModule.AddFromString((Get-FormCode (Join-Path $formsRoot "frmPhraseManager.frm")))
}

function Add-ReviewForm($vbProject) {
    $component = $vbProject.VBComponents.Add(3)
    $component.Name = "frmReview"
    $component.Properties.Item("Caption").Value = "Kontrola provedených změn"
    $component.Properties.Item("Width").Value = 760
    $component.Properties.Item("Height").Value = 440
    $designer = $component.Designer

    [void](Add-Control $designer "Forms.ListBox.1" "lstChanges" $null 12 12 720 320)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdGoTo" "Přejít na buňku" 250 345 105 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdDetail" "Detail" 363 345 75 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdUndo" "Zrušit změnu" 446 345 90 28)
    [void](Add-Control $designer "Forms.CommandButton.1" "cmdUndoAll" "Vrátit vše" 544 345 80 28)
    $finish = Add-Control $designer "Forms.CommandButton.1" "cmdFinish" "Dokončit" 632 345 90 28
    $finish.Default = $true
    $component.CodeModule.AddFromString((Get-FormCode (Join-Path $formsRoot "frmReview.frm")))
}

$excel = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) { $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete() }

    $vbProject = $workbook.VBProject
    if ($null -eq $vbProject) { throw "Excel nevrátil projekt VBA. Ověřte AccessVBOM." }

    foreach ($file in @("modGlobals.bas", "modValueHistory.bas", "CChangeRecord.cls", "modPrefillEngine.bas")) {
        [void]$vbProject.VBComponents.Import((Join-Path $vbaRoot $file))
    }
    Add-PrefillForm $vbProject
    Add-PhraseManagerForm $vbProject
    Add-ReviewForm $vbProject

    $workbook.IsAddin = $true
    if (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Force }
    $workbook.SaveAs($Output, 55)
    $workbook.Close($true); $workbook = $null
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
}

if (!(Test-Path -LiteralPath $Output)) { throw "XLAM nebyl vytvořen: $Output" }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$tempDirectory = Join-Path $env:TEMP ("FormularProPredvyplneni_" + [guid]::NewGuid())
try {
    [IO.Compression.ZipFile]::ExtractToDirectory($Output, $tempDirectory)
    $customUiDirectory = Join-Path $tempDirectory "customUI"
    New-Item -ItemType Directory -Force $customUiDirectory | Out-Null
    Copy-Item -LiteralPath $ribbonFile -Destination (Join-Path $customUiDirectory "customUI14.xml") -Force

    $relationshipsFile = Join-Path $tempDirectory "_rels\.rels"
    [xml]$relationships = Get-Content -LiteralPath $relationshipsFile -Raw
    $namespace = $relationships.DocumentElement.NamespaceURI
    $relationship = $relationships.CreateElement("Relationship", $namespace)
    $relationship.SetAttribute("Id", "rIdCustomUI")
    $relationship.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
    $relationship.SetAttribute("Target", "customUI/customUI14.xml")
    [void]$relationships.DocumentElement.AppendChild($relationship)
    $relationships.Save($relationshipsFile)

    Remove-Item -LiteralPath $Output -Force
    [IO.Compression.ZipFile]::CreateFromDirectory($tempDirectory, $Output)
}
finally {
    if (Test-Path -LiteralPath $tempDirectory) { Remove-Item -LiteralPath $tempDirectory -Recurse -Force }
}
Write-Host "Vytvořeno: $Output"
