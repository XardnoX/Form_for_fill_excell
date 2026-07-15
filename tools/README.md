# Nástroje sestavení

## Pouze XLAM

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\build_xlam.ps1
```

Výstup:

```text
Build\FormularProPredvyplneni.xlam
```

## XLAM a instalační EXE

Spusťte z kořenové složky:

```powershell
powershell -ExecutionPolicy Bypass -File .\Build-Installer.ps1
```

Výstup:

```text
Output\FormularProPredvyplneni_Setup.exe
Output\FormularProPredvyplneni_Setup_SHA256.txt
```

Požadavky: desktopový Excel pro Windows a Inno Setup 7. Před sestavením musí být Excel zavřený.
