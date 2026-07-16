# Formulář pro předvyplnění

Kompletní zdrojový projekt Excel VBA Add-inu. Projekt neobsahuje instalační EXE ani předem sestavený XLAM.

## Funkce

- zobrazení pouze slovních spojení nalezených na vybraných viditelných listech,
- všechny viditelné listy jsou při otevření vybrané,
- počet výskytů každého slovního spojení,
- prázdné vstupní pole bez rozbalovacího tlačítka,
- automatické napovídání podle dříve použitých hodnot při psaní,
- správa, přejmenování a odstranění slovních spojení,
- kontrola duplicitních zásahů do stejné cílové buňky,
- revizní okno a vrácení jedné nebo všech změn,
- návrat z revizního okna zpět do formuláře,
- obnova původní hodnoty, vzorce, číselného formátu a výplně.

## Import do vývojového XLSM

1. Otevřete nový sešit a uložte jej jako `FormularProPredvyplneni_DEV.xlsm`.
2. Stiskněte `Alt+F11`.
3. Importujte postupně soubory `src/vba/*.bas`, `src/vba/*.cls` a `src/forms/*.frm`.
4. Spusťte makro `ShowPrefillForm` přes `Alt+F8`.

Ribbon XML není pro spuštění přes `Alt+F8` nutný. Přidává se až při tvorbě XLAM.

## Git

```bash
git add .
git commit -m "Add complete source project with suggestion history"
git push
```

## Sestavení

Pouze doplněk XLAM:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\build_xlam.ps1
```

Doplněk a instalační EXE:

```powershell
powershell -ExecutionPolicy Bypass -File .\Build-Installer.ps1
```

Vygenerované soubory se ukládají do `Build/` a `Output/`. Tyto složky jsou záměrně uvedeny v `.gitignore`; zdrojové sestavovací skripty `Build-Installer.ps1`, `installer.iss` a `tools/build_xlam.ps1` se naopak verzují v Gitu.
