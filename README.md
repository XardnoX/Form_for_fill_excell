# MR_Helper

Kompletní zdrojový projekt Excel VBA Add-inu. Projekt neobsahuje instalační EXE ani předem sestavený XLAM.

Logo formulářů se instaluje do podsložky `MR_Helper_assets`; soubory obrázků se nesmí ukládat přímo do `XLSTART`, protože by je Excel při spuštění zkusil otevřít jako sešity.
Ribbon používá samostatnou čtvercovou ikonu `assets/MR_Helper_ribbon.png` bez drobného textu, aby značka zůstala čitelná při velikosti 32×32 px.

## Funkce

- zobrazení pouze slovních spojení nalezených na vybraných viditelných listech,
- všechny viditelné listy jsou při otevření vybrané,
- vyhledávání listů podle názvu a hromadný výběr všech/žádného listu,
- počet výskytů každého slovního spojení,
- prázdné vstupní pole bez rozbalovacího tlačítka,
- automatické napovídání podle dříve použitých hodnot při psaní,
- správa, přejmenování a odstranění slovních spojení,
- kontrola duplicitních zásahů do stejné cílové buňky,
- revizní okno a vrácení jedné nebo všech změn,
- návrat z revizního okna zpět do formuláře,
- volitelné rolování seznamů a hodnot kolečkem myši s tichým fallbackem na scrollbary,
- obnova původní hodnoty, vzorce, číselného formátu a výplně.

## Import do vývojového XLSM

1. Otevřete nový sešit a uložte jej jako `MR_Helper_DEV.xlsm`.
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
