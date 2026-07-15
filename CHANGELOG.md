# Changelog

## 1.4.0-dev

- Kompletní zdrojový projekt bez instalačního EXE.
- Uživatelský termín „názvosloví“ nahrazen termínem „slovní spojení“.
- Poslední hodnota se automaticky nevkládá.
- Historie uchovává až 25 unikátních hodnot pro každé slovní spojení.
- Editovatelný ComboBox nabízí doplnění dříve zadaných hodnot.
- Přidána správa slovních spojení, výběr listů, počty výskytů a vrácení všech změn.
- Zachována ochrana proti duplicitnímu zásahu a obnova hodnot, vzorců a formátů.

### Sestavení

- Přidán `tools/build_xlam.ps1` pro vytvoření XLAM.
- Přidán `Build-Installer.ps1` pro vytvoření XLAM, EXE a SHA-256.
- Přidán `installer.iss` pro Inno Setup 7.
- Vygenerované složky `Build/` a `Output/` jsou ignorovány Gitem.

### Oprava sestavení formulářů

- Formuláře se již neimportují přímo přes `VBComponents.Import(.frm)`.
- Build vytváří UserForm komponenty přes VBIDE, přidává ovládací prvky a z `.frm` načítá pouze VBA kód.
- Opravená kompatibilita s chybou `The form class contained ... is not supported in VBE`.
