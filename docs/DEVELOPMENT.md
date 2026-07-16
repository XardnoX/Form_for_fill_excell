# Vývoj

Doporučený vývoj probíhá v samostatném `.xlsm` sešitu. Zdrojové soubory exportujte po každé změně zpět do tohoto repozitáře.

## Doporučené pořadí importu

1. `modGlobals.bas`
2. `modValueHistory.bas`
3. `CChangeRecord.cls`
4. `modPrefillEngine.bas`
5. `frmPrefill.frm`
6. `frmPhraseManager.frm`
7. `frmReview.frm`

## Testovací scénáře

- žádné slovní spojení v sešitu,
- více výskytů stejného spojení,
- skryté a zamčené listy,
- filtrování listů podle názvu se zachování výběru a hromadný výběr vše/žádný,
- sloučené zdrojové i cílové buňky,
- dva různé výskyty mířící do stejné buňky,
- původní vzorec v cílové buňce,
- číslo, datum, text a logická hodnota,
- vrácení jedné změny a vrácení všech změn,
- napovídání dříve použitých hodnot bez automatického vložení a zadání nové hodnoty,
- návrat z kontroly zpět do formuláře a bezpečné vrácení aplikovaných změn.
