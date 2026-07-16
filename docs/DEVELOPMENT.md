# Vývoj

Doporučený vývoj probíhá v samostatném `.xlsm` sešitu. Zdrojové soubory exportujte po každé změně zpět do tohoto repozitáře.

## Doporučené pořadí importu

1. `modGlobals.bas`
2. `modMouseWheel.bas`
3. `modValueHistory.bas`
4. `CChangeRecord.cls`
5. `modPrefillEngine.bas`
6. `frmPrefill.frm`
7. `frmPhraseManager.frm`
8. `frmReview.frm`

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
- import UTF-8 TXT s prázdnými řádky a duplicitami a export pouze uživatelských frází,
- návrat z kontroly zpět do formuláře a bezpečné vrácení aplikovaných změn,
- zavření prázdné kontroly křížkem a potvrzení vrácení neuložených změn,
- automatický centrovaný přechod na vybranou změnu a cyklická navigace tlačítky nahoru/dolů,
- rolování seznamů a oblasti frází kolečkem myši; při nepodporovaném API zůstávají funkční scrollbary.
