# Formula analysis — TFBR monthly close workbooks (Julio 2026)

Source dumps: `scratchpad/dump/formulas/*.tsv` (raw, COM `.Formula` + broken `.Text`) and
`scratchpad/dump/formulas/compact/*_compact.txt` (trimmed to populated cells, `[CellAddr]formula||VAL:value`).
Cross-checked against the reliable value-only CSVs in `scratchpad/dump/*.csv` (pandas/xlrd) where the
COM `.Text` bug left `VAL:` blank. File prefixes: **BM** = Balance Mensual (ARS), **BMR** = Balance Mensual
-Reales (R$), **BA** = Balance acumulado (ARS, spot-check only — dumped *after* the `.Text`→`.Value2` fix,
so its `VAL:` figures are reliable).

**Headline structural finding:** the three files use three *different* SALDOS column layouts. This alone
means any automation cannot assume one fixed cell-mapping across files — it must read a per-file layout
(or better, standardize the template first):

| File | Account-code col | Deudor | Acreedor | Saldo (VLOOKUP) | Staging paste range |
|---|---|---|---|---|---|
| BM (Balance Mensual, ARS) | A | B | C | D | `$A$243:$B$360` |
| BMR (Balance Mensual, R$) | B | C | D | E | `$B$232:$C$289` |
| BA (Balance acumulado, ARS) | B | C | D | E | `$A$235:$B$357` |

The BA staging range uses column A (not B like BMR) even though its Deudor/Acreedor/Saldo columns are
shifted like BMR's — a third distinct hybrid layout. Every SALDOS-driven downstream formula (in Activo,
Pasivo, EERR, Anexo II) also differs correspondingly between BM and BMR (e.g. BM uses `SALDOS!B/C`,
BMR uses `SALDOS!C/D`, `SALDOS!D/E`, etc., depending on the row).

---

## SALDOS (all three files)

**What it computes:** the full fixed chart-of-accounts template (≈230 account rows) that turns a flat
account-code/balance paste into a Deudor/Acreedor/Saldo table, with section subtotals (Activo Corriente,
Activo No Corriente, Pasivo, etc.) and a grand-total check row.

**Sources / pattern:**
- Every account row: `=IF(<Saldo-col><row>&gt;0,<Saldo-col><row>,0)` (Deudor), `=IF(<Saldo-col><row>&lt;0,-<Saldo-col><row>,0)` (Acreedor), and `=IFERROR(VLOOKUP(<code-col><row>,<staging-range>,2,FALSE),0)` (Saldo) — the account code itself (col A in BM, col B in BMR/BA) is a **plain typed/static value**, not a formula (part of the fixed template, doesn't change monthly).
- Subtotal rows use `SUM()` over row ranges of the same sheet (e.g. BM `D50=SUM(D7:D49)`, `E50=SUM(B7:B49)`, `F50=SUM(C7:C49)`, `G50=+E50-F50-D50` as a balance check).
- Grand total / check rows: BM `D236=SUM(B7:B235)`… `D237=+B236-C236` (debit=credit check), `D241=SUM(B242:B360)` (staging-paste checksum). BMR equivalent at row 226/227: `F227=+C227-D227`.
- **Cross-sheet feed-back:** BM `D238=+D236+EERR!C28` — SALDOS pulls the P&L's net result (EERR) back into the balance-sheet total, closing the accounting loop.

**Manual/free-typed inputs each period:**
- The staging paste area itself: `A243:B360` (BM), `B232:C289` (BMR), `A235:B357` (BA) — account code + balance pairs pasted from the Sumas y Saldos export, then (per BM/BMR's current on-disk state) cleared after use. **BA's staging area is still populated** (not yet cleared when captured), which let us directly confirm its shape: column A = `"<code>  <description>"` text, column B = numeric balance (e.g. `A307: "4221000000  PEAJES"`, `B307: 671541.29`).
- Account codes in the template column (A in BM, B in BMR/BA) are static/typed but effectively fixed — only touched if the chart of accounts changes.

**Cross-file / external links:** none found in SALDOS itself.

---

## EESP (Estado de Situación Patrimonial)

**What it computes:** the balance sheet — pulls Activo/Pasivo subtotals and equity from the note schedules and EEPN, with a check that Total Activo = Total Pasivo + PN.

**Sources:**
- `=+Activo!F##` / `=+Pasivo!E##` for each line (note-schedule references).
- `=+EEPN!I18` (BM) / `=+EEPN!J18` (BMR) for Patrimonio Neto.
- `=+SALDOS!E47` (BM F29) vs `=+'Anexo I'!K18` (BMR F29) — **the same conceptual line pulls from a different source sheet between the two files.**
- BA: `F24=+'Anexo I'!K21` for Bienes de Uso.

**Manual/free-typed inputs:**
- `F9` and `G9` (confirmed via BA's readable values: `F9=46234`, `G9=38717`, both plain typed Excel date serials, no formula) — **current period-end date and prior-year comparison date**, typed once per period; every other sheet's date header (`EERR!C9/D9`, `Anexo I`, `Anexo II`, etc.) reads these via `=+EESP!F9`/`=+EESP!F13` chains, so this is a single point of manual date entry that fans out.
- Signature block text (`E35/K35`, `E36/K36`, `E37/K37` in BA) — accountant/company-representative names and titles, static, rarely changes.

**Hardcoded numbers (red flags):**
- **BMR `L33`: `=+L25+L31+0.01`** — a hardcoded `+0.01` rounding plug baked into the PN total formula.
- **BA `F28`: `=+F23+F30-0.02`** — a hardcoded `-0.02` rounding plug in Total Activo.

**Cross-file links (critical):**
- **BM `F13`: `=+'\\Servidor\D\clientes\Empresas\Furlong Equipos y Vehiculos S.A\BALANCES\AÑO 2004\[Balance 2004.xls]Activo'!F21`** — a live external-workbook reference to a **different company** ("Furlong Equipos y Vehiculos S.A.") and a **2004** file, on a UNC path (`\\Servidor\D\...`) that almost certainly no longer exists. This is only in the ARS Mensual file — BMR's `F13` has no formula at all (row is simply skipped in BMR's line-item list). This is the single highest-priority broken link found in the whole review.

---

## EERR (Estado de Resultados)

**What it computes:** the P&L — revenue, cost/expense subtotals sourced from Anexo II by function, financial results, and net income, which feeds back into EEPN and SALDOS.

**Sources:**
- `=+SALDOS!C###`/`D###` combinations for revenue and financial-result lines (column letters differ BM vs BMR per the layout table above).
- `=+'Anexo II'!D###/E###/F###` for the three expense-by-function subtotals (administración/comercialización/financieros).
- `C28 = +C25-C26` is EERR's final net-result cell, which is the value fed back into SALDOS!D238 (BM) and into EEPN (`I16=EERR!C28`).

**Manual/free-typed inputs:** none directly on this sheet in BM/BMR — it is fully formula-driven from SALDOS/Anexo II. **BA is the exception (see below).**

**Critical integrity issue — BA (acumulado) only:**
- `C18 = +C15-C17-C18-C19` — **self-referential** (the formula for C18 includes C18 itself).
- `C24 = +C21+C22-C24+C23+C25+C26` — also self-referential (includes C24).
- `C27 = +C27-C28` — also self-referential (includes C27).
- `C25` ("Impuesto a las Ganancias" / income tax) is a **bare hardcoded `0`**, not a formula.

These self-references almost certainly stem from a row insert/delete that shifted relative references without correcting them (classic Excel corruption pattern), or the workbook relies on iterative calculation being enabled to "resolve" the circularity with a stale/arbitrary value. The cached `VAL:` figures shown (`-580510605.33`, `-574495343`, etc.) cannot be trusted as a correct recalculation — **this sheet's logic needs to be verified against source figures before any automation trusts it**, and should not simply be copied forward as "the formula" — it needs to be fixed at the source, not reproduced.

---

## EEPN (Estado de Evolución del Patrimonio Neto)

**What it computes:** equity rollforward — opening balances plus period movements (capital contributions, distributions, results) reconciled to closing PN, which EESP and Activo's rollforward reference.

**Sources:**
- `I11 = -SALDOS!D126` (BM) / `-SALDOS!E119` (BMR) — opening-equity link to SALDOS.
- `I16 = EERR!C28` (net income for the period, feeds the rollforward).
- Row 13/18 are `SUM()` roll-up formulas.

**Manual/free-typed inputs — this is the sheet's core manual-entry point:**
- **`C11:H11`** (aportes/movements columns feeding `E11=+C11+D11`, `J11=+E11+F11+G11+I11`) — referenced by SUM formulas but **no formula is present in C11, D11, F11, G11, H11 themselves** — these are typed capital-movement figures entered each period (contributions, distributions, revaluations, etc.).
- **`C19:H19`** (a second movement row, `E19=+C19+D19`) — same pattern, likely a second category of equity movement, also manually typed.

---

## Activo / Pasivo (note schedules)

**What they compute:** the detailed line-item breakdown behind EESP's Activo/Pasivo totals, built almost entirely from SALDOS arithmetic (one formula per disclosed line, summed into note subtotals that EESP references).

**Sources:** dense, uniform pattern — `E## = ±SALDOS!C##±SALDOS!B##` (BM) or `±SALDOS!D##±SALDOS!C##` (BMR), one line per account grouping, rolled up via `SUM()`. No sheet-level narrative/manual figures were found in the formula layer.

**Hardcoded numbers (red flags):**
- **BM Pasivo `K30`: `=SUM(#REF!)`** — broken/deleted-reference formula.
- **BMR Pasivo `K35`: `=SUM(#REF!)`** — same break, present in both currency files (so it's a shared structural defect, not currency-specific).
- **BMR Pasivo `E8`: `=-SALDOS!E58-SALDOS!E59-...-SALDOS!E111-0.01`** — a hardcoded `-0.01` rounding plug appended to an otherwise-clean aggregation formula.

**Cross-file links:** none.

**Note on column E/F:** SUM formulas reference an adjacent `F` (or `G`) range (e.g. `E19=SUM(F10:F18)`) whose own cells never appear as formulas in the COM dump — most likely these are merged E:F cells where COM only reports the anchor cell's formula (a known Excel/COM limitation for merged ranges), not evidence of separate hardcoded overrides. Flagged for visual/manual confirmation in Excel rather than as a confirmed defect.

---

## Anexo I (Bienes de Uso / fixed-asset rollforward)

**What it computes:** a small fixed-asset continuity schedule (opening balance, additions, depreciation, disposals, closing balance) per asset category, feeding EESP's "Bienes de uso" line.

**Sources:** BM: `D##=+B##`, `I##=+J##-E##`, `K##=+D##-J##` — mostly intra-sheet arithmetic. BMR: partially different — `I13=-SALDOS!E47`, `I15=-SALDOS!E42`, `I17=-SALDOS!E48` pull directly from SALDOS, a structure BM's version does **not** have (BM's Anexo I is more manual/less SALDOS-linked than BMR's for the same conceptual sheet).

**Manual/free-typed inputs:**
- **BM: `B11:B17`, `C11:C17`, `E11:E17`, `F11:F17`, `J11:J17`** — no formulas appear in these cells; they are the period's typed asset additions/disposals/depreciation figures per category, each month.

**Cross-file links / broken references:**
- **BMR `D17`: `=+SALDOS!#REF!`** — a broken reference, distinct from (but analogous to) the Pasivo `#REF!` issues. Only in the Reales file; BM's Anexo I row 17 has no such formula.

---

## Anexo II (expense breakdown by function)

**What it computes:** the expense-by-function/cost-center schedule behind EERR's three expense lines (administración/comercialización/financieros) — by far the sheet with the heaviest manual load in the whole workbook.

**Sources:** each expense row: `C## = SUM(D##:F##)` with `D##/E##/F##` individually pulling `=±SALDOS!B##` (BM) or `=±SALDOS!C##` (BMR/BA), one per account. Totals row (`113` in BM, `108` in BMR, `85` in BA) sums the whole block per column.

**Hardcoded numbers — extensive, and this is the single biggest automation red flag in the workbook.**
Dozens of cells in columns **J, K, L, M** (cost-center/allocation columns) contain **typed arithmetic of raw invoice/receipt amounts**, not links to any source data, e.g. (BM row numbers; BMR carries the *identical* figures at slightly different rows):

- `K40 = 4800+5000+3000+5000`
- `K41 = 3211.01+4543.6`
- `J42 = 2248.27+2043.98+1241.35`, `K42 = 2543.13+2338.55+2410.41`, `L42 = 1803.75+2112.36+2202.79`
- `K45 = 3200+3115`, `L45 = 4800+1400+3115`, `M45 = 29395.55+31000.65`
- `J70 = 9400+850+550`, `K70 = 850+700`
- `J74 = +$C$18/3+C74`, `K74 = +$C$18/3`, `L74 = +$C$18/3` (a naive even three-way split of `C18`)
- `J79 = 213.333333333333+500`, `K79 = 213.333333333333+500`, `L79 = 213.333333333333+17.55`
- `L83 = 595-250`, `K93 = 723.69-546.19`, `K97 = 37+14+36.5+E84+88`, `L97 = 302.7-37+63`, `K99 = 528.01-1.9`
- `J105 = +E105*27%`, `K105 = +E105*39%`, `L105 = +E105*34%` (hardcoded allocation percentages, sum to 100%)

**Important cross-file finding:** the BMR (R$) version's equivalent cells (`K38`, `K39`, `J40/K40/L40`, `K43/L43/M43`, `J67/K67`, `J71/K71/L71`, `J76/K76/L76`, `K89`, `K92/L92`, `K94`, `J100/K100/L100`) carry **the exact same numeric literals as the ARS file**, unconverted. Either this allocation block is genuinely currency-independent by design (e.g., it represents fixed cost-sharing ratios, not amounts), or — more likely given the mix of raw peso-looking amounts and percentages — the currency-conversion step is simply being skipped/forgotten for this block each month. **This should be confirmed with the preparer before automating**, not silently replicated as correct in both currencies.

- **`G113` (BM) / `G108` (BMR) / `G85` (BA): `=SUM(#REF!)`** — broken reference in the totals row, present in *all three* files. BA's cached value for this cell is `-2146826265` — proof that Excel is still displaying a **stale pre-break cached number** for a formula that would error (`#REF!`) if recalculated today. Do not trust this total in any file.

**Cross-file links (critical, BM only):**
- **`D122`: `=+'\\Estudio-files\e\Users\LOANA\Documents\FURLONG BRASIL\2012\[RESUMEN PAGOS MAYO 2012.xlsx]PRESUPUESTO GASTOS FILIAL'!$C$72`**
- **`D126`**: same external workbook, `$C$85`
- **`D129`: `=+'...[RESUMEN PAGOS MAYO 2012.xlsx]PRESUPUESTO GASTOS FILIAL'!$C$86-66154.26`** — external link *plus* a hardcoded `-66154.26` adjustment.
- **`D130`**: same external workbook, `$C$88`
- **`D132`**: same external workbook, three separate cell references (`$D$64`, `$D$63`, `$D$62`), combined with local `C31+C105`.

All five reference a 2012 file ("RESUMEN PAGOS MAYO 2012.xlsx" — May 2012 payment summary) on a UNC path (`\\Estudio-files\e\...`) — a **14-year-old external workbook** almost certainly unreachable/gone, embedded in a 2026 report. This entire `D121:D134` block (labelled "Total al 31.12.05" for one of the check rows — another stale 2005 label) appears to be a legacy reconciliation section that BMR and BA do **not** carry at all (their Anexo II ends around row 108–112 with no equivalent block) — likely a vestige nobody has cleaned out of the ARS Mensual file specifically.

---

## Does SALDOS differ between ARS and -Reales, confirming currency translation?

Yes — confirmed directly. BMR's SALDOS pulls its VLOOKUP staging from columns shifted one to the right
of BM's (`$B$232:$C$289` vs `$A$243:$B$360`), and every downstream consumer (Activo, Pasivo, EERR, Anexo II)
correspondingly reads `SALDOS!C##/D##` (BMR) instead of `SALDOS!B##/C##` (BM) for the equivalent line. This
is exactly consistent with the "R$" columns of the Sumas y Saldos export (`Saldo (R$)`, recalculated at the
manually-typed "TC CIERRE" rate — see below) being pasted into BMR's staging area, versus the ARS "Saldo ($)"
columns being pasted into BM's. It is **not** simply a relabeled copy of the same figures — the actual pasted
source data differs by file, confirming -Reales is a true currency translation, not a display-only variant.

---

## Sumas y Saldos source files (context for the manual paste step)

`Sumas y Saldos mensual TFBR 07-2026.xls`, sheet `Hoja1`, confirms the manually-typed exchange rate:

```
row5: [blank] [blank] "TC CIERRE"      [blank]
row6: "Balance de Sumas y Saldos" [blank] 293.1757   [blank]
```

i.e. **`Hoja1!C6` = 293.1757**, a plain typed numeric constant (no formula), labelled by `Hoja1!C5` = "TC CIERRE".
This rate is what `Hoja1` uses to recompute `Saldo (R$)` from `Saldo ($)` and derive the exchange-difference
("DIF CAMBIO") figure described in the task background — i.e., this is the upstream manual step that
ultimately determines what numbers get pasted into the -Reales Balance file's SALDOS staging area.

---

## Manual steps inventory (per period)

1. **Run the Bejerman trial-balance export** → produces `Sumas y Saldos mensual TFBR MM-YYYY.xls` (Sheet1 raw + Hoja1 recalculated) and `Sumas y Saldos acumulado TFBR MM-YYYY.xls`.
2. **Type the period-end exchange rate ("TC CIERRE")** into `Sumas y Saldos mensual ....xls` → `Hoja1!C6` (plain numeric constant). Repeat for the acumulado Sumas y Saldos file's own `Hoja1!C6` (not directly dumped this session, but same sheet/cell pattern expected — confirm before automating).
3. **Copy the resulting account-code/balance pairs and paste them into each Balance file's SALDOS staging table**, then clear the staging area after the workbook has recalculated and been saved/printed:
   - `Balance Mensual TFBR Julio 2026.xls` → `SALDOS!A243:B360` (confirmed empty on disk post-use).
   - `Balance Mensual TFBR Julio 2026 -Reales.xls` → `SALDOS!B232:C289` (confirmed empty on disk post-use).
   - `Balance acumulado TFBR Julio 2026.xls` → `SALDOS!A235:B357` (confirmed **still populated** — captured mid-cycle, giving direct proof of the paste format: `"<code>  <description>"` in col A, numeric balance in col B).
   - `Balance acumulado TFBR Julio 2026-Reales.xls` → presumed analogous to BMR's `B232:C289`-style offset; not directly dumped — verify before automating.
4. **Type the period-end date and prior-year comparison date** into `EESP!F9` and `EESP!G9` of each Balance file (plain Excel date serials, no formula) — this single entry point fans out via formula chains to EERR, Anexo I, Anexo II, and EEPN headers in the same workbook.
5. **Type the period's equity/capital movement figures** into `EEPN!C11:H11` and `EEPN!C19:H19` (contributions, distributions, and a second movement category) — no formula present, must be sourced from board/legal records each period.
6. **Type the period's fixed-asset movements** (additions, depreciation, disposals) into Anexo I's detail columns — in the BM layout: `B11:B17`, `C11:C17`, `E11:E17`, `F11:F17`, `J11:J17`.
7. **Type/re-derive dozens of expense-allocation figures** in Anexo II columns J–M (roughly rows 40–105 in BM/rows 38–100 in BMR) — these are hand-typed sums of underlying invoices/receipts split across cost centers, not linked to any source system. This is the most labor-intensive and error-prone manual step found (and the one where ARS/R$ consistency should be double-checked — see finding above).
8. **Type ad hoc rounding/balancing plugs directly into formula text** whenever the books don't tie to the cent: `EESP!L33` (+0.01, BMR), `Pasivo!E8` (-0.01, BMR), `EESP!F28` (-0.02, BA), `Anexo II!D129` (-66154.26, BM, bundled with the stale 2012 external link).
9. **Type/verify the income-tax line each period**: `EERR!C25` in the acumulado (BA) file is a bare hardcoded `0`, not calculated — confirm whether this needs a real monthly value or is intentionally left at 0 pending year-end.
10. *(Low priority, rarely changes)* Signature block names/titles on EESP/EERR/Anexo I/II (`Miguel Ángel Alonso`, `Tomas Furlong`, etc.) are typed static text, only touched when signers change.

### Separate from the manual-input list — pre-existing defects to resolve before automating (not to replicate)

- **Broken external links**: `EESP!F13` (BM only — 2004 file, different company); `Anexo II!D122/D126/D129/D130/D132` (BM only — 2012 file). Both point to UNC paths that are almost certainly gone.
- **`#REF!` formulas**: `Pasivo!K30` (BM), `Pasivo!K35` (BMR), `Anexo II!G113` (BM) / `G108` (BMR) / `G85` (BA), `Anexo I!D17` (BMR). All display stale cached numbers, not live recalculated values.
- **Self-referential (circular) formulas**: `EERR!C18`, `C24`, `C27` in the acumulado (BA) file only — needs source-side investigation; do not trust the cached net-income figure from this sheet without independent verification.
