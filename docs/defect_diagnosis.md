# Defect diagnosis — TFBR monthly close workbooks (Julio 2026)

Companion to `formula_analysis.md`. Source dumps for BAR added this session:
`scratchpad/dump/formulas/BAR_*.tsv` (SALDOS, EESP, EERR, EEPN, Activo, Pasivo, Anexo I, Anexo II, Bienes).

**A methodology note that matters for everything below:** several sheets in the *acumulado* files
(BA and BAR) have a `UsedRange` that does **not** start at row 1 — e.g. BA/BAR's `EESP!UsedRange.Row=5`,
`EERR!UsedRange.Row=4`, and BAR's `Anexo I!UsedRange.Row=3` / `Bienes!UsedRange.Row=3`. (`SALDOS`, `Activo`,
`Pasivo`, `Anexo II`, `EEPN` are `Row=1` in all four files — no offset there.) The bulk-dump tool's array
row *N* then corresponds to Excel row `UsedRange.Row + N − 1`, not row *N*. Reading array-row numbers as
Excel row numbers without checking this **produced a wrong finding in the original `formula_analysis.md`**
— see Section 2, item 6 below, where re-verification via direct single-cell COM reads overturns the
"circular formula" claim for BA's EERR sheet. Every address quoted below was confirmed with a direct
single-cell `Range(addr).Formula`/`.Value2` COM read, not just the bulk array dump, specifically to avoid
repeating that mistake.

---

## Section 1 — BAR (`Balance acumulado TFBR Julio 2026-Reales.xls`)

### SALDOS layout

| Account-code col | Deudor | Acreedor | Saldo (VLOOKUP) | Staging paste range |
|---|---|---|---|---|
| C | D | E | F | `$C$228:$D$315` |

This is a **fourth distinct layout**, different from all three other files (BM=A/B/C/D, BMR=B/C/D/E,
BA=B/C/D/E with staging in A/B). Confirms the "no fixed cell-mapping across files" conclusion from
`formula_analysis.md` — automation must read a per-file layout, full stop.

Pattern is otherwise identical to the other three: `=IF(F##>0,F##,0)` (Deudor), `=IF(F##<0,-F##,0)`
(Acreedor), `=IFERROR(VLOOKUP(C##,$C$228:$D$315,2,FALSE),0)` (Saldo). Staging area (`C228:D315`) is
**still populated** at capture time (mid-cycle, like BA), confirming the paste format: code in C, numeric
balance in D.

### Two things unique to BAR, not seen in BM/BMR/BA

1. **The period-end exchange rate is typed directly into `SALDOS!G1` (`293.17575`)**, labelled by `SALDOS!F1`
   ("T.C. Reales al 31/07/2026"). In BM/BMR/BA the TC rate lives *only* in the separate `Sumas y Saldos`
   source file's `Hoja1!C6` — grepping BMR's and BA's `SALDOS` dumps for "T.C."/"Reales" finds nothing. BAR
   duplicates the rate as a second manual-entry point, which is a reconciliation risk if the two ever
   diverge (e.g., someone re-runs `Sumas y Saldos` with a corrected rate but forgets to update `SALDOS!G1`).
2. **A full "EXPLICACIÓN DIF DE CAMBIO" manual schedule at `SALDOS!D214:D222`** — one typed exchange-difference
   figure per month (`D214`="ENERO 6262.5" … `D220`="JULIO 3231.78"), plus `D221`="ACUMULADO 01/01/2026-30/06/2026"
   = 92982.88, all summed at `D222=SUM(D214:D221)` and cross-checked against `SALDOS!E209/D209` (the
   "DIFERENCIA DE CAMBIO" account row) via `E222=+E209-D222-D209` (=0, ties out currently). This is a genuine
   *new* manual-input point for the automation inventory — not present in BM/BMR/BA — and it must be typed
   once per month for BAR specifically.

### Other manual inputs (same pattern as siblings)

- Staging paste `SALDOS!C228:D315`.
- `EESP!F13`/`G13` — period-end date and prior-year comparison date (plain date serials). **Note:** this is a
  *different real cell* than BA's equivalent (`EESP!F9`/`G9`, confirmed via direct COM read — see the
  correction to the original analysis in the note below) because BAR's `EESP!UsedRange.Row=5` while BA's
  is `Row=1`; the two files' EESP sheets are laid out with a several-row vertical shift relative to each other.
- `EEPN` equity-movement rows (typed contributions/distributions) — same shape as BM/BMR/BA.
- `Bienes` sheet (`Activo`, `Pasivo`, `Patrimonio Neto`, `Impuesto` — a standalone 0.5%-of-PN tax side-calc,
  `=+A6*0.005`) — present in BA and BAR only, titled "Calculo De Bienes Societarios de Transporte Furlong
  Barsil **2009**" (a 17-year-stale title; the formulas themselves are live and reference the current period's
  `EESP`/`Pasivo` totals, so the sheet is functionally fine despite the dead label).

**Correction to the original `formula_analysis.md`:** that document states BA's date entry point is
`EESP!F9`/`G9`, based on a pandas-CSV read. Direct COM verification this session confirms `EESP!F9` and `G9`
are genuinely **blank** in BA, and the real typed dates live at `EESP!F13`/`G13` (value 46234/38717, matching
what every consumer sheet's `=+EESP!F13` formula actually points to). The CSV-based row count in the earlier
session was off by the same kind of index/header mismatch as the EERR issue below. Automation should read
`EESP!F13`/`G13` for BA, not `F9`/`G9`.

### Defects found in BAR (mirroring / extending the BM/BMR/BA defect list)

| Defect | BAR address | Matches |
|---|---|---|
| `#REF!` in a totals row | `Pasivo!K17` = `=SUM(#REF!)`, cached `-2146826265` | Same family as BM `Pasivo!K30`, BMR `Pasivo!K35` |
| `#REF!` in the Anexo II grand-total row | `Anexo II!G84` = `=SUM(#REF!)`, cached `-2146826265` | Same family as BM `G113`, BMR `G108`, BA `G85` |
| `#REF!` — **newly found, not in the original 4-file list** | `EEPN!K17` = `=+#REF!+K13`, cached `-2146826265` | Same broken-reference family, different sheet |
| `EERR` "Impuesto a las Ganancias" not populated | real cell `C27` is **completely blank** — no formula, no value (Excel treats it as 0 in the downstream subtraction, but there's no explicit figure) | BA's equivalent (`C28`) is at least a hardcoded `0`; BAR doesn't even have that |

**Striking cross-file confirmation:** all three `#REF!` cells found in BAR — `Pasivo!K17`, `Anexo II!G84`,
`EEPN!K17` — carry the **exact same stale cached value, `-2146826265`**, as BA's `Anexo II!G85`. Same number,
three different sheets, two different files. That is strong evidence all of these `#REF!`s trace back to one
common broken source range that existed in a shared ancestor template before these workbooks diverged —
not four independent unrelated breakages. See Section 2, item 5 for the fix implication.

### Things checked and found clean / not present in BAR

- `Anexo I`: **no `#REF!`** anywhere (BMR's `D17` defect does not carry over — BAR's Anexo I pulls its
  "Disminuciones"-equivalent figures cleanly, `col2..col11` all resolve without error).
- `EESP!F13`: **clean** — `=+'Anexo I'!K21` (Bienes de uso note reference), not the dead 2004 external link.
  Confirms that defect is BM-only.
- `Anexo II`: **no 2012 external-link block** (`D121:D134`-style). BAR's Anexo II ends cleanly at row 85
  ("Total al 31.12.05" — still a stale label, but a local `SUM`, no external file reference). Confirms BM-only.
- `Anexo II`: **no J/K/L/M cost-center allocation columns at all** — the sheet is only 11 columns wide (A–K,
  and even then only up to column G has content); the hand-typed cost-center block that is "the single
  biggest automation red flag" in BM/BMR simply doesn't exist in the *acumulado* files. See the currency-skip
  spot-check below.

### Spot-check: does the BM/BMR currency-skip pattern (identical unconverted literals in cost-center columns) appear in BA vs BAR?

**No — because the feature itself doesn't exist in the acumulado files.** Both BA's and BMR's... (correction:
both **BA's and BAR's**) Anexo II sheets stop at column G (BA) / G (BAR) with no J–M cost-center columns
present at all. The hand-typed allocation block is a *Balance Mensual*-only feature. There is nothing to
compare for currency consistency in the acumulado pair — this spot-check question doesn't apply to BA/BAR.

---

## Section 2 — Per-cell diagnosis

### 1. BM `EESP!F13` — dead external link to a different company's 2004 file

**Current formula:** `=+'\\Servidor\D\clientes\Empresas\Furlong Equipos y Vehiculos S.A\BALANCES\AÑO 2004\[Balance 2004.xls]Activo'!F21`

**Row/line-item context:** Excel row 13 in BM's EESP is an **unlabeled line sitting between "Caja y bancos
(Nota 4.a)" (row 11/12 area) and "Inversiones temporarias (Nota 4.b)" (row 14)** — it has no "Nota X"
reference of its own. Cross-checked against the pandas value CSV, the row's neighbors are "Creditos
comerciales"/"Creditos Impositivos"-style unlabeled rows. It sits inside the `SUM(F12:F22)` that rolls up
into "Total del Activo corriente" (row 23/24), so whatever this cell resolves to **does currently affect the
reported Activo Corriente total**.

**Cross-file comparison:** Confirmed via direct dump — **BMR has no formula in row 13 at all** (`col6` and
`col12` both empty; the row is skipped entirely in BMR's line-item list). **BA has a blank spacer row at its
own row 13** (nothing there either). **BAR's row 13 is used for something else entirely** — it's the
period-end date cell (`F13`/`G13`), a coincidence of row position, not a related line item. So **no sibling
file has a working formula for "the equivalent line"** — because there is no equivalent line anywhere else.
This row appears to be a vestige of the workbook having originally been cloned from "Furlong Equipos y
Vehiculos S.A."'s own template (explaining why the dead link points at *that* company's balance file) and
never fully genericized for TFBR.

**Proposed fix:** Cannot safely reconstruct what real TFBR figure (if any) this row was ever supposed to
hold — there's no sibling formula to translate and no SALDOS account is obviously implicated. Two honest
options, both requiring a human decision (the "keep layout as-is" constraint rules out deleting the row):
- **Low-risk default:** replace the formula with a static `0`, matching the fact that three of four sibling
  workbooks carry no value here at all. This stops Excel from trying (and failing) to resolve a dead UNC
  link on every open/recalc, and given the row isn't labeled or referenced by any Nota, its removal from the
  total is very unlikely to be noticed.
- **Needs human input:** confirm with the preparer whether row 13 ever represented a real, still-relevant
  balance (in which case someone needs to identify the correct current-period SALDOS source and write a
  proper formula) or whether it's safe to zero out per the above.
**Confidence:** medium-high that `0` is a safe default (three of four files already effectively have no value
there); low confidence on what the "correct" formula would be if the line does need to carry real data —
that part is flagged for the user, not guessed.

---

### 2. BM `Anexo II!D122, D126, D129, D130, D132` — dead 2012 external "RESUMEN PAGOS" links

**Confirmed via direct COM read** (real Excel addresses, not array-dump line numbers):

| Cell | Row label (from adjacent col E) | Formula | Cached value |
|---|---|---|---|
| D121 | "TOTAL GASTOS BALANCE" | local | 102,297,320.33 |
| **D122** | "TOTAL PAGOS" | `=+'...RESUMEN PAGOS MAYO 2012.xlsx'!PRESUPUESTO GASTOS FILIAL'!$C$72` | 3,310,746.86 |
| B123 | "DIFERENCIA" | `=+D121-D122` | 98,986,573.47 |
| **D126** | "TRANSFERENCIAS" (label in E126) | same external file, `$C$85` | 2,583,219.17 |
| **D129** | "ADELANTOS" (label in E129) | same external file, `$C$86` **`-66154.26`** | 78,529.45 |
| **D130** | "TRANSPORTISTAS" (label in E130) | same external file, `$C$88` | 89,824.30 |
| **D132** | "DIFERENCIA EN SUELDOS (NETO VS BRUTO)" (label in E132) | `=+C31+C105-[file]!$D$64-$D$63-$D$62` | 36,354,509.30 |

**What this block is:** a self-contained "Total Pagos" reconciliation schedule — it compares total expenses
booked in the P&L (`D121`, computed locally) against total *payments actually made* (`D122`, sourced from
an external one-off "May 2012 payment summary" file), broken down by category (D126/129/130/132), with
`D123` as the check ("DIFERENCIA").

**Cross-file comparison:** Confirmed BMR and BA and BAR all end their Anexo II sheets around row 85–112
with **no equivalent block at all** — this reconciliation section only exists in BM.

**Does anything downstream depend on it?** Searched all BM formula dumps for references to `Anexo II` rows
121–134 outside the sheet itself — **none found.** `D121:D134` is a dead end: it feeds nothing in EERR,
EESP, or SALDOS. It's purely self-referential decoration.

**Is it still needed?** No evidence it is. The external file is a single **May 2012** payment snapshot — it
has not genuinely updated in over a decade; the "live" link is cosmetic, since the source file is
almost certainly gone (dead UNC path `\\Estudio-files\e\...`) and Excel is just re-displaying whatever value
was last cached in 2012–2013. This is legacy dead weight, not a live reconciliation.

**Proposed fix:** Since nothing downstream reads this block and three of four sibling files never carried
it, the lowest-risk fix is to **freeze the five cells as static values** equal to their currently-cached
numbers (breaking the dead external link without changing what's displayed), rather than deleting the rows
(which would violate the "keep layout as-is" instruction). **Confidence: high** that the block is safe to
detach from the external link and freeze/zero without affecting any other report figure (verified no
downstream dependency). **Flag for the user:** whether to keep the frozen 2012 numbers on display at all
going forward, since they no longer mean anything for the current period — that's a presentation decision,
not a formula-correctness one.

---

### 3. BM `Pasivo!K30`, BMR `Pasivo!K35` — `=SUM(#REF!)`

**Row/line-item context (BM, confirmed via pandas value CSV + COM):** Row 30 is the subtotal for section
"f. Remuneraciones y cs. Sociales" (SUSS a pagar + Sueldos a pagar, rows 28–29). Sibling columns in the same
row: `E30=SUM(E28:E29)` (works, -37,946,731.54) and `G30=SUM(G28:G29)` (works, 938.38). The header row for
this sheet (row 1) carries two stale comparison-year labels: **G column = "31.03.04"**, **K column =
"31.03.03"** — i.e., columns G and K are leftover prior-fiscal-year comparison columns from a template that
predates the current chart of accounts by 20+ years.

**Cross-file comparison:** BMR's `K35` is the analogous subtotal for its own "Cargas Fiscales" section
(rows 32–34: Plan Mis Facilidades / Interes a devengar / SICORE a pagar), with `E35=SUM(E32:E34)` working
and `G35=SUM(G34:G34)` (working, but oddly narrower than E35's range — the sibling columns aren't even
fully consistent with each other, a sign this comparison-column machinery was already half-abandoned before
it broke outright).

**Same defect family, confirmed a third and fourth time this session:** BAR's `Pasivo!K17` (`=SUM(#REF!)`,
cached `-2146826265`) is the identical pattern, and BAR's row-1 header carries the same "31.03.04"/"31.03.03"
labels over columns G/K. All these `#REF!`s cache the exact same stale number as the Anexo II grand-total
`#REF!`s (item 5) — strong evidence they all descend from one common broken reference in a shared ancestor
template.

**Proposed fix:**
- BM `Pasivo!K30` → `=SUM(K28:K29)` (mirrors E30's range exactly). **Confidence: high** — E30 and G30 both
  sum the same two-row block immediately above the total row; K30 almost certainly did too before the
  reference broke.
- BMR `Pasivo!K35` → `=SUM(K32:K34)` (mirrors E35's three-row range) **or** `=SUM(K34:K34)` (mirrors G35's
  narrower one-row range). **Confidence: medium** — the two sibling columns disagree with each other on
  range width, so which one K35 originally mirrored can't be determined with certainty from the surviving
  formulas alone. Recommend `=SUM(K32:K34)` (matching E, the more "complete" column) as the default, but
  flag for the user to confirm against a pre-break backup/printout if one exists.

**Bigger-picture recommendation:** since this whole "K = prior-year comparison" column appears to be dead
across at least BM, BMR, and BAR, and none of the comparison years (`31.03.02/03/04`) have any relevance to
a Julio-2026 close, consider whether the column is worth fixing at all versus just zeroing/hiding it as part
of the "no restructuring" cleanup — a fix that reconstructs a meaningless 20-year-old comparison isn't
obviously better than a fix that acknowledges it's dead. Flagged for the user's judgment.

---

### 4. BMR `Anexo I!D17` — `=+SALDOS!#REF!`

**Row/line-item context:** Row 17 = **"Rodados"** (vehicles), confirmed via the pandas value CSV (row 17
label = "Rodados", the last of the seven asset categories: Software, Hardware, Computadoras, Equipos,
Muebles y Utiles, Instalaciones, Rodados). Column D in BMR's Anexo I is the **"Disminuciones"** (decreases/
disposals) column of the Valor de Origen block. Rows 11–16 (all other categories) have **no formula at all**
in column D — they're plain typed `0`s (manual input, no disposals that period) — `D17` is the *only* row in
that column carrying a formula, and it's the broken one.

**Cross-file comparison:** BM's Anexo I uses column D for a **different concept entirely** — `D##=+B##`
("Valor de Cierre" = Valor de Comienzo, i.e., BM's D means closing value, not disposals) — so there is no
direct "copy BM's formula" fix; the column letter doesn't mean the same thing across files (confirmed
structural difference, matches formula_analysis.md's note that BM's Anexo I is "more manual/less
SALDOS-linked" than BMR's). BMR's own row 17 already has a working, unbroken SALDOS-linked formula in
column I (`I17=-SALDOS!E48`, Depreciación Acumulada al Inicio) — so the sheet's Rodados row isn't uniformly
broken, just this one cell.

**Proposed fix:** Since every sibling category in the same column (D11:D16) is a plain typed `0`, and
disposals are evidently rare enough that BMR's preparer defaults to `0` for six of seven categories, the
lowest-risk fix is to replace `=+SALDOS!#REF!` with a static `0`, matching the established convention for
this column. **Confidence: medium-high** that `0` is safe *if no vehicle was actually disposed of in the
period the formula last worked correctly* — but this cannot be verified from the file alone. **Flag for the
user:** confirm with the preparer whether Rodados ever had a real disposal figure here; if so, the correct
fix is a live link to whatever SALDOS row tracks vehicle disposals (analogous to how row 17's own `I17`
already links to `SALDOS!E48` for the accumulated-depreciation side), not a hardcoded `0`.

---

### 5. `Anexo II!G113` (BM) / `G108` (BMR) / `G85` (BA) / `G84` (BAR) — `=SUM(#REF!)` grand-total row

**Row/line-item context (confirmed via direct COM read on BM):** Row 113 (BM) is the sheet's grand total —
`=+EEPN!B18` labels it, and the sibling columns all work: `C113=SUM(C9:C112)` (102,297,320.33),
`D113=SUM(D11:D112)` (6,217,669.69), `E113=SUM(E9:E112)`, `F113=SUM(F11:F112)`. **`G113=SUM(#REF!)`**,
cached value **`-2,146,826,265`**. The header row (row 7 in BM) confirms column G is labelled
**"TOTALES AL 31.03.02"** — yet another instance of the same stale prior-fiscal-year comparison column found
in the Pasivo sheet (item 3), just with a different vintage year (`31.03.02` here vs `31.03.03/04` in
Pasivo) — consistent with these being remnants of the same old multi-year-comparison template feature,
abandoned at different times across different sheets.

**Cross-file confirmation — same defect, same cell shape, same stale number, in all four files:**
- BMR `G108`: `=SUM(#REF!)`, sibling columns `C108=SUM(C9:C107)`, `D108=SUM(D9:D107)`, `E108=SUM(E9:E107)`,
  `F108=SUM(F11:F107)` all work.
- BA `G85`: `=SUM(#REF!)`, cached **`-2146826265`** — exact same stale number as BM's (confirmed via direct
  COM read this session; BA's own dump already had this correctly).
- BAR `G84`: `=SUM(#REF!)`, cached **`-2146826265`** — same number again.

All four `G`-column grand totals broke the same way and (where readable) cache the identical stale number,
which is strong, repeated evidence of one shared broken source — most likely this column, across every sheet
in every file, once pointed at a "prior period"/comparison SALDOS block that was deleted from the shared
ancestor template at some point, and every descendant file has carried the resulting dangling reference
forward untouched ever since (the earliest surviving comparison-year label found, `31.03.02`, implies this
dates back to at least fiscal year 2002).

**Proposed fix:** Mirror the working sibling columns' range shape for each file:
- BM `G113` → `=SUM(G11:G112)` (matching D/F's start row) or `=SUM(G9:G112)` (matching C/E's start row) —
  the sheet's own C/D/E/F columns disagree on start row (9 vs 11) so this can't be pinned down exactly.
  **Confidence: medium** on the range, **high** on the general shape (a `SUM` over the same row block as
  its neighbors).
- BMR `G108` → `=SUM(G9:G107)` or `=SUM(G11:G107)`, same caveat. **Confidence: medium.**
- BA `G85` → `=SUM(G11:G84)`. **Confidence: medium.**
- BAR `G84` → `=SUM(G11:G83)`. **Confidence: medium.**

**Important caveat, high confidence:** because this is the same dead "prior-year comparison" column
identified in item 3, and there is no surviving G-column *data* anywhere in any of these sheets to sum in
the first place (the underlying account-level G cells for individual expense rows are themselves blank/
`0` in every row checked — the column has no live inputs, only the broken total formula survives), **simply
restoring a `SUM()` formula will just produce `0`, not a meaningful "prior-period total."** The honest fix is
either (a) replace with a hardcoded `0` / blank and stop presenting this column as if it's live, or
(b) remove the column from the report entirely as part of a future template cleanup (out of scope for
"fix, don't restructure," but worth flagging). **Flagged for the user's decision** on which of these two
paths to take — this is a presentation call, not something a formula fix alone resolves.

**Newly found in BAR this session (not in the original 4-file list):** `EEPN!K17` = `=+#REF!+K13`, same
cached `-2,146,826,265`. Row 17 in BAR's EEPN is the "Total al 31.7.2026" patrimonio-neto rollforward total;
column J (`=SUM(J11:J16)`, -1,507,998.50) is the working "Total del patrimonio neto" figure actually used
elsewhere in the workbook (matches `EESP!L23`). Column K appears to be a duplicate/legacy "Total" column
running in parallel to J, broken the same way. **Proposed fix:** replace with `=+J17` (mirror the working
J-column total) if K is meant to just restate J, or `0`/blank if K is genuinely vestigial. **Confidence:
low** — insufficient evidence to know K's original intended purpose distinct from J; flagged for the user.

---

### 6. BA `EERR!C18`, `C24`, `C27` — claimed self-referential circular formulas

**This finding from the original `formula_analysis.md` is incorrect, and the correction is the most
important result of this diagnosis pass.**

The original analysis read the EERR sheet's bulk-dump array assuming array-row *N* = Excel row *N*. But
`EERR!UsedRange.Row = 4` in BA (confirmed via COM this session), so array-row *N* is actually **Excel row
N+3**. The formula text at array-row 18 (`=+C15-C17-C18-C19`) is real and was transcribed correctly — but
it does **not** live in cell C18. Direct single-cell COM verification:

| Real cell | Label | Formula | Value |
|---|---|---|---|
| C21 | "Subtotal" | `=+C15-C17-C18-C19` | -580,510,605.33 |
| C27 | "Ganancia antes de impuesto" | `=+C21+C22-C24+C23+C25+C26` | -574,495,343 |
| C28 | "Impuesto a las Ganancias" | *(bare hardcoded)* `0` | 0 |
| C30 | "Ganancia (Pérdida) del ejercicio" | `=+C27-C28` | -574,495,343 |

None of these formulas reference their own cell. `C21`'s formula references `C15` (Ingresos), `C17`/`C18`/
`C19` (the three expense-by-function subtotals) — none of which is `C21`. `C27`'s formula references `C21`
through `C26` — none of which is `C27`. `C30`'s formula references `C27` and `C28` — neither is `C30`. **This
is a completely ordinary, non-circular income-statement roll-up**, structurally identical to BM's and BMR's
EERR sheets (verified: BM's equivalent chain is `C19=+C12+C13-C15-C16-C17` → `C25=+C19+C20+C22+C21+C23+C24`
→ `C28=+C25-C26`, the exact same shape, also non-circular).

**BAR's equivalent (also re-checked this session, also mis-flagged as circular in my own first pass before
I caught the offset issue):**

| Real cell | Label | Formula | Value |
|---|---|---|---|
| C21 | "Subtotal" | `=+C15-C17-C18-C19` | -2,088,893.68 |
| C26 | "Ganancia antes de impuesto" | `=+C21+C22+C23+C25+C24` | -1,959,559.89 |
| C27 | "Impuesto a las Ganancias" | *(blank — no formula, no value)* | — |
| C29 | "Ganancia (Pérdida) del ejercicio" | `=+C26-C27` | -1,959,559.89 |

Same conclusion: no self-reference anywhere. **The only real issue in either file's EERR sheet is that the
income-tax line is a placeholder** — BA has it hardcoded to `0` (`C28`), BAR has it **completely empty**
(`C27` — worse than BA's, since a blank cell gives no visual signal that it's an intentional zero vs.
simply forgotten). This matches the manual-steps inventory item already flagged in `formula_analysis.md`
("Type/verify the income-tax line each period ... confirm whether this needs a real monthly value or is
intentionally left at 0") — it's a manual-input gap, not a formula defect, and needs no repair beyond
(optionally) typing an explicit `0` into BAR's `C27` for clarity/consistency with BA.

**Confidence: high** (verified via direct single-cell `.Formula`/`.Value2` COM reads on real addresses, not
inferred from array-dump positions) that there is **no circular-formula defect to fix** in either BA's or
BAR's EERR sheet. Recommend striking this item from the fix list entirely; the cached net-income figures on
this sheet can be trusted (subject to the normal caveat that `Impuesto a las Ganancias` is a manually-typed
placeholder, same as in BM/BMR).

---

## Summary of proposed fixes

| # | Cell(s) | Fix | Confidence |
|---|---|---|---|
| 1 | BM `EESP!F13` | Replace dead external link with `0` (no sibling file has an equivalent value) | Medium-high on `0`; low on "correct" value — needs preparer input |
| 2 | BM `Anexo II!D122/D126/D129/D130/D132` | Freeze as static values (dead 2012 link, no downstream dependents, block absent in BMR/BA/BAR) | High that it's safe to detach; presentation decision flagged |
| 3 | BM `Pasivo!K30` | `=SUM(K28:K29)` | High |
| 3 | BMR `Pasivo!K35` | `=SUM(K32:K34)` (or `K34:K34` — sibling columns disagree) | Medium |
| 3 | BAR `Pasivo!K17` | Same family — same treatment as BM/BMR once a decision is made on whether to fix or retire the dead "prior-year comparison" column | Medium (see caveat) |
| 4 | BMR `Anexo I!D17` | `0`, matching sibling categories' convention, unless preparer confirms a real Rodados disposal | Medium-high |
| 5 | BM `G113` / BMR `G108` / BA `G85` / BAR `G84` / BAR `EEPN!K17` | Column has no live underlying data to sum — recommend `0`/blank rather than a reconstructed `SUM()`, and flag to the user whether to retire the column | High that reconstructing a "real" total isn't possible; medium on exact formula if kept |
| 6 | BA `EERR!C18/C24/C27` (real cells C21/C27/C30) | **No fix needed — not actually broken.** Original "circular formula" finding retracted after direct verification. | High |
| 6b | BA `EERR!C28`, BAR `EERR!C27` | Cosmetic only: BAR's blank tax cell could be typed to `0` for consistency with BA | Low priority |

Items 3 and 5 share a root cause (a dead "prior fiscal year comparison" column/feature, vintage ~2002–2004,
present across Pasivo and Anexo II in at least three of the four files, evidenced by matching stale labels
and — where cached — an identical stale number, `-2,146,826,265`, in every occurrence). Worth deciding as one
policy call (retire the column vs. reconstruct a meaningless comparison) rather than fixing each cell in
isolation.
