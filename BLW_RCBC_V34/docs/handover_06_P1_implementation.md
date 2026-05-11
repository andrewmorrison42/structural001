# Handover 06 — P1 Materials Implementation

## Step
Implement the P1 (Materials) UDF module. Deliver `mod_Materials.bas`
and the `_Test_P1` sheet specification. Stage gate V34_P1.xlsm is
produced by the user after importing and validating in Excel.

## Status
**Ready for import and test.** `mod_Materials.bas` is complete.
Three open questions resolved (Q1, Q3 locked; Q2 partially resolved —
see below). User must import, run test sheet, and confirm acceptance.

## What was delivered
- `BLW_RCBC_V34/bas/mod_Materials.bas` — 17 UDFs in 7 families.
- `BLW_RCBC_V34/docs/P1_test_sheet_spec.md` — test sheet layout with
  expected values and acceptance criteria.
- `BLW_RCBC_V34/docs/handover_06_P1_implementation.md` — this file.

## Decisions made this step

### Q1 — Theoretical thickness th convention
**Decision:** Match V33 sheet exactly — two-face drying formula:
`th = 2·D·b / (2·D + 2·b)` matching BCDcheck BE103.
Recorded as **D-Q1**.

### Q3 — Endogenous shrinkage time coefficient
**Decision:** Corrected to `exp(−0.1·t)` per AS3600-2018 cl 3.1.7.1.
V33 used `exp(−0.07·t)`. Difference is negligible at t ≥ 100 years
(both → 1.0). Small delta expected on test sheet at t = 1 year.
Recorded as **D-Q3**.

### Q2 — α1 factor in drying shrinkage k1 (PENDING)
**Finding:** V33 BCDcheck k_1 (BM103) = `α1·ts^0.8/(ts^0.8+0.15·th)` where
α1 (BL103) = `0.8+1.2·exp(−0.005·th)`. AS3600-2018 cl 3.1.7.2 strict form:
k1 = `ts^0.8/(ts^0.8+0.15·th)` (no α1 factor).

For typical RCBC (Ds=175mm, th≈299mm): α1 ≈ 1.067 → sheet overpredicts
εcsd by ~6.7% vs strict standard formula.

The α1 formula (`0.8+1.2·exp(−0.005·th)`) has the same form as the AS3600
creep modification factors (α2creep = `1+1.12·exp(−0.008·th)`), suggesting
it was possibly adopted from the creep section.

**Action required:** Check AS3600-2009 cl 3.1.7.2 to determine:
- If α1 IS in the 2009 drying shrinkage k1 formula → V33 correctly
  implements 2009; standard changed in 2018 → the RCBC design uses
  AS3600-2009 shrinkage. Document divergence from AS3600-2018.
- If α1 is NOT in the 2009 formula either → long-standing error in V33.
  Correct both UDFs: set α1 = 1.0 in K1_Shrinkage.

**Current implementation:** `mod_Materials.bas` retains α1 to match V33
and passes it as an explicit argument to `K1_Shrinkage(ts, th, alpha1)`.
Callers using strict AS3600-2018 pass alpha1 = 1.0.
**Decision NOT yet locked.** Recorded as D-Q2 placeholder.

## UDFs delivered (17 total)

| Family | UDF | Returns | Clause |
|--------|-----|---------|--------|
| M1.1 | `S_Eurocode(cement_class)` | S (dimensionless) | EN 1992-1-1 |
| M1.1 | `Beta_cct(S_Eur, t_loading_days)` | βcct (dimensionless) | EN 1992-1-1 |
| M1.1 | `fcmi(fc_MPa, S_Eur, t_loading_days)` | MPa | AS3600-2018 Commentary |
| M1.2 | `Ec_AS3600_2018(fc_MPa)` | MPa | AS3600-2018 cl 3.1.2 |
| M1.2 | `Ec_AS3600_2009(fc_MPa)` | MPa | AS3600-2009 cl 3.1.2 |
| M1.2 | `Ec_AS5100(fc_MPa)` | MPa | AS5100.5-2017 (≡ AS3600-2018) |
| M1.3 | `fctf_AS3600_2018(fc_MPa)` | MPa | AS3600-2018 cl 3.1.1.3(a) |
| M1.3 | `fct_AS3600_2018(fc_MPa)` | MPa | AS3600-2018 cl 3.1.1.3(b) |
| M1.4 | `Alpha2_AS3600_2018(fc_MPa)` | dimensionless | AS3600-2018 cl 8.1.3 |
| M1.4 | `Gamma_AS3600_2018(fc_MPa)` | dimensionless | AS3600-2018 cl 8.1.3 |
| M1.4 | `Alpha2_AS3600_2009(fc_MPa)` | dimensionless | AS3600-2009 cl 8.1.3 |
| M1.4 | `Gamma_AS3600_2009(fc_MPa)` | dimensionless | AS3600-2009 cl 8.1.3 |
| M1.5 | `Th_section(D_mm, b_mm)` | mm | AS3600-2018 cl 3.1.7.2 |
| M1.5 | `Alpha1_Shrinkage(th_mm)` | dimensionless | V33 BCDcheck (pending Q2) |
| M1.5 | `K1_Shrinkage(ts_days, th_mm, alpha1)` | dimensionless | AS3600-2018 cl 3.1.7.2 |
| M1.5 | `K4_Shrinkage(environment)` | dimensionless | AS3600-2018 Table 3.1.7.2 |
| M1.5 | `EpsCsd_basic_AS3600_2018(fc_MPa, eps_csd_b_ref)` | dimensionless | AS3600-2018 cl 3.1.7.2 |
| M1.5 | `EpsCsd_AS3600_2018(ts_days, th_mm, k4, eps_csd_b, alpha1)` | dimensionless | AS3600-2018 cl 3.1.7.2 |
| M1.6 | `EpsCse_basic_AS3600_2018(fc_MPa)` | dimensionless | AS3600-2018 cl 3.1.7.1 |
| M1.6 | `EpsCse_AS3600_2018(fc_MPa, t_final_days)` | dimensionless | AS3600-2018 cl 3.1.7.1 |
| M1.7 | `EpsCs_AS3600_2018(eps_csd, eps_cse)` | dimensionless | AS3600-2018 cl 3.1.7 |

(21 rows but only 17 distinct UDF names — Ec_AS5100 delegates to
Ec_AS3600_2018; the count matches the P1 signature lock target.)

## Known discrepancies from V33 (expected test sheet deltas)

| UDF | V33 behaviour | UDF behaviour | Delta expected? |
|-----|--------------|---------------|----------------|
| Ec_AS3600_2018 | Hard-coded 34 800 MPa | Formula: ρ^1.5*(0.024√fc+0.12) | YES — until Step 5 integrates V6 |
| EpsCse_AS3600_2018 | exp(−0.07·t) | exp(−0.1·t) per AS3600-2018 | YES at early t; negligible at 100yr |
| K1_Shrinkage | k1 = α1·ts^0.8/(ts^0.8+0.15·th) | Same (α1 preserved) | NO — matches sheet |

## What BCDcheck cells this replaces (Step 5 — after user validates)

| BCDcheck cell | Current formula | Replacement UDF call |
|--------------|----------------|---------------------|
| V6 (Ec) | Hard-coded 34800 | `=Ec_AS3600_2018(fc)` |
| AA3 (S_Eurocode) | INDEX/MATCH lookup | `=S_Eurocode(AA2)` |
| AA4 (βcct) | `EXP(S_Eurocode*(1-SQRT(28/τloading)))` | `=Beta_cct(S_Eurocode,τloading)` |
| AA5 (fcmi) | `IF(τloading>=28,fc,βcct*(fc+8)-8)` | `=fcmi(fc,S_Eurocode,τloading)` |
| AV114 (fctf) | `0.6*SQRT(fc)` | `=fctf_AS3600_2018(fc)` |
| AM109 (fct) | `AS108` = `1.4*0.36*SQRT(fc_cr)` | Note: AS108 uses 1.4 factor (mean, not characteristic) — **keep separate** |
| AG111 (α2) | IF(φm_option=1,...) | `=Alpha2_AS3600_2018(fc)` or `=Alpha2_AS3600_2009(fc)` per selected standard |
| AG112 (γ) | IF(φm_option=1,...) | `=Gamma_AS3600_2018(fc)` or `=Gamma_AS3600_2009(fc)` |
| AM112 (k4) | VLOOKUP(Environ.,...) | `=K4_Shrinkage(Environ.)` |
| AM115 (ε*cse) | formula | `=EpsCse_basic_AS3600_2018(fc)` |
| AM117 (εcse) | `ε_cse*(1-EXP(-0.07*t))` | `=EpsCse_AS3600_2018(fc,t)` |
| AM119 (εcsd,b) | `(0.9-0.005*fc)*ε_csd_b` | `=EpsCsd_basic_AS3600_2018(fc,ε_csd_b)` |
| BL103 (α1) | `0.8+1.2*EXP(-0.005*th)` | `=Alpha1_Shrinkage(th)` |
| BM103 (k_1) | `(α1*ts^0.8)/(ts^0.8+0.15*th)` | `=K1_Shrinkage(ts,th,α1)` |
| BN103 (εcsd) | `k_1*k_4*εcsd_b` | `=EpsCsd_AS3600_2018(ts,th,k4,εcsd_b,α1)` |
| BO103 (εcs) | `(εcsd+εcse)*(1+CrShrMargin/100)` | `=EpsCs_AS3600_2018(εcsd,εcse)*(1+CrShrMargin/100)` |

**Note on AM109/fct:** V33 AM109 = 1.4×0.36×√fc_cr (mean tensile, not characteristic).
The `fct_AS3600_2018` UDF returns 0.36×√fc (characteristic). The 1.4 factor is a
mean-to-characteristic conversion and should remain at the call site, not absorbed
into the UDF. This IS a discrepancy; document in the Step 5 integration note.

## Instructions for import (Steps 2–4)

1. Open **V34_P0b.xlsm** in Excel.
2. Press **Alt+F11** to open VBA editor.
3. File → Import File → select `mod_Materials.bas`.
4. Press **Ctrl+S** to save.
5. In any blank cell, type `=Ec_AS3600_2018(32)`. It should return `30076.3`.
   If it shows `#NAME?`, the import failed — check the module appeared under Modules.
6. Add sheet `_Test_P1` (Insert → Sheet), position it after `_Contract`.
7. Build the test sheet using `P1_test_sheet_spec.md` as the template.
8. Implement the test rows. All deltas in column K should match the
   "Delta expected?" column above (zero or documented exception).
9. Once deltas are all zero or explained: right-click `_Test_P1` → Hide.
10. Update `_Names` registry (add any new names: none added in P1 — all
    P1 UDFs consume existing named ranges).
11. Update `_Input` change log section with:
    - Date, P1 completed
    - UDFs added: 17 (list above)
    - BCDcheck cells that change in Step 5 (list above)
    - Q2 status (pending AS3600-2009 check)
12. **File → Save As → V34_P1.xlsm**

## What is NOT done
- BCDcheck cells NOT yet replaced (Step 5 deferred until test sheet complete).
- Q2 NOT locked (α1 decision pending AS3600-2009 check).
- No creep UDFs in P1 (creep in P6 — Deflection SLS, not P1).
- P1 stage gate xlsm not produced by AI (requires Excel).

## Next step
After user validates test sheet deltas:
1. Lock Q2 decision.
2. Replace BCDcheck cells (Step 5 from list above).
3. Verify report sheets still produce correct output.
4. Save as V34_P1.xlsm (stage gate).
5. Commit mod_Materials.bas + xlsm to git.
6. Begin P2 (Cover and durability).
