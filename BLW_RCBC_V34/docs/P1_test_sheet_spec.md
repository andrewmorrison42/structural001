# _Test_P1 Sheet Specification
## P1 — Materials | BLW RCBC Section Analysis V34

Add a hidden sheet named `_Test_P1` to the workbook immediately after
`_Contract`. Layout follows the convention in the _Contract sheet (columns A–M).

---

### Column layout (all test rows)

| Col | Content |
|-----|---------|
| A | UDF name (text label) |
| B–H | Input values (one per column, labelled row 1) |
| I | UDF call result |
| J | Reference value (BCDcheck cell or hand-computed) |
| K | Delta = I − J |
| L | Tolerance OK = ABS(K) < 1E-6 |
| M | Scenario notes |

---

### M1.1 — Mean compressive strength

#### S_Eurocode
| Row | B (cement_class) | I (UDF) | J (Reference) | Expected |
|-----|-----------------|---------|---------------|---------|
| 2 | "Class N" | =S_Eurocode(B2) | 0.25 | 0.25 |
| 3 | "Class R" | =S_Eurocode(B3) | 0.20 | 0.20 |
| 4 | "Class S" | =S_Eurocode(B4) | 0.38 | 0.38 |
| 5 | "Class X" | =S_Eurocode(B5) | #VALUE! | error |

#### Beta_cct
Inputs: B=S_Eur, C=t_loading_days. Reference: BCDcheck AA4.

| Row | B | C | I | J | Scenario |
|-----|---|---|---|---|---------|
| 7 | 0.25 | 28 | =Beta_cct(B7,C7) | 1.0 | t = 28 d → β = 1.0 |
| 8 | 0.25 | 3 | =Beta_cct(B8,C8) | BCDcheck!AA4 (with τloading=3) | Early age |
| 9 | 0.38 | 7 | =Beta_cct(B9,C9) | hand: exp(0.38*(1-√(28/7))) = exp(0.38*(-1)) ≈ 0.684 | Class S, 7d |
| 10 | 0.25 | 0 | =Beta_cct(B10,C10) | #VALUE! | edge: t=0 |

#### fcmi
Inputs: B=fc_MPa, C=S_Eur, D=t_loading_days. Reference: BCDcheck AA5.

| Row | B | C | D | I | J | Scenario |
|-----|---|---|---|---|---|---------|
| 12 | 32 | 0.25 | 28 | =fcmi(B12,C12,D12) | 32 | t≥28 → fcmi = fc |
| 13 | 32 | 0.25 | 3 | =fcmi(B13,C13,D13) | BCDcheck!AA5 | match sheet |
| 14 | 50 | 0.25 | 7 | =fcmi(B14,C14,D14) | hand-compute | mid-age |

---

### M1.2 — Modulus of elasticity

Inputs: B=fc_MPa. References: BCDcheck V6 (hard-coded 34800 — will differ).

| Row | B | I (AS3600-2018) | I (AS3600-2009) | Hand-computed | Scenario |
|-----|---|---------|---------|---------|---------|
| 16 | 32 | =Ec_AS3600_2018(B16) | =Ec_AS3600_2009(B16) | 2400^1.5*(0.024*√32+0.12) = 30 076 | fc=32 |
| 17 | 40 | =Ec_AS3600_2018(B17) | | 2400^1.5*(0.024*√40+0.12) = 32 320 | fc=40 |
| 18 | 50 | =Ec_AS3600_2018(B18) | | 2400^1.5*(0.024*√50+0.12) = 34 074 | fc=50 |
| 19 | 65 | =Ec_AS3600_2018(B19) | | 2400^1.5*(0.024*√65+0.12) = 36 432 | fc=65 |
| 20 | 80 | =Ec_AS3600_2018(B20) | | 2400^1.5*(0.024*√80+0.12) = 38 606 | fc=80 |

**Note:** BCDcheck V6 = 34 800 (hard-coded). The UDF gives 34 074 at fc=50.
The delta here is EXPECTED until V6 is replaced with =Ec_AS3600_2018(fc) in Step 5.

---

### M1.3 — Tensile strength

Inputs: B=fc_MPa. References: BCDcheck AV114 (fctf), AM109 (fct).

| Row | B | fctf call | fctf ref | fct call | fct ref | Scenario |
|-----|---|-----------|----------|----------|---------|---------|
| 22 | 32 | =fctf_AS3600_2018(B22) | =0.6*SQRT(B22) = 3.394 | =fct_AS3600_2018(B22) | =0.36*SQRT(B22) = 2.037 | fc=32 |
| 23 | 40 | =fctf_AS3600_2018(B23) | 3.795 | =fct_AS3600_2018(B23) | 2.277 | fc=40 |
| 24 | 50 | =fctf_AS3600_2018(B24) | 4.243 | =fct_AS3600_2018(B24) | 2.546 | fc=50 |

---

### M1.4 — Stress block parameters

Inputs: B=fc_MPa.

**Acceptance:** Match AS3600-2018 Table 8.1.3 to 4 decimal places at fc = 25, 32, 40, 50, 65, 80, 100 MPa.

| Row | B (fc) | α2-2018 | γ-2018 | α2-2009 | γ-2009 |
|-----|--------|---------|--------|---------|--------|
| 26 | 25 | =Alpha2_AS3600_2018(B26) | =Gamma_AS3600_2018(B26) | =Alpha2_AS3600_2009(B26) | =Gamma_AS3600_2009(B26) |
| 27 | 32 | | | | |
| 28 | 40 | | | | |
| 29 | 50 | | | | |
| 30 | 65 | | | | |
| 31 | 80 | | | | |
| 32 | 100 | | | | |

Expected values (AS3600-2018 formula α2 = MAX(0.67, 0.85−0.0015·fc)):

| fc | α2-2018 | γ-2018 | α2-2009 | γ-2009 |
|----|---------|--------|---------|--------|
| 25 | 0.8125 | 0.9075 | 0.8500 | 0.8750 |
| 32 | 0.8020 | 0.8900 | 0.8500 | 0.8250 |
| 40 | 0.7900 | 0.8700 | 0.8500 | 0.7700 |
| 50 | 0.7750 | 0.8450 | 0.8500 | 0.7000 |
| 65 | 0.7525 | 0.8075 | 0.8050 | 0.6700 |
| 80 | 0.7300 | 0.7700 | 0.7600 | 0.6700 |
| 100 | 0.7000 | 0.7200 | 0.7000 | 0.6700 |

**Note:** Verify these expected values against the actual AS3600-2018 Table 8.1.3
before signing off P1. If the table disagrees, the formula has a discrepancy.

---

### M1.5 — Drying shrinkage

Inputs: B=D_mm (or ts_days), C=b_mm (or th_mm), etc.

#### Th_section
| Row | B (D_mm) | C (b_mm) | Call | Expected | Scenario |
|-----|----------|----------|------|----------|---------|
| 34 | 175 | 1000 | =Th_section(B34,C34) | 2*175*1000/(2*175+2*1000) = 298.7 | RCBC typical |
| 35 | 300 | 1000 | =Th_section(B35,C35) | 2*300*1000/(2*300+2*1000) = 461.5 | thick slab |
| 36 | 175 | 1200 | =Th_section(B36,C36) | 2*175*1200/(2*175+2*1200) = 305.9 | 1.2m unit |

#### Alpha1_Shrinkage
| Row | B (th_mm) | Call | Expected | BCDcheck ref |
|-----|----------|------|----------|-------------|
| 38 | 298.7 | =Alpha1_Shrinkage(B38) | 0.8+1.2*EXP(-0.005*298.7) ≈ 1.0668 | BCDcheck BL103 |
| 39 | 100 | =Alpha1_Shrinkage(B39) | 0.8+1.2*EXP(-0.5) ≈ 1.5274 | thin |
| 40 | 1000 | =Alpha1_Shrinkage(B40) | ≈ 0.8081 | thick |

#### K1_Shrinkage
Inputs: B=ts_days, C=th_mm, D=alpha1.

| Row | B | C | D | Call | BCDcheck ref | Scenario |
|-----|---|---|---|------|-------------|---------|
| 42 | 36500 (100yr) | 298.7 | =Alpha1_Shrinkage(C42) | =K1_Shrinkage(B42,C42,D42) | BCDcheck BM103 | Final time |
| 43 | 365 (1yr) | 298.7 | =Alpha1_Shrinkage(C43) | =K1_Shrinkage(B43,C43,D43) | | 1 year |
| 44 | 36500 | 298.7 | 1.0 | =K1_Shrinkage(B44,C44,D44) | | strict AS3600 (α1=1) |

#### K4_Shrinkage
| Row | B (environment) | Call | Expected |
|-----|----------------|------|----------|
| 46 | "Arid" | =K4_Shrinkage(B46) | 0.70 |
| 47 | "Interior" | =K4_Shrinkage(B47) | 0.65 |
| 48 | "Temperate inland" | =K4_Shrinkage(B48) | 0.60 |
| 49 | "Tropical or near coastal or coastal" | =K4_Shrinkage(B49) | 0.50 |
| 50 | "Invalid" | =K4_Shrinkage(B50) | #VALUE! |

#### EpsCsd_basic_AS3600_2018
Inputs: B=fc_MPa, C=eps_csd_b_ref.

| Row | B | C | Call | BCDcheck ref | Scenario |
|-----|---|---|------|-------------|---------|
| 52 | 32 | 0.0008 | =EpsCsd_basic_AS3600_2018(B52,C52) | BCDcheck AM119 at fc=32 → 0.000592 | fc=32 Arid |
| 53 | 50 | 0.0008 | =EpsCsd_basic_AS3600_2018(B53,C53) | 0.000520 | fc=50 |
| 54 | 20 | 0.0008 | =EpsCsd_basic_AS3600_2018(B54,C54) | 0.000640 | fc=20 |

#### EpsCsd_AS3600_2018 (full)
Cross-check against BCDcheck BN103 (= εcsd named range array element 1).

| Row | ts_days | th_mm | k4 | eps_csd_b | alpha1 | Call | BCDcheck ref |
|-----|---------|-------|----|-----------|--------|------|-------------|
| 56 | =Design_Life*365-τdrying | =Th_section(Ds,Lunit) | =K4_Shrinkage(Environ.) | =EpsCsd_basic_AS3600_2018(fc,0.0008) | =Alpha1_Shrinkage(th) | =EpsCsd_AS3600_2018(…) | =BN103 |

**Delta at row 56 must be < 1E-6 for P1 to be implementation-complete.**

---

### M1.6 — Endogenous shrinkage

#### EpsCse_basic_AS3600_2018
| Row | B (fc_MPa) | Call | BCDcheck ref (AM115) | Notes |
|-----|-----------|------|---------------------|-------|
| 58 | 32 | =EpsCse_basic_AS3600_2018(B58) | BCDcheck AM115 formula | |
| 59 | 50 | =EpsCse_basic_AS3600_2018(B59) | | boundary at fc=50 |
| 60 | 65 | =EpsCse_basic_AS3600_2018(B60) | | fc>50 branch |

#### EpsCse_AS3600_2018
Inputs: B=fc_MPa, C=t_final_days.

| Row | B | C | Call | BCDcheck ref (AM117) | Delta notes |
|-----|---|---|------|---------------------|-------------|
| 62 | 32 | 36500 | =EpsCse_AS3600_2018(B62,C62) | BCDcheck AM117 | EXPECTED small delta: UDF uses exp(−0.1t), V33 used exp(−0.07t). Document. |
| 63 | 50 | 36500 | =EpsCse_AS3600_2018(B63,C63) | | |
| 64 | 32 | 365 | =EpsCse_AS3600_2018(B64,C64) | | 1 year — larger delta here |

---

### M1.7 — Total shrinkage

#### EpsCs_AS3600_2018
| Row | B (εcsd) | C (εcse) | Call | Expected | Notes |
|-----|---------|---------|------|----------|-------|
| 66 | =EpsCsd_AS3600_2018(…) | =EpsCse_AS3600_2018(…) | =EpsCs_AS3600_2018(B66,C66) | B66+C66 | No margin |
| 67 | 0.0004 | 0.0001 | =EpsCs_AS3600_2018(B67,C67) | 0.0005 | hand check |

---

## Acceptance gate

P1 is implementation-complete when:
1. All delta cells in column K show < 1E-6 (or documented as expected exceptions).
2. Exception rows (Ec hardcode mismatch, EpsCse coefficient divergence) are flagged in column M.
3. α2 and γ match AS3600-2018 Table 8.1.3 at fc = 25, 32, 40, 50, 65, 80, 100 MPa (user to verify).
4. Q2 decision recorded (α1 in k1 confirmed or corrected after checking AS3600-2009).

After acceptance: hide this sheet, update _Names, update _Input change log, save V34_P1.xlsm.
