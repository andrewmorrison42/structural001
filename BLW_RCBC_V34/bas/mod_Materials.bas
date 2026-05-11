Attribute VB_Name = "mod_Materials"
'=============================================================================
' mod_Materials.bas
' P1 — Materials
' BLW RCBC Section Analysis V34
'
' UDF library for concrete and steel material properties.
' All UDFs are pure functions (no global state, no sheet reads).
' All internals in SI base units: N, N·mm, mm, MPa.
' Boundary conversions (kN→N, kNm→N·mm, m→mm) happen at the call site.
'
' Contract rules (per _Contract sheet R1–R10):
'   R1 – Pure function. No ActiveSheet, Range(), module-level variables.
'   R2 – As Variant return type (allows Double or CVErr).
'   R3 – Arguments declared As Double (or As String for enum args).
'   R4 – One UDF per code version. No internal code-string switches.
'   R5 – Unit-suffixed argument names (fc_MPa, d_mm, M_star_Nmm).
'   R6 – Invalid input → CVErr(xlErrValue). Never return sentinel -1/0/"ERR".
'   R7 – On Error GoTo bad_input pattern in every UDF.
'   R8 – SI internals only. No unit conversion inside UDFs.
'   R9 – Cite the standard clause in every function header.
'  R10 – No magic numbers without a comment tracing them to their source.
'
' Decisions recorded:
'   D-Q1 : th convention — match V33 sheet exactly (two-face drying formula).
'   D-Q2 : α1 in k1 shrinkage — α1 factor included to match V33 BCDcheck;
'           pending user verification against AS3600-2009 cl 3.1.7.2 vs
'           AS3600-2018 cl 3.1.7.2. See Alpha1_Shrinkage and K1_Shrinkage headers.
'   D-Q3 : EpsCse time coefficient — corrected to exp(−0.1·t) per
'           AS3600-2018 cl 3.1.7.1. V33 used exp(−0.07·t); difference
'           is negligible at t ≥ 100 years.
'=============================================================================
Option Explicit

'─────────────────────────────────────────────────────────────────────────────
' M1.1  Mean compressive strength (Eurocode-based maturity)
'─────────────────────────────────────────────────────────────────────────────

Public Function S_Eurocode(cement_class As String) As Variant
    ' Cement class coefficient S used in the Eurocode maturity function.
    ' EN 1992-1-1 Table 3.1 / AS3600-2018 Commentary.
    ' cement_class: "Class N" | "Class R" | "Class S"
    ' Returns: S coefficient (dimensionless).
    ' Pure function. No SI boundary issues.
    On Error GoTo bad_input
    Select Case Trim(cement_class)
        Case "Class N": S_Eurocode = 0.25
        Case "Class R": S_Eurocode = 0.20
        Case "Class S": S_Eurocode = 0.38
        Case Else:       GoTo bad_input
    End Select
    Exit Function
bad_input:
    S_Eurocode = CVErr(xlErrValue)
End Function


Public Function Beta_cct(S_Eur As Double, t_loading_days As Double) As Variant
    ' Maturity factor βcct for concrete compressive strength at age t.
    ' Ref: Eurocode EN 1992-1-1 cl 3.1.2 (same form used in AS3600 Commentary).
    ' βcct = exp( S_Eur × (1 − √(28 / t_loading_days)) )
    ' Returns: βcct (dimensionless, > 0).
    ' Pure function. SI internals: t in days (dimensionless ratio).
    On Error GoTo bad_input
    If t_loading_days <= 0 Then GoTo bad_input
    Beta_cct = Exp(S_Eur * (1# - Sqr(28# / t_loading_days)))
    Exit Function
bad_input:
    Beta_cct = CVErr(xlErrValue)
End Function


Public Function fcmi(fc_MPa As Double, S_Eur As Double, t_loading_days As Double) As Variant
    ' Mean compressive strength at age t_loading_days (MPa).
    ' Ref: AS3600-2018 Commentary (Eurocode maturity model).
    ' BCDcheck reference: AA5 = IF(τloading>=28, fc, βcct*(fc+8)-8)
    '   where βcct is computed by Beta_cct(S_Eur, t_loading_days).
    ' Returns: fcmi in MPa.
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Or t_loading_days <= 0 Then GoTo bad_input
    If t_loading_days >= 28# Then
        fcmi = fc_MPa
    Else
        Dim b As Double
        b = Beta_cct(S_Eur, t_loading_days)
        If IsError(b) Then GoTo bad_input
        fcmi = b * (fc_MPa + 8#) - 8#
    End If
    Exit Function
bad_input:
    fcmi = CVErr(xlErrValue)
End Function


'─────────────────────────────────────────────────────────────────────────────
' M1.2  Modulus of elasticity
'─────────────────────────────────────────────────────────────────────────────

Public Function Ec_AS3600_2018(fc_MPa As Double) As Variant
    ' Mean modulus of elasticity for normal weight concrete (ρ = 2400 kg/m³).
    ' Ref: AS3600-2018 cl 3.1.2
    ' Ec = ρ^1.5 × (0.024√fc + 0.12)   (MPa, with ρ in kg/m³, fc in MPa)
    ' At ρ = 2400: ρ^1.5 = 117 576 (kg^1.5/m^4.5 → unitless coefficient = 117576)
    ' Returns: Ec in MPa.
    ' NOTE: V33 BCDcheck hard-coded Ec = 34 800 MPa regardless of fc.
    '        This UDF corrects that. BCDcheck cell V6 should be replaced
    '        with =Ec_AS3600_2018(fc) in the Step 5 integrate pass.
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    Const rho15 As Double = 117576#   ' 2400^1.5 kg^1.5/m^4.5
    Ec_AS3600_2018 = rho15 * (0.024# * Sqr(fc_MPa) + 0.12#)
    Exit Function
bad_input:
    Ec_AS3600_2018 = CVErr(xlErrValue)
End Function


Public Function Ec_AS3600_2009(fc_MPa As Double) As Variant
    ' Mean modulus of elasticity for normal weight concrete (ρ = 2400 kg/m³).
    ' Ref: AS3600-2009 cl 3.1.2
    ' Ec = ρ^1.5 × (0.024√fc + 0.12)   — identical formula to AS3600-2018.
    ' Returns: Ec in MPa.
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    Ec_AS3600_2009 = Ec_AS3600_2018(fc_MPa)
    Exit Function
bad_input:
    Ec_AS3600_2009 = CVErr(xlErrValue)
End Function


Public Function Ec_AS5100(fc_MPa As Double) As Variant
    ' Mean modulus of elasticity for normal weight concrete (ρ = 2400 kg/m³).
    ' Ref: AS5100.5-2017 (treated identical to AS3600-2018 per D5).
    ' Returns: Ec in MPa.
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    Ec_AS5100 = Ec_AS3600_2018(fc_MPa)
    Exit Function
bad_input:
    Ec_AS5100 = CVErr(xlErrValue)
End Function


'─────────────────────────────────────────────────────────────────────────────
' M1.3  Tensile strength
'─────────────────────────────────────────────────────────────────────────────

Public Function fctf_AS3600_2018(fc_MPa As Double) As Variant
    ' Characteristic flexural tensile strength.
    ' Ref: AS3600-2018 cl 3.1.1.3(a)
    ' f'ct,f = 0.6 × √f'c   (MPa)
    ' BCDcheck reference: AV114 = 0.6*SQRT(fc)
    ' Returns: fctf in MPa.
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    fctf_AS3600_2018 = 0.6# * Sqr(fc_MPa)
    Exit Function
bad_input:
    fctf_AS3600_2018 = CVErr(xlErrValue)
End Function


Public Function fct_AS3600_2018(fc_MPa As Double) As Variant
    ' Characteristic uniaxial (splitting) tensile strength.
    ' Ref: AS3600-2018 cl 3.1.1.3(b)
    ' f'ct = 0.36 × √f'c   (MPa)
    ' Returns: fct in MPa.
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    fct_AS3600_2018 = 0.36# * Sqr(fc_MPa)
    Exit Function
bad_input:
    fct_AS3600_2018 = CVErr(xlErrValue)
End Function


'─────────────────────────────────────────────────────────────────────────────
' M1.4  Equivalent rectangular stress block parameters
'─────────────────────────────────────────────────────────────────────────────

Public Function Alpha2_AS3600_2018(fc_MPa As Double) As Variant
    ' Stress block intensity factor α2 (dimensionless).
    ' Ref: AS3600-2018 cl 8.1.3 (2021 Amendment 2); same as AS1597.2-2013.
    ' φm_option = 1 branch in BCDcheck (BCDcheck AG111).
    ' α2 = MAX(0.67, 0.85 − 0.0015·f'c)
    ' Returns: α2 (dimensionless, 0.67 ≤ α2 ≤ 0.85).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    Alpha2_AS3600_2018 = Application.WorksheetFunction.Max(0.67#, 0.85# - 0.0015# * fc_MPa)
    Exit Function
bad_input:
    Alpha2_AS3600_2018 = CVErr(xlErrValue)
End Function


Public Function Gamma_AS3600_2018(fc_MPa As Double) As Variant
    ' Stress block depth factor γ (dimensionless).
    ' Ref: AS3600-2018 cl 8.1.3 (2021 Amendment 2); same as AS1597.2-2013.
    ' φm_option = 1 branch in BCDcheck (BCDcheck AG112).
    ' γ = MAX(0.67, 0.97 − 0.0025·f'c)
    ' Returns: γ (dimensionless, 0.67 ≤ γ ≤ 0.97).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    Gamma_AS3600_2018 = Application.WorksheetFunction.Max(0.67#, 0.97# - 0.0025# * fc_MPa)
    Exit Function
bad_input:
    Gamma_AS3600_2018 = CVErr(xlErrValue)
End Function


Public Function Alpha2_AS3600_2009(fc_MPa As Double) As Variant
    ' Stress block intensity factor α2 (dimensionless).
    ' Ref: AS3600-2009 cl 8.1.3 (φm_option ≠ 1 branch in BCDcheck).
    ' α2 = MAX(0.67, MIN(0.85, 1.0 − 0.003·f'c))
    ' Returns: α2 (dimensionless, 0.67 ≤ α2 ≤ 0.85).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    Alpha2_AS3600_2009 = Application.WorksheetFunction.Max( _
        0.67#, Application.WorksheetFunction.Min(0.85#, 1# - 0.003# * fc_MPa))
    Exit Function
bad_input:
    Alpha2_AS3600_2009 = CVErr(xlErrValue)
End Function


Public Function Gamma_AS3600_2009(fc_MPa As Double) As Variant
    ' Stress block depth factor γ (dimensionless).
    ' Ref: AS3600-2009 cl 8.1.3 (φm_option ≠ 1 branch in BCDcheck).
    ' γ = MAX(0.67, MIN(0.85, 1.05 − 0.007·f'c))
    ' Returns: γ (dimensionless, 0.67 ≤ γ ≤ 0.85).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    Gamma_AS3600_2009 = Application.WorksheetFunction.Max( _
        0.67#, Application.WorksheetFunction.Min(0.85#, 1.05# - 0.007# * fc_MPa))
    Exit Function
bad_input:
    Gamma_AS3600_2009 = CVErr(xlErrValue)
End Function


'─────────────────────────────────────────────────────────────────────────────
' M1.5  Drying shrinkage
'─────────────────────────────────────────────────────────────────────────────

Public Function Th_section(D_mm As Double, b_mm As Double) As Variant
    ' Theoretical thickness for two-face drying (AS3600-2018 cl 3.1.7.2).
    ' th = 2·Ac / ue = 2·(D·b) / (2·D + 2·b)   (mm)
    ' This is the hydraulic radius formula for a rectangular slab exposed
    ' on two faces (top and bottom). Convention matches V33 BCDcheck
    ' formula in BE103: 2*(Ds*Lunit)/(2*Lunit+2*Ds).
    ' D_mm: section depth (mm); b_mm: section width (mm, typically 1000 for 1m strip).
    ' Returns: th in mm.
    ' Decision D-Q1: match V33 sheet — two-face drying formula.
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If D_mm <= 0 Or b_mm <= 0 Then GoTo bad_input
    Th_section = 2# * D_mm * b_mm / (2# * D_mm + 2# * b_mm)
    Exit Function
bad_input:
    Th_section = CVErr(xlErrValue)
End Function


Public Function Alpha1_Shrinkage(th_mm As Double) As Variant
    ' Size-dependent shrinkage amplitude factor α1.
    ' Source: V33 BCDcheck named range α1 = BL103 = 0.8 + 1.2·exp(−0.005·th).
    ' Decision D-Q2: PENDING — this factor appears in the V33 drying shrinkage
    '   k1 formula but is NOT present in the strict AS3600-2018 cl 3.1.7.2 text.
    '   A similar factor appears in the AS3600 creep section (α2creep in V33).
    '   Verify against AS3600-2009 cl 3.1.7.2 to determine if this is:
    '     (a) correct per the 2009 standard and was removed in 2018, or
    '     (b) a long-standing labelling error in the sheet.
    '   Until resolved, this UDF reproduces the V33 BCDcheck behaviour.
    '   If AS3600-2018 strict formula is required: remove α1 from K1_Shrinkage
    '   and set α1 = 1.0 (i.e., the factor cancels out).
    ' Returns: α1 (dimensionless). Ranges from ~0.8 (thick) to ~2.0 (thin).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If th_mm <= 0 Then GoTo bad_input
    Alpha1_Shrinkage = 0.8# + 1.2# * Exp(-0.005# * th_mm)
    Exit Function
bad_input:
    Alpha1_Shrinkage = CVErr(xlErrValue)
End Function


Public Function K1_Shrinkage(ts_days As Double, th_mm As Double, alpha1 As Double) As Variant
    ' Time-dependent drying shrinkage factor k1.
    ' Source: V33 BCDcheck k_1 (BM103) — includes α1 factor (see D-Q2).
    ' Formula (V33): k1 = α1·ts^0.8 / (ts^0.8 + 0.15·th)
    ' Strict AS3600-2018 cl 3.1.7.2: k1 = ts^0.8 / (ts^0.8 + 0.15·th)  [no α1]
    ' Decision D-Q2: α1 retained from V33. Pass alpha1 = 1.0 to use strict standard.
    ' ts_days: duration of drying = t − τdrying (days, > 0).
    ' th_mm: theoretical thickness (mm, from Th_section).
    ' alpha1: size-dependent factor (from Alpha1_Shrinkage, or 1.0 for strict standard).
    ' Returns: k1 (dimensionless, 0 ≤ k1 ≤ α1, approaches α1 at long durations).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If ts_days <= 0 Or th_mm <= 0 Or alpha1 <= 0 Then GoTo bad_input
    Dim ts08 As Double
    ts08 = ts_days ^ 0.8#
    K1_Shrinkage = (alpha1 * ts08) / (ts08 + 0.15# * th_mm)
    Exit Function
bad_input:
    K1_Shrinkage = CVErr(xlErrValue)
End Function


Public Function K4_Shrinkage(environment As String) As Variant
    ' Environment factor k4 for drying shrinkage.
    ' Ref: AS3600-2018 Table 3.1.7.2.
    ' Source: V33 DROP DOWNS sheet rows 31–34 (named range k_4 = AM112).
    ' environment: "Arid" | "Interior" | "Temperate inland" |
    '              "Tropical or near coastal or coastal"
    ' Returns: k4 (dimensionless).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    Select Case Trim(environment)
        Case "Arid":                                    K4_Shrinkage = 0.7#
        Case "Interior":                                K4_Shrinkage = 0.65#
        Case "Temperate inland":                        K4_Shrinkage = 0.6#
        Case "Tropical or near coastal or coastal":     K4_Shrinkage = 0.5#
        Case Else:                                      GoTo bad_input
    End Select
    Exit Function
bad_input:
    K4_Shrinkage = CVErr(xlErrValue)
End Function


Public Function EpsCsd_basic_AS3600_2018(fc_MPa As Double, eps_csd_b_ref As Double) As Variant
    ' Basic drying shrinkage strain adjusted for concrete strength.
    ' Ref: AS3600-2018 cl 3.1.7.2 / V33 BCDcheck AM119.
    ' Formula: εcsd,b = (0.9 − 0.005·f'c) × ε*csd,b_ref
    ' eps_csd_b_ref: tabulated base value from Table 3.1.7.2 for the design
    '   environment (V33 default: 0.0008 for Arid, stored in AM118).
    ' Returns: εcsd,b (dimensionless strain).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Or eps_csd_b_ref <= 0 Then GoTo bad_input
    EpsCsd_basic_AS3600_2018 = (0.9# - 0.005# * fc_MPa) * eps_csd_b_ref
    Exit Function
bad_input:
    EpsCsd_basic_AS3600_2018 = CVErr(xlErrValue)
End Function


Public Function EpsCsd_AS3600_2018(ts_days As Double, th_mm As Double, _
                                    k4 As Double, eps_csd_b As Double, _
                                    alpha1 As Double) As Variant
    ' Design drying shrinkage strain εcsd.
    ' Ref: AS3600-2018 cl 3.1.7.2.
    ' εcsd = k1 × k4 × εcsd,b   (V33 BCDcheck BN103: k_1*k_4*εcsd_b)
    ' ts_days:   duration of drying t − τdrying (days).
    ' th_mm:     theoretical thickness (mm, from Th_section).
    ' k4:        environment factor (from K4_Shrinkage, or direct value).
    ' eps_csd_b: basic drying shrinkage strain (from EpsCsd_basic_AS3600_2018).
    ' alpha1:    size factor (from Alpha1_Shrinkage, or 1.0 for strict standard).
    ' Returns: εcsd (dimensionless strain, positive = shrinkage).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If ts_days <= 0 Or th_mm <= 0 Or k4 <= 0 Or eps_csd_b <= 0 Or alpha1 <= 0 Then GoTo bad_input
    Dim k1 As Variant
    k1 = K1_Shrinkage(ts_days, th_mm, alpha1)
    If IsError(k1) Then GoTo bad_input
    EpsCsd_AS3600_2018 = k1 * k4 * eps_csd_b
    Exit Function
bad_input:
    EpsCsd_AS3600_2018 = CVErr(xlErrValue)
End Function


'─────────────────────────────────────────────────────────────────────────────
' M1.6  Endogenous (autogenous) shrinkage
'─────────────────────────────────────────────────────────────────────────────

Public Function EpsCse_basic_AS3600_2018(fc_MPa As Double) As Variant
    ' Basic endogenous shrinkage strain ε*cse (dimensionless).
    ' Ref: AS3600-2018 cl 3.1.7.1.
    ' ε*cse = (0.07·f'c − 0.50) × 10^−4   for f'c ≤ 50 MPa
    ' ε*cse = (0.08·f'c − 0.50) × 10^−4   for f'c > 50 MPa
    ' BCDcheck reference: AM115 formula matches this clause.
    ' Returns: ε*cse (dimensionless). Negative (shrinkage) for fc > ~7 MPa.
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Then GoTo bad_input
    If fc_MPa <= 50# Then
        EpsCse_basic_AS3600_2018 = (0.07# * fc_MPa - 0.5#) * 0.00005#
    Else
        EpsCse_basic_AS3600_2018 = (0.08# * fc_MPa - 0.5#) * 0.00005#
    End If
    Exit Function
bad_input:
    EpsCse_basic_AS3600_2018 = CVErr(xlErrValue)
End Function


Public Function EpsCse_AS3600_2018(fc_MPa As Double, t_final_days As Double) As Variant
    ' Design endogenous shrinkage strain εcse at age t_final_days (dimensionless).
    ' Ref: AS3600-2018 cl 3.1.7.1
    ' εcse = ε*cse × (1 − exp(−0.1·t))
    ' Decision D-Q3: coefficient corrected to 0.1 per AS3600-2018 cl 3.1.7.1.
    '   V33 BCDcheck AM117 used exp(−0.07·t); difference is negligible at t ≥ 100 yrs
    '   (both → 1.0), but this UDF is corrected to the current standard.
    ' Returns: εcse (dimensionless strain, positive = shrinkage convention matches εcsd).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    If fc_MPa <= 0 Or t_final_days <= 0 Then GoTo bad_input
    Dim eps_cse_b As Variant
    eps_cse_b = EpsCse_basic_AS3600_2018(fc_MPa)
    If IsError(eps_cse_b) Then GoTo bad_input
    EpsCse_AS3600_2018 = eps_cse_b * (1# - Exp(-0.1# * t_final_days))
    Exit Function
bad_input:
    EpsCse_AS3600_2018 = CVErr(xlErrValue)
End Function


'─────────────────────────────────────────────────────────────────────────────
' M1.7  Total design shrinkage
'─────────────────────────────────────────────────────────────────────────────

Public Function EpsCs_AS3600_2018(eps_csd As Double, eps_cse As Double) As Variant
    ' Total design shrinkage strain εcs = εcsd + εcse (dimensionless).
    ' Ref: AS3600-2018 cl 3.1.7
    ' BCDcheck reference: BO103 = (εcsd + εcse) × (1 + CrShrMargin/100).
    '   NOTE: This UDF does NOT apply the CrShrMargin variation factor.
    '   Apply the margin at the call site: =EpsCs_AS3600_2018(...) * (1 + CrShrMargin/100)
    ' Returns: εcs (dimensionless, sum of drying + endogenous shrinkage strains).
    ' Pure function. SI internals.
    On Error GoTo bad_input
    EpsCs_AS3600_2018 = eps_csd + eps_cse
    Exit Function
bad_input:
    EpsCs_AS3600_2018 = CVErr(xlErrValue)
End Function
