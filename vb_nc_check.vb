' VB/VBA skeleton for running one check
Option Explicit

Public Type CheckResult
    Passed As Boolean
    MaxUtil As Double
    Message As String
End Type

Public Function CheckTimberBeam_SP64(ByVal L_mm As Double, ByVal b_mm As Double, ByVal h_mm As Double, _
                                    ByVal q_kN_per_m As Double, ByVal materialName As String) As CheckResult
    Dim r As CheckResult
    On Error GoTo EH

    Dim ncApiR As Object
    Set ncApiR = CreateObject("ncApi.Report") ' NCAPI.dll 

    ' 1) Choose module / task
    ncApiR.Norm = "СП 64.13330.2017 Деревянные конструкции"  ' пример: уточни по своему модулю
    ncApiR.TaskName = "Балка"                                ' взять из сгенерированного проекта
    ncApiR.Unit = "1"                                        ' взять из сгенерированного проекта

    ncApiR.ClcLoadNorm

    ' 2) Vars object (ProgID берёшь из сгенерированного VB проекта)
    Dim Vars As Object
    Set Vars = CreateObject("NC_XXXXXXXXXXXXXXX.Vars")

    ' 3) Assign input vars (имена полей — из сгенерированного проекта!)
    Vars.L = L_mm
    Vars.b = b_mm
    Vars.h = h_mm
    Vars.q = q_kN_per_m
    Vars.Material = materialName

    ncApiR.SetVars Vars

    ' 4) Conditions array (если нужно) — тоже видно в сгенерированном проекте
    Dim Conds(0 To 0) As String
    Conds(0) = "..."   ' например: схема опирания / длительность нагрузки / класс эксплуатации
    ncApiR.SetConds Conds

    ncApiR.ClcLoadData
    ncApiR.ClcLoadConds
    ncApiR.ClcCalc

    r.MaxUtil = ncApiR.MaxResult  ' 
    r.Passed = (r.MaxUtil <= 1#)

    If r.Passed Then
        r.Message = "OK"
    Else
        r.Message = "FAIL: Max utilization = " & CStr(r.MaxUtil)
    End If

    CheckTimberBeam_SP64 = r
    Exit Function

EH:
    r.Passed = False
    r.MaxUtil = 9999
    r.Message = "ERROR: " & Err.Description
    CheckTimberBeam_SP64 = r
End Function
