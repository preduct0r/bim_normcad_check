Option Explicit
Const c As Long = 5
Dim Ar(1 To c) As Single

Public Sub ArmX()
        On Error GoTo 1
        Dim Vars As Object
        Dim Conds As Object
        Set Vars = CreateObject("NC_167258518177598E02.Vars")
        Set Conds = Vars.Conds
        
        Vars(VN("gr_g__b1")).Value = 1
        Vars(VN("m__kp")).Value = 1

        Vars(VN("s__íx")).Value = 0.1
        Vars(VN("s__âx")).Value = 0.1
        Vars(VN("s__íy")).Value = 0.1
        Vars(VN("s__ây")).Value = 0.1
        Vars(VN("a__íx")).Value = 0.04
        Vars(VN("a__âx")).Value = 0.04
        Vars(VN("a__íy")).Value = 0.04
        Vars(VN("a__ây")).Value = 0.04
        Vars(VN("h")).Value = 0.2
        Vars(VN("b")).Value = 1
        
        Conds.Add "Àðìàòóðà ðàñïîëîæåíà ïî êîíòóðó ñå÷åíèÿ - íå ðàâíîìåðíî"
        Conds.Add "Ãðóïïà ïðåäåëüíûõ ñîñòîÿíèé - ïåðâàÿ"
        Conds.Add "Êîíñòðóêöèÿ - æåëåçîáåòîííàÿ"
        Conds.Add "Íàçíà÷åíèå êëàññà áåòîíà - ïî ïðî÷íîñòè íà ñæàòèå"
        Conds.Add "Îòíîñèòåëüíàÿ âëàæíîñòü âîçäóõà îêðóæàþùåé ñðåäû - 40 - 75%"
        Conds.Add "Ïîïåðåìåííîå çàìîðàæèâàíèå è îòòàèâàíèå ïðè òåìïåðàòóðå < 20°C - îòñóòñòâóåò"
        Conds.Add "Àðìàòóðà ïëèò - âåðõíÿÿ è íèæíÿÿ (èçãèá. ìîìåíòû ââîäÿòñÿ ñî ñâîèìè çíàêàìè)"
        Conds.Add "Ñå÷åíèå - ïðÿìîóãîëüíîå"
        Conds.Add "Ýëåìåíò - èçãèáàåìûé"
        Conds.Add "Ïðîãðåññèðóþùåå ðàçðóøåíèå - íå ðàññìàòðèâàåòñÿ â äàííîì ðàñ÷åòå"
        Conds.Add "Êîíñòðóêöèÿ áåòîíèðóåòñÿ - â ãîðèçîíòàëüíîì ïîëîæåíèè"
        Conds.Add "Êëàññ áåòîíà - B30"
        Conds.Add "Äåéñòâèå íàãðóçêè - íåïðîäîëæèòåëüíîå"
        Conds.Add "Ñåéñìè÷íîñòü ïëîùàäêè ñòðîèòåëüñòâà - íå áîëåå 6 áàëëîâ"
        Conds.Add "Êëàññ ïðîäîëüíîé àðìàòóðû - A400"
        Conds.Add "Ïîïåðå÷íàÿ àðìàòóðà - íå ðàññìàòðèâàåòñÿ â äàííîì ðàñ÷åòå"
        
        Ar(1) = 10
        Ar(2) = 12
        Ar(3) = 14
        Ar(4) = 16
        Ar(5) = 20
        
        Dim Row As Long
        Dim Col As Long
        Dim CellText As String
        Dim NCResult As Single
        
        Dim ix As Long
        Dim jx As Long
        Dim iy As Long
        Dim jy As Long
        
        Columns("F:G").Select
        Selection.ClearContents
        Range("A1").Select
        
        Vars.Ex ("S_" & VN("5.1.8"))
        Vars.Ex ("S_" & VN("5.1.9"))
        Vars.Ex ("S_" & VN("5.1.10"))
        Vars.Ex ("S_" & VN("5.2.7"))
        Vars.Ex ("S_" & VN("5.2.10"))
        
        Do
           Row = Row + 1

           CellText = Cells(Row, 1)
           If CellText = "" Then Exit Do

           Vars("M__x").Value = Cells(Row, 1)
           Vars("M__y").Value = Cells(Row, 2)
           Vars("M__xy").Value = Cells(Row, 3)
           Vars("Q__x").Value = Cells(Row, 4)
           Vars("Q__y").Value = Cells(Row, 5)
           
                           
           For ix = 1 To c
               For iy = 1 To c
                   For jx = 1 To c
                       For jy = 1 To c
                       
                           Vars("d__síx").Value = Ar(ix)
                           Vars("d__sâx").Value = Ar(jx)
                           Vars("d__síy").Value = Ar(iy)
                           Vars("d__sây").Value = Ar(jy)
                       
                           NCResult = 0
                           Vars.Result = 0

                           Vars.Ex ("S_" & VN("6.2.7"))
                           NCResult = Max(NCResult, Vars.Result)
                           Vars.Ex ("S_" & VN("8.4 ÑÏ 52-103"))
                           NCResult = Max(NCResult, Vars.Result)
                           Vars.Ex ("S_" & VN("8.5 ÑÏ 52-103"))
                           NCResult = Max(NCResult, Vars.Result)
                           Vars.Ex ("S_" & VN("8.3.4"))
                           NCResult = Max(NCResult, Vars.Result)
                           
                           If NCResult <= 1 Then
                               Cells(Row, 6) = Ar(ix) & " x " & Ar(jx) & " / " & Ar(iy) & " x " & Ar(jy)
                               Exit For
                           End If
                       Next
                       If NCResult <= 1 Then Exit For
                   Next
                   If NCResult <= 1 Then Exit For
               Next
               If NCResult <= 1 Then Exit For
           Next
           
           Cells(Row, 7) = NCResult
           
        Loop
        
        Exit Sub
1:
        If Err.Number <> 56401 Then MsgBox "Error " & Err.Number & "; Description: " & Err.Description, vbCritical, "Error"
End Sub
Private Function Max(A As Single, B As Single) As Single
        Max = IIf(A > B, A, B)
End Function
Private Function VN(Name As String) As String
        VN = Replace(Name, " ", "_spc_")
        VN = Replace(VN, ".", "_pnt_")
        VN = Replace(VN, "-", "_minus_")
        VN = Replace(VN, "(", "_bkt1_")
        VN = Replace(VN, ")", "_bkt2_")
End Function