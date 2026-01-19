Option Explicit
Public Function NCResult() As Single
        On Error GoTo 1
        Dim Vars As Object
        Dim Conds As Object
        Set Vars = CreateObject("NC_137667756294139E02.Vars")
        Set Conds = Vars.Conds
        
        Vars(VN("gr_g__b1")).Value = 1
        Vars(VN("m__kp")).Value = 1
        Vars(VN("M__x")).Value = 4.90332500325983E-02
        Vars(VN("d__síx")).Value = 10
        Vars(VN("s__íx")).Value = 0.02
        Vars(VN("d__sâx")).Value = 10
        Vars(VN("s__âx")).Value = 0.02
        Vars(VN("d__síy")).Value = 10
        Vars(VN("s__íy")).Value = 0.02
        Vars(VN("d__sây")).Value = 10
        Vars(VN("s__ây")).Value = 0.02
        Vars(VN("a__íx")).Value = 0.05
        Vars(VN("a__âx")).Value = 0.04
        Vars(VN("k__max")).Value = 1000
        Vars(VN("gr_d_")).Value = 0.1
        Vars(VN("M")).Value = 3.92266000260786E-02
        Vars(VN("h")).Value = 0.03
        Vars(VN("b")).Value = 1
        Vars(VN("a")).Value = 0.05
        Vars(VN("a_vert")).Value = 0.04
        Vars(VN("A__s")).Value = 0.00393
        Vars(VN("A_vert__s")).Value = 0.00393
        
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
        Conds.Add "Êëàññ áåòîíà - B10"
        Conds.Add "Äåéñòâèå íàãðóçêè - íåïðîäîëæèòåëüíîå"
        Conds.Add "Ñåéñìè÷íîñòü ïëîùàäêè ñòðîèòåëüñòâà - íå áîëåå 6 áàëëîâ"
        Conds.Add "Êëàññ ïðîäîëüíîé àðìàòóðû - A240"
        Conds.Add "Ïîïåðå÷íàÿ àðìàòóðà - íå ðàññìàòðèâàåòñÿ â äàííîì ðàñ÷åòå"
        
        Vars.Result = 0
        Vars.Ex ("S_" & VN("2.6"))
        NCResult = Max(NCResult, Vars.Result)
        Vars.Ex ("S_" & VN("2.7"))
        NCResult = Max(NCResult, Vars.Result)
        Vars.Ex ("S_" & VN("2.8"))
        NCResult = Max(NCResult, Vars.Result)
        Vars.Ex ("S_" & VN("2.18"))
        NCResult = Max(NCResult, Vars.Result)
        Vars.Ex ("S_" & VN("2.20"))
        NCResult = Max(NCResult, Vars.Result)
        Vars.Ex ("S_" & VN("8.4 ÑÏ 52-103"))
        NCResult = Max(NCResult, Vars.Result)
        Vars.Ex ("S_" & VN("8.5 ÑÏ 52-103"))
        NCResult = Max(NCResult, Vars.Result)
        Vars.Ex ("S_" & VN("5.12"))
        NCResult = Max(NCResult, Vars.Result)
        
        Exit Function
1:
        If Err.Number <> 56401 Then MsgBox "Error " & Err.Number & "; Description: " & Err.Description, vbCritical, "Error"
End Function
Private Function Max(A As Single, B As Single) As Single
        Max = IIf(A > B, A, B)
End Function
Private Function VN(Name As String) As String
        VN = Replace(Name, " ", "_spc_")
        VN = Replace(VN, "..", "_zpt_")
        VN = Replace(VN, ".", "_pnt_")
        VN = Replace(VN, "-", "_minus_")
        VN = Replace(VN, "(", "_bkt1_")
        VN = Replace(VN, ")", "_bkt2_")
End Function