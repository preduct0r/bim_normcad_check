Attribute VB_Name = "Определение_нагрузки_от_нависания_снега_на_краю_ската_покрытия"
Option Explicit
Public Function NCResult() As Single
	On Error GoTo 1
	Dim Vars As Object
	Dim Conds As Object
	Set Vars = CreateObject("NC_873301143084689E03.Vars")
	Set Conds = Vars.Conds
	
	Vars(VN("C__t")).Value = 1
	Vars(VN("gr_a")).Value = 4
	Vars(VN("gr_g__Qi")).Value = 1.5
	Vars(VN("s__k")).Value = 1.935
	Vars(VN("Z")).Value = 3
	Vars(VN("A___A")).Value = 0.04
	
	Conds.Add "Покрытия - без повышенной теплоотдачи"
	Conds.Add "Климатический регион - Альпийский регион"
	Conds.Add "Условия местности - не защищенные от ветра"
	Conds.Add "Форма кровли - односкатное покрытие"
	
	Vars.Result = 0
	Vars.Ex("S_" & VN("прил. C"))
	NCResult = Max(NCResult, Vars.Result)
	Vars.Ex("S_" & VN("6.3"))
	NCResult = Max(NCResult, Vars.Result)
	
	Exit Function
1:
	If Err.Number <> 56401 Then MsgBox "Error " & Err.Number & "; Description: " & Err.Description, vbCritical, "Error"
End Function
Private Function Max(A As Single, B As Single) As Single
	Max =  IIf(A > B, A, B)
End Function
Private Function VN(Name As String) As String
	VN = Replace(Name, " ", "_spc_")
	VN = Replace(VN, "..", "_zpt_")
	VN = Replace(VN, ".", "_pnt_")
	VN = Replace(VN, "-", "_minus_")
	VN = Replace(VN, "(", "_bkt1_")
	VN = Replace(VN, ")", "_bkt2_")
End Function
