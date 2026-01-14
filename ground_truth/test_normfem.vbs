' Тестовый VBScript для проверки API NormFEM с данными
' Запуск: C:\Windows\SysWOW64\cscript.exe //NoLogo test_normfem.vbs

Option Explicit

Dim nfApi
Dim result
Dim arrM00, arrG01, arrG02, arrG03, arrL00
Dim ArrZ, ArrNM, ArrQ

WScript.Echo "======================================"
WScript.Echo "ТЕСТ API NormFEM (VBScript + данные)"
WScript.Echo "======================================"

' Создание COM объекта
On Error Resume Next

WScript.Echo ""
WScript.Echo "[INFO] Создание COM объекта ncfem.main..."
Set nfApi = CreateObject("ncfem.main")

If Err.Number <> 0 Then
    WScript.Echo "[FAIL] Ошибка: " & Err.Description
    WScript.Quit 1
End If
WScript.Echo "[OK] COM объект создан"

' SetPath
WScript.Echo ""
WScript.Echo "[INFO] SetPath..."
nfApi.SetPath "C:\Program Files (x86)\NormCAD", "C:\Temp", "TestProject", Nothing

If Err.Number <> 0 Then
    WScript.Echo "[FAIL] SetPath: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] SetPath выполнен"
End If

' ========== ПЕРЕДАЧА ДАННЫХ ==========
WScript.Echo ""
WScript.Echo "[INFO] Передача данных через SetArr..."

' m00 - Материалы
ReDim arrM00(0)
arrM00(0) = "Steel; 2.0e5; 0.3"
nfApi.SetArr "m00", arrM00
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] m00: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] m00 (материалы): 1 шт"
End If

' g01 - Узлы
ReDim arrG01(1)
arrG01(0) = "1; 0; 0; 0"
arrG01(1) = "2; 3000; 0; 0"
nfApi.SetArr "g01", arrG01
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] g01: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] g01 (узлы): 2 шт"
End If

' g02 - Элементы
ReDim arrG02(0)
arrG02(0) = "1; 1; 2; 1; 40000; 1.33e8; 1.33e8; 2.67e8"
nfApi.SetArr "g02", arrG02
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] g02: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] g02 (элементы): 1 шт"
End If

' g03 - Связи
ReDim arrG03(1)
arrG03(0) = "1; 1; 1; 1; 0; 0; 0"
arrG03(1) = "2; 0; 1; 1; 0; 0; 0"
nfApi.SetArr "g03", arrG03
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] g03: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] g03 (связи): 2 шт"
End If

' l00 - Нагрузки
ReDim arrL00(0)
arrL00(0) = "1; 1; 0; 0; -1.5; 0; 0; 0"
nfApi.SetArr "l00", arrL00
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] l00: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] l00 (нагрузки): 1 шт"
End If

' ========== РАСЧЁТ ==========
WScript.Echo ""
WScript.Echo "[INFO] Вызов Calc..."
nfApi.Calc

If Err.Number <> 0 Then
    WScript.Echo "[FAIL] Calc: " & Err.Description & " (код " & Err.Number & ")"
    Err.Clear
Else
    WScript.Echo "[OK] Calc выполнен"
End If

' Result
WScript.Echo ""
WScript.Echo "[INFO] Проверка Result..."
result = nfApi.Result

If Err.Number <> 0 Then
    WScript.Echo "[FAIL] Result: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] Result = " & result
End If

' ========== ПОЛУЧЕНИЕ РЕЗУЛЬТАТОВ ==========
WScript.Echo ""
WScript.Echo "[INFO] Получение результатов..."

nfApi.GetArrZ ArrZ
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] GetArrZ: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] ArrZ получен"
End If

nfApi.GetArrNM ArrNM
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] GetArrNM: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] ArrNM получен"
End If

nfApi.GetArrQ ArrQ
If Err.Number <> 0 Then
    WScript.Echo "[FAIL] GetArrQ: " & Err.Description
    Err.Clear
Else
    WScript.Echo "[OK] ArrQ получен"
End If

' Освобождение
Set nfApi = Nothing

WScript.Echo ""
WScript.Echo "======================================"
WScript.Echo "ТЕСТ ЗАВЕРШЁН"
WScript.Echo "======================================"
