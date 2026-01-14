# -*- coding: utf-8 -*-
"""
Точный перевод normfem.vb на Python (1:1).

ОРИГИНАЛ VB:
    Public Function Calc(Gp As Double, Gs As Double) As Boolean
        ReportAdd "Расчёт усилий"
        Gpkr = Gp
        Gstn = Gs
        AddAllArr
        nfApi.Calc
        Calc = nfApi.Result
        If Not Calc Then Exit Function
        nfApi.GetArrZ ArrZ
        nfApi.GetArrNM ArrNM
        nfApi.GetArrQ ArrQ
        GetComb
        MaxMinN
        Set nfApi = Nothing
    End Function
"""

import win32com.client

# Глобальные переменные (как в VB)
nfApi = None      # объект API NormFEM
Gpkr = 0.0        # расчётный вес покрытия
Gstn = 0.0        # расчётный вес стен
ArrZ = None       # перемещения узлов
ArrNM = None      # нормальные силы и моменты
ArrQ = None       # поперечные силы


def ReportAdd(message):
    """VB: ReportAdd "текст" """
    print(message)


def AddAllArr():
    """VB: AddAllArr - заполнение массивов данных для расчёта"""
    # Заглушка - в реальном коде здесь передача данных через nfApi.SetArr
    pass


def GetComb():
    """VB: GetComb - расчёт усилий от сочетаний нагрузок"""
    pass


def MaxMinN():
    """VB: MaxMinN - определение max/min усилий в элементах ферм"""
    pass


def Calc(Gp, Gs):
    """
    VB: Public Function Calc(Gp As Double, Gs As Double) As Boolean
    
    Точный перевод строка за строкой.
    """
    global nfApi, Gpkr, Gstn, ArrZ, ArrNM, ArrQ
    
    # VB: ReportAdd "Расчёт усилий"
    ReportAdd("Расчёт усилий")
    
    # VB: Gpkr = Gp
    Gpkr = Gp
    
    # VB: Gstn = Gs
    Gstn = Gs
    
    # VB: AddAllArr
    AddAllArr()
    
    # VB: nfApi.Calc
    nfApi.Calc()
    
    # VB: Calc = nfApi.Result
    result = nfApi.Result
    
    # VB: If Not Calc Then Exit Function
    if not result:
        return result
    
    # VB: nfApi.GetArrZ ArrZ
    # В VB это ByRef - массив заполняется внутри метода
    # Вариант 1: передача как аргумент (если метод возвращает None и модифицирует аргумент)
    # Вариант 2: метод возвращает массив
    try:
        # Попытка 1: метод возвращает значение
        ArrZ = nfApi.GetArrZ()
    except:
        # Попытка 2: передаём пустой массив для заполнения
        ArrZ = []
        nfApi.GetArrZ(ArrZ)
    
    # VB: nfApi.GetArrNM ArrNM
    try:
        ArrNM = nfApi.GetArrNM()
    except:
        ArrNM = []
        nfApi.GetArrNM(ArrNM)
    
    # VB: nfApi.GetArrQ ArrQ
    try:
        ArrQ = nfApi.GetArrQ()
    except:
        ArrQ = []
        nfApi.GetArrQ(ArrQ)
    
    # VB: GetComb
    GetComb()
    
    # VB: MaxMinN
    MaxMinN()
    
    # VB: Set nfApi = Nothing
    nfApi = None
    
    # VB: End Function (возврат Calc)
    return result


if __name__ == "__main__":
    # Инициализация nfApi (должна быть ДО вызова Calc)
    # VB: Set nfApi = CreateObject("ncfem.main")
    nfApi = win32com.client.Dispatch("ncfem.main")
    
    # VB: nfApi.SetPath AppPath, TempDir, Project, ParentForm
    nfApi.SetPath("C:\\Program Files (x86)\\NormCAD", "C:\\Temp", "TestProject", None)
    
    # Тестовый вызов
    result = Calc(1.5, 2.0)
    
    print(f"Result: {result}")
    print(f"ArrZ: {ArrZ}")
    print(f"ArrNM: {ArrNM}")
    print(f"ArrQ: {ArrQ}")
