Public Function Calc(Gp As Double, Gs As Double) As Boolean
    'Функция расчёта усилий для ферм

    ReportAdd "Расчёт усилий"

    Gpkr = Gp    ' расчётный вес покрытия
    Gstn = Gs    ' расчётный вес стен

    'Заполнение массивов данных для расчёта
    AddAllArr

    'Запуск расчёта через API NormFEM
    nfApi.Calc
    Calc = nfApi.Result  ' Проверка успешности расчёта

    If Not Calc Then Exit Function

    'Получение массивов результатов из API NormFEM
    nfApi.GetArrZ ArrZ   ' перемещения узлов
    nfApi.GetArrNM ArrNM ' нормальные силы и моменты
    nfApi.GetArrQ ArrQ   ' поперечные силы

    'Расчёт усилий от сочетаний нагрузок
    GetComb

    'Определение максимальных и минимальных усилий в элементах ферм
    MaxMinN

    'Освобождение объекта API
    Set nfApi = Nothing
End Function
