import win32com.client

# Создаём COM‑объект отчёта
nc_report = win32com.client.Dispatch("ncApi.Report")

# Выбираем нормативный документ и задание
nc_report.Norm = "EN 1991-1-3___Снеговые нагрузки"            # название модуля, как в NormCAD
nc_report.TaskName = "Определение нагрузки от нависания снега на краю ската покрытия"
nc_report.Unit = "1"                     # список пунктов; пустая строка = все пункты

# Загружаем модуль расчёта
nc_report.ClcLoadNorm()

# Способ 1: Загружаем исходные данные и условия из заранее сохранённых файлов
nc_report.LoadDat(r"C:\Users\servuser\Desktop\test_normcad\bim_normcad_check\self_made_vb_modules\Определение нагрузки от нависания снега на краю ската покрытия.dat")
nc_report.LoadNr1(r"C:\Users\servuser\Desktop\test_normcad\bim_normcad_check\self_made_vb_modules\Определение нагрузки от нависания снега на краю ската покрытия.nr1")
nc_report.ClcLoadData()
nc_report.ClcLoadConds()

# Способ 2 (альтернативно): создаём объект переменных и задаём значения в коде
# varlib = "NC_219532362921554E02"   # имя вашей библиотеки расчётного модуля
# vars_obj = win32com.client.Dispatch(f"{varlib}.Vars")
# vars_obj.b = 0.3
# vars_obj.h = 0.5
# ...
# nc_report.SetVars(vars_obj)

# Запускаем расчёт
nc_report.ClcCalc()

# Проверяем аппаратный ключ (важно для полной версии NormCAD)
if nc_report.TestKey():
    # Сохраняем полный отчёт в RTF или Word
    nc_report.MakeReport(r"C:\Users\servuser\Desktop\test_normcad\bim_normcad_check\self_made_vb_modules\test_report.rtf")
    # либо, если нужен отчёт с рамкой и штампом
    nc_report.SendToWord(r"C:\Users\servuser\Desktop\test_normcad\bim_normcad_check\self_made_vb_modules\test_report.doc")
else:
    print("Не найден аппаратный ключ NormCAD – отчёт не создан")
