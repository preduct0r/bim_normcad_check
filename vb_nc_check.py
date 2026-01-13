# pip install pywin32
import win32com.client

def run_normcad_check(norm, task_name, unit, vars_progid, vars_values: dict, conds=None):
    # 1) Create API object
    nc = win32com.client.Dispatch("ncApi.Report")  # :contentReference[oaicite:1]{index=1}

    # 2) Select module/task
    nc.Norm = norm
    nc.TaskName = task_name
    nc.Unit = unit

    nc.ClcLoadNorm  # иногда как свойство, иногда как метод (зависит от COM обертки)
    try:
        nc.ClcLoadNorm()
    except TypeError:
        pass

    # 3) Create Vars object of конкретного расчетного модуля
    Vars = win32com.client.Dispatch(vars_progid)  # пример: "NC_219532362921554E02.Vars" :contentReference[oaicite:2]{index=2}

    # 4) Fill input variables (имена полей берешь из “Создать VB проект”)
    for k, v in vars_values.items():
        setattr(Vars, k, v)

    nc.SetVars(Vars)  # :contentReference[oaicite:3]{index=3}

    # 5) Conditions (если надо)
    if conds:
        nc.SetConds(conds)  # :contentReference[oaicite:4]{index=4}

    # 6) Load data/conds and calculate
    nc.ClcLoadData()
    nc.ClcLoadConds()
    nc.ClcCalc()  # 

    max_util = float(nc.MaxResult)  # 
    return {"passed": max_util <= 1.0, "max_util": max_util}

# пример вызова (значения norm/task/unit/vars_progid и имена Vars.* — из сгенерированного VB-проекта)
# res = run_normcad_check("СП 64.13330...", "Балка", "1", "NC_....Vars", {"L": 5800, "b": 200, "h": 220, "q": 2.5})
# print(res)
