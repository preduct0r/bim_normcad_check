from __future__ import annotations

import struct
from pathlib import Path

import win32com.client


def _is_32bit_python() -> bool:
    return struct.calcsize("P") * 8 == 32


def _step(name: str, fn) -> None:
    try:
        fn()
        print(f"[OK] {name}")
    except Exception as e:
        msg = str(e).strip() or e.__class__.__name__
        raise RuntimeError(f"[FAIL] {name}: {msg}") from e


def main() -> int:
    if not _is_32bit_python():
        print("ERROR: NormCAD COM API requires 32-bit Python (run via env_32).")
        return 2

    module_dir = Path(__file__).resolve().parent
    dat_path = module_dir / "Определение нагрузки от нависания снега на краю ската покрытия.dat"
    nr1_path = module_dir / "Определение нагрузки от нависания снега на краю ската покрытия.nr1"
    out_rtf = module_dir / "test_report.rtf"
    out_doc = module_dir / "test_report.doc"

    for p in (dat_path, nr1_path):
        if not p.exists():
            print(f"ERROR: input file not found: {p}")
            return 3

    # Создаём COM‑объект отчёта
    # Per official docs (NCBkP.pdf p.53): Set ncApiR = New ncApi.Report
    nc_report = win32com.client.Dispatch("ncApi.Report")
    print("[OK] COM Dispatch(ncApi.Report)")

    # According to official docs, these are "variables" (properties):
    # - Norm: Module name (e.g., "СП 16.13330.2017___Стальные конструкции")
    # - TaskName: Task name within the module
    # - Unit: List of calculation sections (empty = all)
    norm_val = "EN 1991-1-3___Снеговые нагрузки"
    task_val = "Определение нагрузки от нависания снега на краю ската покрытия"
    # Unit specifies which calculation sections to run
    # From your .nr1 file: Unit=п.п. прил. C;6.3
    # Empty string may mean "no sections", not "all sections"
    unit_val = "п.п. прил. C;6.3"  # Specific sections from .nr1 file
    
    # In pywin32, COM properties without type library may appear as methods.
    # Try calling them as methods (VB property setters become method calls)
    def set_com_property(obj, prop_name, value):
        """Try setting COM property using various approaches."""
        # Check what type of attribute we're dealing with
        attr = getattr(obj, prop_name, None)
        print(f"[DEBUG] {prop_name} attribute type: {type(attr)}")
        
        # Approach 1: Call as method (most common for late-bound COM)
        if callable(attr):
            try:
                attr(value)
                print(f"[OK] {prop_name}({value!r}) - called as method")
                return True
            except Exception as e:
                print(f"[WARN] {prop_name}() as method failed: {e}")
        
        # Approach 2: Direct property assignment
        try:
            setattr(obj, prop_name, value)
            print(f"[OK] {prop_name} = {value!r} - property assignment")
            return True
        except Exception as e:
            print(f"[WARN] {prop_name} property assignment failed: {e}")
        
        return False
    
    set_com_property(nc_report, "Norm", norm_val)
    set_com_property(nc_report, "TaskName", task_val)
    set_com_property(nc_report, "Unit", unit_val)

    # Загружаем модуль расчёта
    _step("ClcLoadNorm()", lambda: nc_report.ClcLoadNorm())

    # Per docs (p.52): SetVars(Vars As Object) - Передает объект переменных
    # The .bas file uses: Set Vars = CreateObject("NC_873301143084689E03.Vars")
    # Try both approaches: LoadDat/LoadNr1 AND SetVars
    
    # Approach A: Load from files (current approach)
    _step(f"LoadDat({dat_path.name})", lambda: nc_report.LoadDat(str(dat_path)))
    _step(f"LoadNr1({nr1_path.name})", lambda: nc_report.LoadNr1(str(nr1_path)))
    
    # Approach B: Try to create and use Vars object directly
    # The ProgID is from the .bas file: NC_873301143084689E03.Vars
    vars_progid = "NC_873301143084689E03.Vars"
    try:
        vars_obj = win32com.client.Dispatch(vars_progid)
        print(f"[OK] Created Vars object: {vars_progid}")
        
        # Get conditions from Vars
        conds = vars_obj.Conds
        
        # Set variables (from .dat file)
        def VN(name):
            """Variable name transformation (same as in .bas)"""
            name = name.replace(" ", "_spc_")
            name = name.replace("..", "_zpt_")
            name = name.replace(".", "_pnt_")
            name = name.replace("-", "_minus_")
            name = name.replace("(", "_bkt1_")
            name = name.replace(")", "_bkt2_")
            return name
        
        # Set values from .dat file
        vars_obj[VN("C__t")].Value = 1
        vars_obj[VN("gr_a")].Value = 4
        vars_obj[VN("gr_g__Qi")].Value = 1.5
        vars_obj[VN("s__k")].Value = 2.577
        vars_obj[VN("Z")].Value = 4
        vars_obj[VN("A___A")].Value = 0.03
        print("[OK] Set variable values")
        
        # Add conditions (from .nr1 file)
        conds.Add("Покрытия - без повышенной теплоотдачи")
        conds.Add("Климатический регион - Альпийский регион")
        conds.Add("Условия местности - не защищенные от ветра")
        conds.Add("Форма кровли - односкатное покрытие")
        print("[OK] Added conditions")
        
        # Pass Vars object to Report
        nc_report.SetVars(vars_obj)
        print("[OK] SetVars(vars_obj)")
        
    except Exception as e:
        print(f"[INFO] Vars object approach failed: {e}")
        print("[INFO] Continuing with LoadDat/LoadNr1 approach only...")
    
    # Re-set Unit after loading (in case LoadNr1 overwrote it)
    print("[INFO] Re-setting Unit...")
    set_com_property(nc_report, "Unit", unit_val)
    
    _step("ClcLoadData()", lambda: nc_report.ClcLoadData())
    _step("ClcLoadConds()", lambda: nc_report.ClcLoadConds())
    
    # Per docs (p.53): LoadProp loads calculation parameters from registry
    try:
        nc_report.LoadProp()
        print("[OK] LoadProp()")
    except Exception as e:
        print(f"[INFO] LoadProp() not available or failed: {e}")

    # Запускаем расчёт
    _step("ClcCalc()", lambda: nc_report.ClcCalc())

    # Check if calculation produced any results
    try:
        max_result = nc_report.MaxResult
        print(f"[INFO] MaxResult => {max_result}")
    except Exception as e:
        print(f"[INFO] MaxResult not available: {e}")
    
    try:
        result = nc_report.Result
        print(f"[INFO] Result => {result}")
    except Exception as e:
        print(f"[INFO] Result not available: {e}")

    # Check MadeCalculations flag (from .nr1 file it was FALSE)
    try:
        made_calc = nc_report.MadeCalculations
        print(f"[INFO] MadeCalculations => {made_calc}")
    except Exception as e:
        print(f"[INFO] MadeCalculations not available: {e}")

    # Проверяем аппаратный ключ (важно для полной версии NormCAD)
    has_key = bool(nc_report.TestKey())
    print(f"[INFO] TestKey() => {has_key}")
    if not has_key:
        print("ERROR: hardware key/license not found – report may be blocked.")
        return 4

    # Сохраняем полный отчёт
    _step(f"MakeReport({out_rtf.name})", lambda: nc_report.MakeReport(str(out_rtf)))
    try:
        size = out_rtf.stat().st_size
    except Exception:
        size = -1
    print(f"[INFO] RTF size: {size} bytes")
    
    # Read and inspect RTF content
    try:
        with open(out_rtf, "rb") as f:
            rtf_content = f.read()
        # Show first 200 bytes to diagnose
        print(f"[INFO] RTF content preview: {rtf_content[:200]!r}")
        if 0 <= size < 300:
            print("WARNING: RTF is suspiciously small (likely empty report content).")
            print("         Most common cause: Unit is restricting sections or Calc didn't produce results.")
    except Exception as e:
        print(f"[WARNING] Could not read RTF: {e}")

    # Альтернатива: Word-отчёт с рамкой/штампом (требует Word/настроек)
    try:
        _step(f"SendToWord({out_doc.name})", lambda: nc_report.SendToWord(str(out_doc)))
        doc_size = out_doc.stat().st_size
        print(f"[INFO] DOC size: {doc_size} bytes")
    except Exception as e:
        # Word export is optional; don't fail the whole run.
        print(str(e))

    print(f"[DONE] Outputs: {out_rtf} ; {out_doc}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
