#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Python port of: Определение нагрузки от нависания снега на краю ската покрытия.bas

This script:
1. Performs the same calculation as the VB module using NC_873301143084689E03.Vars
2. Generates a full report using ncApi.Report API

Based on official NormCAD API documentation (NCBkP.pdf p.52-53)

Run with 32-bit Python:
  C:\\Users\\servuser\\Desktop\\test_normcad\\env_32\\Scripts\\python.exe snow_overhang_calc.py
"""

from __future__ import annotations

import struct
from pathlib import Path
from dataclasses import dataclass
from typing import Optional

import win32com.client


def _is_32bit_python() -> bool:
    return struct.calcsize("P") * 8 == 32


def VN(name: str) -> str:
    """
    Variable name transformation (same as VN function in .bas file).
    Converts special characters to safe identifiers.
    """
    name = name.replace(" ", "_spc_")
    name = name.replace("..", "_zpt_")
    name = name.replace(".", "_pnt_")
    name = name.replace("-", "_minus_")
    name = name.replace("(", "_bkt1_")
    name = name.replace(")", "_bkt2_")
    return name


@dataclass
class SnowLoadInput:
    """Input parameters for snow overhang load calculation."""
    # Temperature coefficient
    C_t: float = 1.0
    # Roof slope angle (degrees)
    gr_a: float = 4.0
    # Reliability coefficient for temporary load
    gr_g_Qi: float = 1.5
    # Characteristic snow load on ground (kPa)
    s_k: float = 1.935
    # Snow zone number
    Z: int = 3
    # Altitude above sea level (m)
    A_A: float = 0.04
    
    # Conditions (from .nr1 file)
    condition_thermal: str = "Покрытия - без повышенной теплоотдачи"
    condition_climate: str = "Климатический регион - Альпийский регион"
    condition_wind: str = "Условия местности - не защищенные от ветра"
    condition_roof: str = "Форма кровли - односкатное покрытие"


@dataclass
class SnowLoadResult:
    """Results from snow overhang load calculation."""
    max_result: float  # Maximum utilization coefficient
    section_results: dict  # Results for each calculation section


class SnowOverhangCalculator:
    """
    Calculator for snow overhang load at roof edge.
    Based on EN 1991-1-3 (Eurocode snow loads).
    """
    
    # COM ProgID for this specific calculation module
    VARS_PROGID = "NC_873301143084689E03.Vars"
    
    # NormCAD module identification
    NORM = "EN 1991-1-3___Снеговые нагрузки"
    TASK_NAME = "Определение нагрузки от нависания снега на краю ската покрытия"
    UNIT = "п.п. прил. C;6.3"  # Calculation sections
    
    def __init__(self):
        self.vars_obj = None
        self.conds = None
        self.report_obj = None
    
    def calculate(self, input_data: SnowLoadInput) -> SnowLoadResult:
        """
        Perform the calculation using direct Vars object (like in .bas file).
        Returns the maximum utilization coefficient.
        """
        # Create Vars object
        self.vars_obj = win32com.client.Dispatch(self.VARS_PROGID)
        self.conds = self.vars_obj.Conds
        
        # Set input variables
        self.vars_obj[VN("C__t")].Value = input_data.C_t
        self.vars_obj[VN("gr_a")].Value = input_data.gr_a
        self.vars_obj[VN("gr_g__Qi")].Value = input_data.gr_g_Qi
        self.vars_obj[VN("s__k")].Value = input_data.s_k
        self.vars_obj[VN("Z")].Value = input_data.Z
        self.vars_obj[VN("A___A")].Value = input_data.A_A
        
        # Add conditions
        self.conds.Add(input_data.condition_thermal)
        self.conds.Add(input_data.condition_climate)
        self.conds.Add(input_data.condition_wind)
        self.conds.Add(input_data.condition_roof)
        
        # Execute calculations (same as .bas file)
        self.vars_obj.Result = 0
        section_results = {}
        max_result = 0.0
        
        # Section: прил. C (Appendix C)
        self.vars_obj.Ex("S_" + VN("прил. C"))
        result_c = float(self.vars_obj.Result)
        section_results["прил. C"] = result_c
        max_result = max(max_result, result_c)
        
        # Section: 6.3
        self.vars_obj.Ex("S_" + VN("6.3"))
        result_63 = float(self.vars_obj.Result)
        section_results["6.3"] = result_63
        max_result = max(max_result, result_63)
        
        return SnowLoadResult(
            max_result=max_result,
            section_results=section_results
        )
    
    def generate_report(
        self,
        input_data: SnowLoadInput,
        dat_file: Path,
        nr1_file: Path,
        output_rtf: Optional[Path] = None,
        output_doc: Optional[Path] = None,
    ) -> bool:
        """
        Generate full calculation report using ncApi.Report.
        
        This matches the working report_example.py exactly:
        1. Load data from .dat and .nr1 files
        2. Create Vars object, set values, call SetVars()
        3. Call LoadProp() before ClcCalc()
        
        Args:
            input_data: Calculation input parameters
            dat_file: Path to .dat file with input data (required)
            nr1_file: Path to .nr1 file with conditions (required)
            output_rtf: Path for RTF report (MakeReport)
            output_doc: Path for Word report with stamp (SendToWord)
            
        Returns:
            True if report generated successfully
        """
        # Create Report object
        self.report_obj = win32com.client.Dispatch("ncApi.Report")
        
        # Set module identification (must be set BEFORE ClcLoadNorm)
        self.report_obj.Norm = self.NORM
        self.report_obj.TaskName = self.TASK_NAME
        self.report_obj.Unit = self.UNIT
        
        # Load calculation module
        self.report_obj.ClcLoadNorm()
        print("[OK] ClcLoadNorm()")
        
        # Load data from files
        self.report_obj.LoadDat(str(dat_file))
        print(f"[OK] LoadDat({dat_file.name})")
        
        self.report_obj.LoadNr1(str(nr1_file))
        print(f"[OK] LoadNr1({nr1_file.name})")
        
        # Create and configure Vars object (like report_example.py)
        vars_obj = win32com.client.Dispatch(self.VARS_PROGID)
        conds = vars_obj.Conds
        
        # Set input variables
        vars_obj[VN("C__t")].Value = input_data.C_t
        vars_obj[VN("gr_a")].Value = input_data.gr_a
        vars_obj[VN("gr_g__Qi")].Value = input_data.gr_g_Qi
        vars_obj[VN("s__k")].Value = input_data.s_k
        vars_obj[VN("Z")].Value = input_data.Z
        vars_obj[VN("A___A")].Value = input_data.A_A
        print("[OK] Set variable values")
        
        # Add conditions
        conds.Add(input_data.condition_thermal)
        conds.Add(input_data.condition_climate)
        conds.Add(input_data.condition_wind)
        conds.Add(input_data.condition_roof)
        print("[OK] Added conditions")
        
        # Pass Vars to Report
        self.report_obj.SetVars(vars_obj)
        print("[OK] SetVars(vars_obj)")
        
        # Re-set Unit after loading files (LoadNr1 may override it)
        self.report_obj.Unit = self.UNIT
        print(f"[OK] Unit = {self.UNIT!r}")
        
        # Load data and conditions into module
        self.report_obj.ClcLoadData()
        print("[OK] ClcLoadData()")
        
        self.report_obj.ClcLoadConds()
        print("[OK] ClcLoadConds()")
        
        # Load calculation properties from registry (like report_example.py)
        try:
            self.report_obj.LoadProp()
            print("[OK] LoadProp()")
        except Exception as e:
            print(f"[INFO] LoadProp() not available: {e}")
        
        # Run calculation
        self.report_obj.ClcCalc()
        print("[OK] ClcCalc()")
        
        # Check license
        if not self.report_obj.TestKey():
            print("ERROR: Hardware key not found - report may be incomplete")
            return False
        print("[OK] TestKey() = True")
        
        # Generate reports
        if output_rtf:
            self.report_obj.MakeReport(str(output_rtf))
            size = output_rtf.stat().st_size
            print(f"[OK] MakeReport -> {output_rtf.name} ({size} bytes)")
            if size < 500:
                print("[WARNING] RTF file is suspiciously small!")
        
        if output_doc:
            self.report_obj.SendToWord(str(output_doc))
            size = output_doc.stat().st_size
            print(f"[OK] SendToWord -> {output_doc.name} ({size} bytes)")
        
        return True


def main() -> int:
    if not _is_32bit_python():
        print("ERROR: NormCAD COM API requires 32-bit Python.")
        print("Run with: C:\\Users\\servuser\\Desktop\\test_normcad\\env_32\\Scripts\\python.exe")
        return 2
    
    # File paths
    module_dir = Path(__file__).resolve().parent
    dat_file = module_dir / "Определение нагрузки от нависания снега на краю ската покрытия.dat"
    nr1_file = module_dir / "Определение нагрузки от нависания снега на краю ската покрытия.nr1"
    output_rtf = module_dir / "snow_overhang_report.rtf"
    output_doc = module_dir / "snow_overhang_report.doc"
    
    # Check required files exist
    for f in (dat_file, nr1_file):
        if not f.exists():
            print(f"ERROR: Required file not found: {f}")
            return 3
    
    # Input data (from .bas file values)
    input_data = SnowLoadInput(
        C_t=1.0,
        gr_a=4.0,
        gr_g_Qi=1.5,
        s_k=1.935,  # kPa - characteristic snow load
        Z=3,        # Snow zone
        A_A=0.04,   # Altitude in meters (0.04 = 4 cm above sea level)
    )
    
    print("=" * 60)
    print("Snow Overhang Load Calculation")
    print("EN 1991-1-3 (Eurocode)")
    print("=" * 60)
    
    print("\n--- Input Data ---")
    print(f"  Temperature coefficient C_t = {input_data.C_t}")
    print(f"  Roof slope angle α = {input_data.gr_a}°")
    print(f"  Reliability coefficient γ_Qi = {input_data.gr_g_Qi}")
    print(f"  Characteristic snow load s_k = {input_data.s_k} kPa")
    print(f"  Snow zone Z = {input_data.Z}")
    print(f"  Altitude A = {input_data.A_A} m")
    
    print("\n--- Conditions ---")
    print(f"  - {input_data.condition_thermal}")
    print(f"  - {input_data.condition_climate}")
    print(f"  - {input_data.condition_wind}")
    print(f"  - {input_data.condition_roof}")
    
    # Create calculator
    calc = SnowOverhangCalculator()
    
    # Perform calculation (like .bas file)
    print("\n--- Calculation ---")
    try:
        result = calc.calculate(input_data)
        print(f"[OK] Calculation completed")
        print(f"  Section 'прил. C' result: {result.section_results.get('прил. C', 'N/A')}")
        print(f"  Section '6.3' result: {result.section_results.get('6.3', 'N/A')}")
        print(f"  Maximum utilization coefficient: {result.max_result}")
    except Exception as e:
        print(f"[FAIL] Calculation error: {e}")
        return 1
    
    # Generate reports (matching report_example.py exactly)
    print("\n--- Report Generation ---")
    try:
        success = calc.generate_report(
            input_data=input_data,
            dat_file=dat_file,
            nr1_file=nr1_file,
            output_rtf=output_rtf,
            output_doc=output_doc,
        )
        if success:
            print(f"\n[SUCCESS] Reports generated:")
            print(f"  RTF: {output_rtf}")
            print(f"  DOC: {output_doc}")
        else:
            print("[WARNING] Report generation may be incomplete")
    except Exception as e:
        print(f"[FAIL] Report generation error: {e}")
        return 1
    
    print("\n" + "=" * 60)
    print(f"Final Result: NCResult = {result.max_result}")
    print("=" * 60)
    
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
