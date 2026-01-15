#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
API availability checker for NormCAD / NormFEM (COM).

Focus: NormFEM API used in `ground_truth/normfem.vb`:
  - create COM object
  - SetPath(...)
  - SetArr(...)
  - Calc()
  - (optional) GetArrZ/GetArrNM/GetArrQ

Run with 32-bit Python (required for these COM components):
  C:\\Users\\servuser\\Desktop\\test_normcad\\env_32\\Scripts\\python.exe checks\\api_check_normfem.py
"""

from __future__ import annotations

import argparse
import os
import platform
import struct
import sys
import tempfile
import traceback
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Tuple


def _is_32bit_python() -> bool:
    return struct.calcsize("P") * 8 == 32


def _fmt_bool(v: Optional[bool]) -> str:
    if v is True:
        return "OK"
    if v is False:
        return "FAIL"
    return "-"


@dataclass
class CheckResult:
    name: str
    ok: Optional[bool]
    details: str = ""


def _run_step(name: str, fn: Callable[[], Any]) -> CheckResult:
    try:
        fn()
        return CheckResult(name=name, ok=True, details="")
    except Exception as e:
        # Keep the exception message short but useful
        msg = str(e).strip()
        if not msg:
            msg = e.__class__.__name__
        return CheckResult(name=name, ok=False, details=msg)


def _try_import_win32com() -> Tuple[bool, Optional[str]]:
    try:
        import win32com.client  # noqa: F401
        return True, None
    except Exception as e:
        return False, f"{e.__class__.__name__}: {e}"


def _dispatch(progid: str):
    import win32com.client
    return win32com.client.Dispatch(progid)


def _default_normfem_paths() -> List[str]:
    candidates = [
        os.environ.get("NORMFEM_PATH", ""),
        r"C:\Program Files\NormFEM",
        r"C:\Program Files (x86)\NormFEM",
        r"C:\NormFEM",
    ]
    out: List[str] = []
    for p in candidates:
        p = (p or "").strip()
        if p and p not in out:
            out.append(p)
    return out


def _pick_existing_dir(paths: List[str]) -> Optional[str]:
    for p in paths:
        try:
            if p and os.path.isdir(p):
                return p
        except Exception:
            continue
    return None


def _print_table(title: str, rows: List[CheckResult]) -> None:
    print("\n" + title)
    print("-" * len(title))
    width = max([len(r.name) for r in rows] + [10])
    for r in rows:
        status = _fmt_bool(r.ok)
        details = f" - {r.details}" if r.details else ""
        print(f"{r.name:<{width}}  {status}{details}")


def check_normfem(
    *,
    app_path: Optional[str],
    temp_dir: Optional[str],
    project: str,
    try_calc: bool,
) -> List[CheckResult]:
    import win32com.client  # type: ignore

    results: List[CheckResult] = []

    nf: Dict[str, Any] = {"obj": None}

    def step_create():
        # Common ProgID seen in field tests for NormFEM
        nf["obj"] = win32com.client.Dispatch("ncfem.main")

    results.append(_run_step("COM object ncfem.main", step_create))
    if results[-1].ok is not True:
        return results

    obj = nf["obj"]

    chosen_app_path = app_path or _pick_existing_dir(_default_normfem_paths())
    chosen_temp_dir = temp_dir or tempfile.gettempdir()

    def step_setpath():
        # In VB docs: nfApi.SetPath AppPath:="...", TempDir:="...", Project:="...", ParentForm:=Me
        # In Python we don't have a UI form; passing None typically works for COM expecting an object pointer.
        obj.SetPath(chosen_app_path, chosen_temp_dir, project, None)

    results.append(
        _run_step(
            f'SetPath(app_path="{chosen_app_path}", temp_dir="{chosen_temp_dir}", project="{project}")',
            step_setpath,
        )
    )

    # Minimal SetArr checks: we don't assume strict table row format; this is *availability* testing.
    # We try empty list first (best-case), then a harmless placeholder row if the API rejects empty arrays.
    def setarr(tbl: str):
        try:
            obj.SetArr(tbl, [])
        except Exception:
            obj.SetArr(tbl, [""])

    for tbl in ("m00", "g01", "g02", "g03", "l00", "s00"):
        results.append(_run_step(f'SetArr("{tbl}", ...)', lambda t=tbl: setarr(t)))

    if try_calc:
        results.append(_run_step("Calc()", lambda: obj.Calc()))

        # If Calc succeeded, attempt to read arrays (best-effort)
        def _try_getarr(method_name: str):
            # COM methods with ByRef arrays: pywin32 can often pass a dummy list and it gets replaced.
            arr: Any = []
            getattr(obj, method_name)(arr)

        results.append(_run_step("GetArrZ()", lambda: _try_getarr("GetArrZ")))
        results.append(_run_step("GetArrNM()", lambda: _try_getarr("GetArrNM")))
        results.append(_run_step("GetArrQ()", lambda: _try_getarr("GetArrQ")))

    return results


def check_normcad() -> List[CheckResult]:
    import win32com.client  # type: ignore

    results: List[CheckResult] = []
    nc: Dict[str, Any] = {"obj": None}

    def step_create():
        nc["obj"] = win32com.client.Dispatch("ncApi.Report")

    results.append(_run_step('COM object ncApi.Report', step_create))
    if results[-1].ok is not True:
        return results

    obj = nc["obj"]

    # License check is a good proxy for a "real" installation.
    results.append(_run_step("TestKey()", lambda: bool(obj.TestKey())))
    return results


def main() -> int:
    ap = argparse.ArgumentParser(description="Check availability of NormFEM/NormCAD COM APIs")
    ap.add_argument("--app-path", default="", help="NormFEM installation folder (optional)")
    ap.add_argument("--temp-dir", default="", help="Temp folder for NormFEM (optional)")
    ap.add_argument("--project", default="ApiCheckProject", help="Project name for NormFEM SetPath")
    ap.add_argument("--no-calc", action="store_true", help="Skip Calc()/GetArr* calls (license may block Calc)")
    ap.add_argument("--verbose", action="store_true", help="Print full exception tracebacks")
    args = ap.parse_args()

    print("Environment")
    print("-----------")
    print(f"python: {sys.executable}")
    print(f"python_version: {platform.python_version()}")
    print(f"python_bits: {struct.calcsize('P') * 8}")
    print(f"os: {platform.platform()}")

    if not _is_32bit_python():
        print("\nERROR: NormCAD/NormFEM COM APIs require 32-bit Python. Please run with env_32.")
        return 2

    has_win32com, win32com_err = _try_import_win32com()
    if not has_win32com:
        print(f"\nERROR: pywin32/win32com is not available: {win32com_err}")
        return 3

    app_path = args.app_path.strip() or None
    temp_dir = args.temp_dir.strip() or None

    try:
        nf_rows = check_normfem(
            app_path=app_path,
            temp_dir=temp_dir,
            project=args.project,
            try_calc=(not args.no_calc),
        )
        _print_table("NormFEM API checks", nf_rows)

        nc_rows = check_normcad()
        _print_table("NormCAD API checks", nc_rows)

        # Exit code policy: fail if we couldn't create either COM object.
        fatal = any(
            (r.name.startswith("COM object") and r.ok is False)
            for r in (nf_rows + nc_rows)
        )
        return 1 if fatal else 0

    except Exception:
        print("\nUNHANDLED ERROR")
        print("--------------")
        if args.verbose:
            traceback.print_exc()
        else:
            print("Run with --verbose for full traceback.")
            print(traceback.format_exc().strip().splitlines()[-1])
        return 10


if __name__ == "__main__":
    raise SystemExit(main())

