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


def _find_install_location_from_uninstall() -> Optional[str]:
    """
    Best-effort lookup of NormFEM install folder via Windows registry uninstall entries.
    Works for both 32/64-bit views.
    """
    try:
        import winreg
    except Exception:
        return None

    uninstall_roots = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]

    def _try_get_str(hkey, name: str) -> Optional[str]:
        try:
            v, t = winreg.QueryValueEx(hkey, name)
            if t in (winreg.REG_SZ, winreg.REG_EXPAND_SZ) and isinstance(v, str) and v.strip():
                return v.strip()
        except Exception:
            return None
        return None

    for hive, root in uninstall_roots:
        try:
            with winreg.OpenKey(hive, root) as k:
                i = 0
                while True:
                    try:
                        sub = winreg.EnumKey(k, i)
                    except OSError:
                        break
                    i += 1
                    try:
                        with winreg.OpenKey(k, sub) as sk:
                            name = _try_get_str(sk, "DisplayName") or ""
                            if "normfem" not in name.lower():
                                continue
                            loc = _try_get_str(sk, "InstallLocation")
                            if loc and os.path.isdir(loc):
                                return loc
                            # Sometimes only DisplayIcon is set (path to exe)
                            icon = _try_get_str(sk, "DisplayIcon")
                            if icon:
                                icon_path = icon.split(",")[0].strip().strip('"')
                                base = os.path.dirname(icon_path)
                                if base and os.path.isdir(base):
                                    return base
                    except Exception:
                        continue
        except Exception:
            continue
    return None


def _default_normfem_paths() -> List[str]:
    reg_loc = _find_install_location_from_uninstall() or ""
    # User-provided custom install example: C:\NormCAD\NormFEM.exe
    normcad_dir = r"C:\NormCAD"
    normcad_exe = os.path.join(normcad_dir, "NormFEM.exe")
    candidates = [
        os.environ.get("NORMFEM_PATH", ""),
        reg_loc,
        normcad_dir if os.path.isdir(normcad_dir) else "",
        normcad_dir if os.path.isfile(normcad_exe) else "",
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


def _normalize_normfem_app_path(p: Optional[str]) -> Optional[str]:
    """
    Accept either a directory (NormFEM app folder) or a direct path to NormFEM.exe.
    Return a directory path or None.
    """
    if not p:
        return None
    p = p.strip().strip('"').strip()
    if not p:
        return None
    try:
        if os.path.isfile(p):
            return os.path.dirname(p) or None
        if os.path.isdir(p):
            return p
    except Exception:
        return None
    return None


def _print_table(title: str, rows: List[CheckResult]) -> None:
    print("\n" + title)
    print("-" * len(title))
    width = max([len(r.name) for r in rows] + [10])
    for r in rows:
        status = _fmt_bool(r.ok)
        details = f" - {r.details}" if r.details else ""
        print(f"{r.name:<{width}}  {status}{details}")


def _read_text_lines(path: str) -> List[str]:
    """
    NormFEM tables are typically line-oriented text files. We read them robustly
    because many installs use ANSI/Windows-1251.
    """
    # Try UTF-8 first, then fall back to CP1251 (common for RU installations).
    for enc in ("utf-8", "cp1251"):
        try:
            with open(path, "r", encoding=enc, errors="strict") as f:
                return [ln.rstrip("\r\n") for ln in f.readlines()]
        except UnicodeDecodeError:
            continue
        except Exception:
            break
    with open(path, "r", encoding="cp1251", errors="replace") as f:
        return [ln.rstrip("\r\n") for ln in f.readlines()]


def _default_example_dirs(app_path: Optional[str]) -> List[str]:
    """
    Common places where NormCAD ships example projects/tables.
    """
    candidates = [
        r"C:\NormCAD\NormFEM\Exs\РСУ СП",
        r"C:\NormCAD\Exs\РСУ СП",
        r"C:\NormCAD\NormFEM\Exs",
        r"C:\NormCAD\Exs",
    ]
    if app_path:
        ap = app_path.rstrip("\\/")
        candidates = [
            os.path.join(ap, "NormFEM", "Exs", "РСУ СП"),
            os.path.join(ap, "Exs", "РСУ СП"),
            os.path.join(ap, "NormFEM", "Exs"),
            os.path.join(ap, "Exs"),
        ] + candidates
    out: List[str] = []
    for p in candidates:
        if p and p not in out:
            out.append(p)
    return out


def _pick_example_dir(app_path: Optional[str], example_dir: Optional[str]) -> Optional[str]:
    if example_dir:
        return example_dir if os.path.isdir(example_dir) else None
    for p in _default_example_dirs(app_path):
        try:
            if os.path.isdir(p):
                return p
        except Exception:
            continue
    return None


def _iter_example_tables(example_dir: str, max_tables: int = 80) -> List[Tuple[str, str, List[str]]]:
    """
    Returns list of (table_name, file_path, lines).
    Table name is derived from file extension (e.g. RSU.G01 -> 'g01').
    """
    items: List[Tuple[str, str, List[str]]] = []
    for name in sorted(os.listdir(example_dir)):
        path = os.path.join(example_dir, name)
        if not os.path.isfile(path):
            continue
        _, ext = os.path.splitext(name)
        if not ext:
            continue
        tbl = ext.lstrip(".").strip().lower()
        # Skip obviously non-table/binary stuff
        if tbl in {"exe", "dll", "ocx", "png", "jpg", "jpeg", "gif", "ico", "dwg", "bak"}:
            continue
        try:
            if os.path.getsize(path) > 2_000_000:
                continue
        except Exception:
            continue
        try:
            lines = _read_text_lines(path)
        except Exception:
            continue
        items.append((tbl, path, lines))
        if len(items) >= max_tables:
            break
    return items


def _make_parent_form(mode: str):
    """
    NormFEM VB examples pass a Form instance to SetPath (ParentForm:=Me) that contains
    txtProgress and txtReport controls. Some builds appear to require a non-null
    ParentForm and will crash in Calc() if it's missing.
    """
    mode = (mode or "none").strip().lower()
    if mode in {"none", "null"}:
        return None
    if mode in {"dummy", "form"}:
        # Expose a minimal IDispatch object with txtProgress/txtReport having .Text.
        from win32com.server.util import wrap

        class _TextBox:
            _public_methods_: List[str] = []
            _public_attrs_ = ["Text"]

            def __init__(self):
                self.Text = ""

        class _Form:
            _public_methods_: List[str] = []
            _public_attrs_ = ["txtProgress", "txtReport"]

            def __init__(self):
                self.txtProgress = wrap(_TextBox())
                self.txtReport = wrap(_TextBox())

        return wrap(_Form())
    if mode in {"dict", "scripting.dictionary"}:
        # Some COM code only checks for object presence; this is a cheap non-null IDispatch.
        return _dispatch("Scripting.Dictionary")
    raise ValueError(f"Unknown --parent-form mode: {mode!r}")


def check_normfem(
    *,
    app_path: Optional[str],
    temp_dir: Optional[str],
    project: str,
    try_calc: bool,
    example_dir: Optional[str],
    use_example: bool,
    parent_form_mode: str,
    call_prop: bool,
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

    chosen_app_path = _normalize_normfem_app_path(app_path) or _pick_existing_dir(_default_normfem_paths())
    chosen_temp_dir = temp_dir or tempfile.gettempdir()
    parent_form = _make_parent_form(parent_form_mode)

    def step_setpath():
        # In VB docs: nfApi.SetPath AppPath:="...", TempDir:="...", Project:="...", ParentForm:=Me
        # In Python we don't have a UI form; passing None typically works for COM expecting an object pointer.
        if not chosen_app_path:
            raise RuntimeError(
                "NormFEM install folder was not detected. "
                "Pass --app-path \"C:\\Path\\To\\NormFEM\" or set NORMFEM_PATH."
            )
        obj.SetPath(chosen_app_path, chosen_temp_dir, project, parent_form)

    results.append(
        _run_step(
            f'SetPath(app_path="{chosen_app_path}", temp_dir="{chosen_temp_dir}", project="{project}", parent_form="{parent_form_mode}")',
            step_setpath,
        )
    )
    if results[-1].ok is not True:
        return results

    # Prefer a real shipped example dataset if present: this makes Calc() meaningful.
    if use_example:
        picked = _pick_example_dir(chosen_app_path, example_dir)
        if picked:
            tables = _iter_example_tables(picked)
            results.append(_run_step(f'Load example dir "{picked}"', lambda: None if tables else (_ for _ in ()).throw(RuntimeError("No readable tables found"))))
            for tbl, path, lines in tables:
                results.append(_run_step(f'SetArr("{tbl}", from {os.path.basename(path)})', lambda t=tbl, arr=lines: obj.SetArr(t, arr)))
        else:
            results.append(
                CheckResult(
                    name="Example dataset",
                    ok=False,
                    details="Not found (pass --example-dir to point to NormFEM example folder)",
                )
            )
    else:
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
        if call_prop:
            results.append(_run_step("Prop()", lambda: obj.Prop()))
        calc_res = _run_step("Calc()", lambda: obj.Calc())
        results.append(calc_res)

        if calc_res.ok is True:
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
    ap.add_argument("--example-dir", default="", help="Folder with NormFEM example tables to load (optional)")
    ap.add_argument("--no-example", action="store_true", help="Do not try to load shipped example tables; only do minimal SetArr checks")
    ap.add_argument(
        "--parent-form",
        default="none",
        help='ParentForm for nfApi.SetPath: "none" (default), "dummy" (recommended), or "dict".',
    )
    ap.add_argument("--call-prop", action="store_true", help="Call nfApi.Prop() before Calc() (some builds may require it)")
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
    example_dir = args.example_dir.strip() or None

    try:
        nf_rows = check_normfem(
            app_path=app_path,
            temp_dir=temp_dir,
            project=args.project,
            try_calc=(not args.no_calc),
            example_dir=example_dir,
            use_example=(not args.no_example),
            parent_form_mode=args.parent_form,
            call_prop=args.call_prop,
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

