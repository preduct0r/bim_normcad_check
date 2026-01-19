#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Порт `self_made_vb_modules/1.vb` на Python (COM NormCAD Vars).

Оригинальный VB делает:
- CreateObject("NC_137667756294139E02.Vars")
- заполняет набор Vars(...)
- добавляет Conds.Add(...) (набор условий)
- выполняет ряд проверок Vars.Ex("S_"+VN(...))
- возвращает максимум коэффициента (NCResult).

Требования:
- Windows
- установленный NormCAD с зарегистрированным COM модулем
- Python 32-bit + pywin32 (в проекте: env_32)
"""

from __future__ import annotations

import argparse
import os
import struct
import sys
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional, Sequence, Tuple


def _is_32bit_python() -> bool:
    return struct.calcsize("P") * 8 == 32


def _fix_mojibake_cp1251(s: str) -> str:
    """
    В VB файлах RU-строки и имена переменных иногда видны как 'Àðì...'
    (CP1251 → Latin-1). Конвертируем обратно для корректной передачи в COM.
    """
    try:
        return s.encode("latin-1", errors="strict").decode("cp1251", errors="strict")
    except Exception:
        return s


def _s(x: str) -> str:
    return _fix_mojibake_cp1251(x)


def VN(name: str) -> str:
    """
    Полная копия VN из `self_made_vb_modules/1.vb`:
    - пробел -> _spc_
    - '..' -> _zpt_   (важно: сначала двойная точка)
    - '.' -> _pnt_
    - '-' -> _minus_
    - '(' -> _bkt1_
    - ')' -> _bkt2_
    """
    name = name.replace(" ", "_spc_")
    name = name.replace("..", "_zpt_")
    name = name.replace(".", "_pnt_")
    name = name.replace("-", "_minus_")
    name = name.replace("(", "_bkt1_")
    name = name.replace(")", "_bkt2_")
    return name


def _dispatch(progid: str):
    import win32com.client

    return win32com.client.Dispatch(progid)


def _max(a: float, b: float) -> float:
    return a if a > b else b


@dataclass(frozen=True)
class Params:
    # COM модуль (как в VB)
    progid: str = "NC_137667756294139E02.Vars"

    # Vars(...) значения (как в VB)
    gr_g__b1: float = 1.0
    m__kp: float = 1.0
    M__x: float = 4.90332500325983e-02

    d_s_low_x: float = 10.0
    s_low_x: float = 0.02
    d_s_up_x: float = 10.0
    s_up_x: float = 0.02
    d_s_low_y: float = 10.0
    s_low_y: float = 0.02
    d_s_up_y: float = 10.0
    s_up_y: float = 0.02

    a_low_x: float = 0.05
    a_up_x: float = 0.04

    k__max: float = 1000.0
    gr_d_: float = 0.1
    M: float = 3.92266000260786e-02
    h: float = 0.03
    b: float = 1.0
    a: float = 0.05
    a_vert: float = 0.04
    A__s: float = 0.00393
    A_vert__s: float = 0.00393

    # Управление условиями/проверками
    add_conds: bool = True
    checks: Tuple[str, ...] = ("2.6", "2.7", "2.8", "2.18", "2.20", "8.4 ÑÏ 52-103", "8.5 ÑÏ 52-103", "5.12")


CONDS_MOJIBAKE: Tuple[str, ...] = (
    "Àðìàòóðà ðàñïîëîæåíà ïî êîíòóðó ñå÷åíèÿ - íå ðàâíîìåðíî",
    "Ãðóïïà ïðåäåëüíûõ ñîñòîÿíèé - ïåðâàÿ",
    "Êîíñòðóêöèÿ - æåëåçîáåòîííàÿ",
    "Íàçíà÷åíèå êëàññà áåòîíà - ïî ïðî÷íîñòè íà ñæàòèå",
    "Îòíîñèòåëüíàÿ âëàæíîñòü âîçäóõà îêðóæàþùåé ñðåäû - 40 - 75%",
    "Ïîïåðåìåííîå çàìîðàæèâàíèå è îòòàèâàíèå ïðè òåìïåðàòóðå < 20°C - îòñóòñòâóåò",
    "Àðìàòóðà ïëèò - âåðõíÿÿ è íèæíÿÿ (èçãèá. ìîìåíòû ââîäÿòñÿ ñî ñâîèìè çíàêàìè)",
    "Ñå÷åíèå - ïðÿìîóãîëüíîå",
    "Ýëåìåíò - èçãèáàåìûé",
    "Ïðîãðåññèðóþùåå ðàçðóøåíèå - íå ðàññìàòðèâàåòñÿ â äàííîì ðàñ÷åòå",
    "Êîíñòðóêöèÿ áåòîíèðóåòñÿ - â ãîðèçîíòàëüíîì ïîëîæåíèè",
    "Êëàññ áåòîíà - B10",
    "Äåéñòâèå íàãðóçêè - íåïðîäîëæèòåëüíîå",
    "Ñåéñìè÷íîñòü ïëîùàäêè ñòðîèòåëüñòâà - íå áîëåå 6 áàëëîâ",
    "Êëàññ ïðîäîëüíîé àðìàòóðû - A240",
    "Ïîïåðå÷íàÿ àðìàòóðà - íå ðàññìàòðèâàåòñÿ â äàííîì ðàñ÷åòå",
)


def _set(vars_obj, name: str, value: float) -> None:
    """VB: Vars(VN(name)).Value = value"""
    vars_obj[VN(_s(name))].Value = value


def _make_steps_report_text(params: Params, per_check: List[Tuple[str, float]], nc_result: float) -> str:
    lines: List[str] = []
    lines.append("NormCAD calculation trace (ported from self_made_vb_modules/1.vb)")
    lines.append(f"Timestamp: {datetime.now().isoformat(timespec='seconds')}")
    lines.append(f"COM ProgID: {params.progid}")
    lines.append("")
    lines.append("Inputs (Vars):")
    lines.append(f"  gr_g__b1 = {params.gr_g__b1}")
    lines.append(f"  m__kp = {params.m__kp}")
    lines.append(f"  M__x = {params.M__x}")
    lines.append(f"  d__síx = {params.d_s_low_x}")
    lines.append(f"  s__íx = {params.s_low_x}")
    lines.append(f"  d__sâx = {params.d_s_up_x}")
    lines.append(f"  s__âx = {params.s_up_x}")
    lines.append(f"  d__síy = {params.d_s_low_y}")
    lines.append(f"  s__íy = {params.s_low_y}")
    lines.append(f"  d__sây = {params.d_s_up_y}")
    lines.append(f"  s__ây = {params.s_up_y}")
    lines.append(f"  a__íx = {params.a_low_x}")
    lines.append(f"  a__âx = {params.a_up_x}")
    lines.append(f"  k__max = {params.k__max}")
    lines.append(f"  gr_d_ = {params.gr_d_}")
    lines.append(f"  M = {params.M}")
    lines.append(f"  h = {params.h}")
    lines.append(f"  b = {params.b}")
    lines.append(f"  a = {params.a}")
    lines.append(f"  a_vert = {params.a_vert}")
    lines.append(f"  A__s = {params.A__s}")
    lines.append(f"  A_vert__s = {params.A_vert__s}")
    lines.append("")
    lines.append("Conds.Add used: " + ("YES" if params.add_conds else "NO"))
    if params.add_conds:
        lines.append("Conditions:")
        for c in CONDS_MOJIBAKE:
            lines.append("  - " + _s(c))
    lines.append("")
    lines.append("Checks (Vars.Ex):")
    for chk, r in per_check:
        lines.append(f"  S_{chk}: {r}")
    lines.append("")
    lines.append(f"NCResult (max): {nc_result}")
    lines.append("")
    return "\n".join(lines)


def calc_ncresult(params: Params) -> Tuple[float, List[Tuple[str, float]]]:
    """
    Полный аналог VB функции NCResult().
    Возвращает (nc_result, per_check_results)
    """
    vars_obj = _dispatch(params.progid)
    conds = vars_obj.Conds

    # Vars(...)
    _set(vars_obj, "gr_g__b1", params.gr_g__b1)
    _set(vars_obj, "m__kp", params.m__kp)
    _set(vars_obj, "M__x", params.M__x)

    _set(vars_obj, "d__síx", params.d_s_low_x)
    _set(vars_obj, "s__íx", params.s_low_x)
    _set(vars_obj, "d__sâx", params.d_s_up_x)
    _set(vars_obj, "s__âx", params.s_up_x)
    _set(vars_obj, "d__síy", params.d_s_low_y)
    _set(vars_obj, "s__íy", params.s_low_y)
    _set(vars_obj, "d__sây", params.d_s_up_y)
    _set(vars_obj, "s__ây", params.s_up_y)

    _set(vars_obj, "a__íx", params.a_low_x)
    _set(vars_obj, "a__âx", params.a_up_x)
    _set(vars_obj, "k__max", params.k__max)
    _set(vars_obj, "gr_d_", params.gr_d_)
    _set(vars_obj, "M", params.M)
    _set(vars_obj, "h", params.h)
    _set(vars_obj, "b", params.b)
    _set(vars_obj, "a", params.a)
    _set(vars_obj, "a_vert", params.a_vert)
    _set(vars_obj, "A__s", params.A__s)
    _set(vars_obj, "A_vert__s", params.A_vert__s)

    # Conds.Add(...)
    if params.add_conds:
        for c in CONDS_MOJIBAKE:
            conds.Add(_s(c))

    # checks
    try:
        vars_obj.Result = 0
    except Exception:
        pass

    nc_result = 0.0
    per_check: List[Tuple[str, float]] = []
    for chk in params.checks:
        vars_obj.Ex("S_" + VN(_s(chk)))
        try:
            r = float(vars_obj.Result)
        except Exception:
            r = 1e9
        per_check.append((chk, r))
        nc_result = _max(nc_result, r)
    return float(nc_result), per_check


def _try_make_normcad_report(
    *,
    vars_obj,
    conds_list: List[str],
    norm: str,
    task_name: str,
    unit: str,
    out_txt: Optional[str],
    out_doc: Optional[str],
) -> None:
    """
    Попытка сделать официальный отчёт через ncApi.Report (NCAPI.dll), как в NCBkP.pdf.
    Это доп. к нашему текстовому "trace" и может требовать корректных Norm/Unit.
    """
    import win32com.client

    nc = win32com.client.Dispatch("ncApi.Report")
    nc.Norm = norm
    nc.TaskName = task_name
    nc.Unit = unit
    nc.ClcLoadNorm()
    nc.SetVars(vars_obj)
    nc.ClcLoadData()
    if conds_list:
        nc.SetConds(conds_list)
        nc.ClcLoadConds()
    nc.ClcCalc()
    if out_txt:
        nc.MakeReport(out_txt)
    if out_doc:
        nc.SendToWord(out_doc)


def _parse_csv_strings(s: str) -> Tuple[str, ...]:
    parts = [p.strip() for p in (s or "").split(",")]
    out = tuple([p for p in parts if p])
    if not out:
        raise ValueError("empty list")
    return out


def main(argv: Optional[Sequence[str]] = None) -> int:
    d = Params()
    ap = argparse.ArgumentParser(description="Порт self_made_vb_modules/1.vb → Python (NormCAD COM)")

    ap.add_argument("--progid", default=d.progid)
    ap.add_argument("--no-conds", action="store_true", help="Не добавлять Conds.Add (оставить условия модуля по умолчанию).")
    ap.add_argument("--checks", default=",".join(d.checks), help="Список проверок без префикса S_ (через запятую).")

    # Отчёты
    ap.add_argument("--trace", help="Сохранить пошаговый текстовый лог (наш отчёт) в файл.")
    ap.add_argument("--nc-report", help="Сохранить официальный отчёт NormCAD (MakeReport) в файл .txt.")
    ap.add_argument("--word", help="Сохранить официальный отчёт NormCAD в Word (SendToWord) в файл .doc/.docx.")
    ap.add_argument("--norm", default="", help="NormCAD: ncApi.Report.Norm (например, 'СП 52-103-2007').")
    ap.add_argument("--task-name", default="self_made_vb_modules/1", help="NormCAD: ncApi.Report.TaskName.")
    ap.add_argument("--unit", default="", help="NormCAD: ncApi.Report.Unit (например, 'п.2.6, п.2.7').")

    # Полный набор переменных (дефолты как в VB)
    ap.add_argument("--gr_g__b1", type=float, default=d.gr_g__b1)
    ap.add_argument("--m__kp", type=float, default=d.m__kp)
    ap.add_argument("--M__x", type=float, default=d.M__x)

    ap.add_argument("--d_s_low_x", type=float, default=d.d_s_low_x)
    ap.add_argument("--s_low_x", type=float, default=d.s_low_x)
    ap.add_argument("--d_s_up_x", type=float, default=d.d_s_up_x)
    ap.add_argument("--s_up_x", type=float, default=d.s_up_x)
    ap.add_argument("--d_s_low_y", type=float, default=d.d_s_low_y)
    ap.add_argument("--s_low_y", type=float, default=d.s_low_y)
    ap.add_argument("--d_s_up_y", type=float, default=d.d_s_up_y)
    ap.add_argument("--s_up_y", type=float, default=d.s_up_y)

    ap.add_argument("--a_low_x", type=float, default=d.a_low_x)
    ap.add_argument("--a_up_x", type=float, default=d.a_up_x)
    ap.add_argument("--k__max", type=float, default=d.k__max)
    ap.add_argument("--gr_d_", type=float, default=d.gr_d_)
    ap.add_argument("--M", type=float, default=d.M)
    ap.add_argument("--h", type=float, default=d.h)
    ap.add_argument("--b", type=float, default=d.b)
    ap.add_argument("--a", type=float, default=d.a)
    ap.add_argument("--a_vert", type=float, default=d.a_vert)
    ap.add_argument("--A__s", type=float, default=d.A__s)
    ap.add_argument("--A_vert__s", type=float, default=d.A_vert__s)

    args = ap.parse_args(argv)

    if not _is_32bit_python():
        print("ERROR: Требуется 32-bit Python для COM-компонентов NormCAD.", file=sys.stderr)
        return 2

    params = Params(
        progid=args.progid,
        add_conds=(not args.no_conds),
        checks=_parse_csv_strings(args.checks),
        gr_g__b1=float(args.gr_g__b1),
        m__kp=float(args.m__kp),
        M__x=float(args.M__x),
        d_s_low_x=float(args.d_s_low_x),
        s_low_x=float(args.s_low_x),
        d_s_up_x=float(args.d_s_up_x),
        s_up_x=float(args.s_up_x),
        d_s_low_y=float(args.d_s_low_y),
        s_low_y=float(args.s_low_y),
        d_s_up_y=float(args.d_s_up_y),
        s_up_y=float(args.s_up_y),
        a_low_x=float(args.a_low_x),
        a_up_x=float(args.a_up_x),
        k__max=float(args.k__max),
        gr_d_=float(args.gr_d_),
        M=float(args.M),
        h=float(args.h),
        b=float(args.b),
        a=float(args.a),
        a_vert=float(args.a_vert),
        A__s=float(args.A__s),
        A_vert__s=float(args.A_vert__s),
    )

    try:
        res, per_check = calc_ncresult(params)
    except Exception as e:
        print(f"ERROR: {e.__class__.__name__}: {e}", file=sys.stderr)
        return 1

    # 1) Всегда печатаем численный итог (видно в терминале/PowerShell)
    print(res)

    # 2) Наш пошаговый отчёт (всегда доступен)
    if args.trace:
        trace_path = os.path.abspath(args.trace)
        os.makedirs(os.path.dirname(trace_path) or ".", exist_ok=True)
        txt = _make_steps_report_text(params, per_check, res)
        with open(trace_path, "w", encoding="utf-8") as f:
            f.write(txt)
        print(f"TRACE_WRITTEN: {trace_path}", file=sys.stderr)

    # 3) Официальный отчёт NormCAD (если попросили)
    if args.nc_report or args.word:
        # Для ncApi.Report нам нужен сам vars_obj, поэтому создадим ещё раз и передадим.
        # (win32com объект не сериализуется между вызовами, проще повторить установку)
        vars_obj = _dispatch(params.progid)
        conds_obj = vars_obj.Conds

        # Vars(...) повторно
        _set(vars_obj, "gr_g__b1", params.gr_g__b1)
        _set(vars_obj, "m__kp", params.m__kp)
        _set(vars_obj, "M__x", params.M__x)
        _set(vars_obj, "d__síx", params.d_s_low_x)
        _set(vars_obj, "s__íx", params.s_low_x)
        _set(vars_obj, "d__sâx", params.d_s_up_x)
        _set(vars_obj, "s__âx", params.s_up_x)
        _set(vars_obj, "d__síy", params.d_s_low_y)
        _set(vars_obj, "s__íy", params.s_low_y)
        _set(vars_obj, "d__sây", params.d_s_up_y)
        _set(vars_obj, "s__ây", params.s_up_y)
        _set(vars_obj, "a__íx", params.a_low_x)
        _set(vars_obj, "a__âx", params.a_up_x)
        _set(vars_obj, "k__max", params.k__max)
        _set(vars_obj, "gr_d_", params.gr_d_)
        _set(vars_obj, "M", params.M)
        _set(vars_obj, "h", params.h)
        _set(vars_obj, "b", params.b)
        _set(vars_obj, "a", params.a)
        _set(vars_obj, "a_vert", params.a_vert)
        _set(vars_obj, "A__s", params.A__s)
        _set(vars_obj, "A_vert__s", params.A_vert__s)

        conds_list: List[str] = []
        if params.add_conds:
            for c in CONDS_MOJIBAKE:
                cc = _s(c)
                conds_obj.Add(cc)
                conds_list.append(cc)

        out_txt = os.path.abspath(args.nc_report) if args.nc_report else None
        out_doc = os.path.abspath(args.word) if args.word else None
        if out_txt:
            os.makedirs(os.path.dirname(out_txt) or ".", exist_ok=True)
        if out_doc:
            os.makedirs(os.path.dirname(out_doc) or ".", exist_ok=True)

        try:
            _try_make_normcad_report(
                vars_obj=vars_obj,
                conds_list=conds_list,
                norm=str(args.norm or ""),
                task_name=str(args.task_name or ""),
                unit=str(args.unit or ""),
                out_txt=out_txt,
                out_doc=out_doc,
            )
            if out_txt:
                print(f"NC_REPORT_WRITTEN: {out_txt}", file=sys.stderr)
            if out_doc:
                print(f"WORD_WRITTEN: {out_doc}", file=sys.stderr)
        except Exception as e:
            print(f"NC_REPORT_ERROR: {e.__class__.__name__}: {e}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

