#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Порт VB-скрипта `official_examples/armirovanie_pliti.vb` на Python.

Что делает скрипт (как оригинал):
- Создаёт COM-объект NormCAD Vars: "NC_167258518177598E02.Vars"
- Задаёт исходные переменные (геометрия/шаги/защитные слои и т.п.)
- Задаёт условия расчёта (Conds.Add ...)
- Для каждой строки усилий (Mx, My, Mxy, Qx, Qy) перебирает диаметры арматуры
  для (низ X, верх X, низ Y, верх Y) и выполняет проверки через Vars.Ex("S_"+VN(...)).
  Критерий: max(Vars.Result) <= 1.
- Пишет подобранную комбинацию и итоговый коэффициент.

Варианты ввода/вывода:
- Excel (рекомендуется для полного соответствия VBA): читает A:E, пишет F:G, как макрос.
- CSV: читает 5 колонок без заголовка/с заголовком, печатает и пишет *_out.csv.

Требования:
- Windows + установленный NormCAD/NormFEM с зарегистрированными COM-компонентами
- Python 32-bit + pywin32
"""

from __future__ import annotations

import argparse
import csv
import os
import struct
import sys
from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence, Tuple


def _is_32bit_python() -> bool:
    return struct.calcsize("P") * 8 == 32


def _parse_csv_floats(s: str) -> Tuple[float, ...]:
    parts = [p.strip() for p in (s or "").split(",")]
    out: List[float] = []
    for p in parts:
        if not p:
            continue
        out.append(float(p))
    if not out:
        raise ValueError("empty float list")
    return tuple(out)


def _parse_csv_strings(s: str) -> Tuple[str, ...]:
    parts = [p.strip() for p in (s or "").split(",")]
    out = tuple([p for p in parts if p])
    if not out:
        raise ValueError("empty string list")
    return out


def _fix_mojibake_cp1251(s: str) -> str:
    """
    В репозитории VB-файл часто отображается с "кракозябрами" (CP1251 → Latin-1).
    Чтобы гарантированно передать в COM корректные русские строки, держим в коде
    текст как он "виден" (латинские символы Àðì...), и конвертируем обратно.
    """
    try:
        return s.encode("latin-1", errors="strict").decode("cp1251", errors="strict")
    except Exception:
        # Если строка уже нормальная Unicode (или не латин-1), оставим как есть.
        return s


def VN(name: str) -> str:
    """
    Полная копия функции VN из VBA:
    - пробел -> _spc_
    - . -> _pnt_
    - - -> _minus_
    - ( -> _bkt1_
    - ) -> _bkt2_
    """
    name = name.replace(" ", "_spc_")
    name = name.replace(".", "_pnt_")
    name = name.replace("-", "_minus_")
    name = name.replace("(", "_bkt1_")
    name = name.replace(")", "_bkt2_")
    return name


def _s(x: str) -> str:
    """Shortcut: fix mojibake and return."""
    return _fix_mojibake_cp1251(x)


@dataclass(frozen=True)
class RowForces:
    mx: float
    my: float
    mxy: float
    qx: float
    qy: float


@dataclass(frozen=True)
class ArmParams:
    """
    Все параметры имеют значения по умолчанию (как в VB-скрипте),
    но их можно переопределять через CLI.
    """

    # COM ProgID нормативного модуля
    progid: str = "NC_167258518177598E02.Vars"

    # Переменные задания
    gr_g__b1: float = 1.0
    m__kp: float = 1.0

    # Геометрия/шаги/защитные слои (в VB имена с кириллицей отображаются "кракозябрами")
    s_low_x: float = 0.1
    s_up_x: float = 0.1
    s_low_y: float = 0.1
    s_up_y: float = 0.1
    a_low_x: float = 0.04
    a_up_x: float = 0.04
    a_low_y: float = 0.04
    a_up_y: float = 0.04
    h: float = 0.2
    b: float = 1.0

    # Списки проверок/диаметров
    diameters: Tuple[float, ...] = (10, 12, 14, 16, 20)
    pre_ex: Tuple[str, ...] = ("5.1.8", "5.1.9", "5.1.10", "5.2.7", "5.2.10")
    check_ex: Tuple[str, ...] = ("6.2.7", "8.4 ÑÏ 52-103", "8.5 ÑÏ 52-103", "8.3.4")

    # Условия расчёта
    add_conds: bool = True

# --- "кракозябры" из просмотра файла; конвертим в cp1251 через _s(...) ---
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
    "Êëàññ áåòîíà - B30",
    "Äåéñòâèå íàãðóçêè - íåïðîäîëæèòåëüíîå",
    "Ñåéñìè÷íîñòü ïëîùàäêè ñòðîèòåëüñòâà - íå áîëåå 6 áàëëîâ",
    "Êëàññ ïðîäîëüíîé àðìàòóðû - A400",
    "Ïîïåðå÷íàÿ àðìàòóðà - íå ðàññìàòðèâàåòñÿ â äàííîì ðàñ÷åòå",
)

def _dispatch(progid: str):
    import win32com.client

    return win32com.client.Dispatch(progid)


def _init_vars(params: ArmParams):
    """
    Создаёт Vars/Conds и заполняет всё как в VBA (строки/значения).
    Возвращает Vars COM-объект.
    """
    vars_obj = _dispatch(params.progid)
    conds = vars_obj.Conds

    # Переменные (как в VBA)
    vars_obj[VN("gr_g__b1")].Value = params.gr_g__b1
    vars_obj[VN("m__kp")].Value = params.m__kp

    # Эти имена в VB идут с кириллицей (низ/верх), в репе они "кракозябрами".
    vars_obj[VN(_s("s__íx"))].Value = params.s_low_x
    vars_obj[VN(_s("s__âx"))].Value = params.s_up_x
    vars_obj[VN(_s("s__íy"))].Value = params.s_low_y
    vars_obj[VN(_s("s__ây"))].Value = params.s_up_y
    vars_obj[VN(_s("a__íx"))].Value = params.a_low_x
    vars_obj[VN(_s("a__âx"))].Value = params.a_up_x
    vars_obj[VN(_s("a__íy"))].Value = params.a_low_y
    vars_obj[VN(_s("a__ây"))].Value = params.a_up_y
    vars_obj["h"].Value = params.h
    vars_obj["b"].Value = params.b

    # Условия (как в VBA)
    if params.add_conds:
        for c in CONDS_MOJIBAKE:
            conds.Add(_s(c))

    # Предварительные вычисления (как в VBA)
    for ex_name in params.pre_ex:
        vars_obj.Ex("S_" + VN(_s(ex_name)))

    return vars_obj


def _calc_for_row(vars_obj, forces: RowForces, params: ArmParams) -> Tuple[Optional[str], float]:
    """
    Возвращает (combo_text_or_None, max_result).
    combo_text_or_None == None, если комбинацию не удалось подобрать.
    """
    vars_obj["M__x"].Value = forces.mx
    vars_obj["M__y"].Value = forces.my
    vars_obj["M__xy"].Value = forces.mxy
    vars_obj["Q__x"].Value = forces.qx
    vars_obj["Q__y"].Value = forces.qy

    best_combo: Optional[str] = None
    best_result: float = 0.0

    for ix in params.diameters:
        for iy in params.diameters:
            for jx in params.diameters:
                for jy in params.diameters:
                    # В VBA: Vars("d__...").Value = ...
                    # Для полной совместимости прогоняем имя через VN(...)
                    vars_obj[VN(_s("d__síx"))].Value = ix
                    vars_obj[VN(_s("d__sâx"))].Value = jx
                    vars_obj[VN(_s("d__síy"))].Value = iy
                    vars_obj[VN(_s("d__sây"))].Value = jy

                    max_r = 0.0
                    # В VBA делается и NCResult=0 и Vars.Result=0
                    try:
                        vars_obj.Result = 0
                    except Exception:
                        pass

                    for ex_name in params.check_ex:
                        vars_obj.Ex("S_" + VN(_s(ex_name)))
                        try:
                            r = float(vars_obj.Result)
                        except Exception:
                            # Если Result не приводится, считаем провалом
                            r = 1e9
                        if r > max_r:
                            max_r = r

                    if max_r <= 1.0:
                        best_combo = f"{int(ix)} x {int(jx)} / {int(iy)} x {int(jy)}"
                        best_result = max_r
                        return best_combo, best_result

                    # если ни одна комбинация не подошла — запомним последний max_r
                    best_result = max(best_result, max_r)

    return best_combo, best_result


def _iter_csv_rows(path: str) -> List[RowForces]:
    with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        has_header = csv.Sniffer().has_header(sample)
        reader: Iterable[Sequence[str]]
        if has_header:
            dr = csv.DictReader(f, dialect=dialect)
            rows: List[RowForces] = []
            for r in dr:
                rows.append(
                    RowForces(
                        mx=float(r.get("Mx") or r.get("M__x") or r.get("mx") or 0),
                        my=float(r.get("My") or r.get("M__y") or r.get("my") or 0),
                        mxy=float(r.get("Mxy") or r.get("M__xy") or r.get("mxy") or 0),
                        qx=float(r.get("Qx") or r.get("Q__x") or r.get("qx") or 0),
                        qy=float(r.get("Qy") or r.get("Q__y") or r.get("qy") or 0),
                    )
                )
            return rows
        else:
            sr = csv.reader(f, dialect=dialect)
            rows = []
            for row in sr:
                if not row:
                    continue
                # поддержка строк с пробелами
                vals = [c.strip() for c in row if c is not None]
                if len(vals) < 5:
                    continue
                rows.append(RowForces(mx=float(vals[0]), my=float(vals[1]), mxy=float(vals[2]), qx=float(vals[3]), qy=float(vals[4])))
            return rows


def _run_excel(path: str, sheet_name: Optional[str]) -> None:
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(os.path.abspath(path))
    try:
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet

        # Как в VBA: очистить F:G, начать с A1
        ws.Range("F:G").ClearContents()

        # Параметры берём из глобального значения, проставленного в main()
        params = _PARAMS
        vars_obj = _init_vars(params)

        row = 0
        while True:
            row += 1
            cell_text = ws.Cells(row, 1).Value
            if cell_text is None or str(cell_text).strip() == "":
                break

            # Поддержка шаблона/файлов с заголовком: если 1-я строка нечисловая — пропускаем.
            try:
                forces = RowForces(
                    mx=float(ws.Cells(row, 1).Value or 0),
                    my=float(ws.Cells(row, 2).Value or 0),
                    mxy=float(ws.Cells(row, 3).Value or 0),
                    qx=float(ws.Cells(row, 4).Value or 0),
                    qy=float(ws.Cells(row, 5).Value or 0),
                )
            except Exception:
                if row == 1:
                    continue
                raise
            combo, nc_result = _calc_for_row(vars_obj, forces, params)
            if combo is not None:
                ws.Cells(row, 6).Value = combo
            ws.Cells(row, 7).Value = float(nc_result)

        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


def _run_csv(path: str) -> None:
    params = _PARAMS
    vars_obj = _init_vars(params)
    rows = _iter_csv_rows(path)

    out_path = os.path.splitext(path)[0] + "_out.csv"
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Mx", "My", "Mxy", "Qx", "Qy", "Arm", "NCResult"])
        for forces in rows:
            combo, nc_result = _calc_for_row(vars_obj, forces, params)
            w.writerow([forces.mx, forces.my, forces.mxy, forces.qx, forces.qy, combo or "", nc_result])

    print(f"Wrote: {out_path}")


def _create_excel_template(path: str) -> str:
    """
    Создаёт новый Excel-файл-шаблон для ввода усилий (A:E) и вывода (F:G).
    Файл создаётся только если пользователь не передал --excel/--csv.
    """
    import win32com.client

    abspath = os.path.abspath(path)
    os.makedirs(os.path.dirname(abspath) or ".", exist_ok=True)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Add()
    try:
        ws = wb.ActiveSheet
        ws.Name = "Forces"

        # "Чистый" шаблон: только входные колонки A:E.
        # Колонки F:G будут заполняться скриптом при расчёте.
        headers = ["M__x", "M__y", "M__xy", "Q__x", "Q__y"]
        for i, h in enumerate(headers, start=1):
            ws.Cells(1, i).Value = h

        # Немного косметики (не влияет на функционал)
        ws.Columns("A:E").AutoFit()

        # Сохраняем
        # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.SaveAs(abspath, FileFormat=51)
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()
    return abspath


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Порт armirovanie_pliti.vb на Python (NormCAD COM)")
    io = ap.add_mutually_exclusive_group(required=False)
    io.add_argument("--excel", help="Путь к .xlsx/.xlsm. Читает A:E, пишет F:G (как VBA).")
    io.add_argument("--csv", help="Путь к CSV (колонки Mx,My,Mxy,Qx,Qy или 5 колонок без заголовка).")
    ap.add_argument("--sheet", help="Имя листа (если не указано — ActiveSheet).")
    ap.add_argument(
        "--new-excel",
        default="armirovanie_pliti_input.xlsx",
        help="Путь для создания нового Excel-шаблона, если не передан --excel/--csv.",
    )

    # Все параметры имеют значения по умолчанию (как в VB-скрипте)
    d = ArmParams()  # только для дефолтов argparse
    ap.add_argument("--progid", default=d.progid, help="ProgID COM объекта Vars (NormCAD модуль).")
    ap.add_argument("--gr_g__b1", type=float, default=d.gr_g__b1)
    ap.add_argument("--m__kp", type=float, default=d.m__kp)

    ap.add_argument("--s_low_x", type=float, default=d.s_low_x)
    ap.add_argument("--s_up_x", type=float, default=d.s_up_x)
    ap.add_argument("--s_low_y", type=float, default=d.s_low_y)
    ap.add_argument("--s_up_y", type=float, default=d.s_up_y)
    ap.add_argument("--a_low_x", type=float, default=d.a_low_x)
    ap.add_argument("--a_up_x", type=float, default=d.a_up_x)
    ap.add_argument("--a_low_y", type=float, default=d.a_low_y)
    ap.add_argument("--a_up_y", type=float, default=d.a_up_y)
    ap.add_argument("--h", type=float, default=d.h)
    ap.add_argument("--b", type=float, default=d.b)

    ap.add_argument(
        "--diameters",
        default=",".join(str(int(x)) for x in d.diameters),
        help="Список диаметров через запятую, например: 10,12,14,16,20",
    )
    ap.add_argument(
        "--pre-ex",
        default=",".join(d.pre_ex),
        help="Список предварительных S_ проверок через запятую (как в VB до цикла).",
    )
    ap.add_argument(
        "--check-ex",
        default=",".join(d.check_ex),
        help="Список S_ проверок через запятую (как в VB внутри цикла).",
    )
    ap.add_argument(
        "--no-conds",
        action="store_true",
        help="Не добавлять условия Conds.Add (использовать условия модуля по умолчанию).",
    )
    args = ap.parse_args(argv)

    if not _is_32bit_python():
        print("ERROR: Требуется 32-bit Python для COM-компонентов NormCAD/NormFEM.", file=sys.stderr)
        return 2

    global _PARAMS
    _PARAMS = ArmParams(
        progid=args.progid,
        gr_g__b1=float(args.gr_g__b1),
        m__kp=float(args.m__kp),
        s_low_x=float(args.s_low_x),
        s_up_x=float(args.s_up_x),
        s_low_y=float(args.s_low_y),
        s_up_y=float(args.s_up_y),
        a_low_x=float(args.a_low_x),
        a_up_x=float(args.a_up_x),
        a_low_y=float(args.a_low_y),
        a_up_y=float(args.a_up_y),
        h=float(args.h),
        b=float(args.b),
        diameters=_parse_csv_floats(args.diameters),
        pre_ex=_parse_csv_strings(args.pre_ex),
        check_ex=_parse_csv_strings(args.check_ex),
        add_conds=(not args.no_conds),
    )

    if not args.excel and not args.csv:
        try:
            out = _create_excel_template(args.new_excel)
            print(f"Created template: {out}")
            print("Fill columns A:E (M__x, M__y, M__xy, Q__x, Q__y), then re-run with --excel <file>.")
            return 0
        except Exception as e:
            print(f"ERROR: {e.__class__.__name__}: {e}", file=sys.stderr)
            return 1

    try:
        if args.excel:
            _run_excel(args.excel, args.sheet)
        if args.csv:
            _run_csv(args.csv)
    except Exception as e:
        print(f"ERROR: {e.__class__.__name__}: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


# Параметры, собранные в main(). Нужны, чтобы не прокидывать params через все уровни CLI → IO.
_PARAMS: ArmParams = ArmParams()

