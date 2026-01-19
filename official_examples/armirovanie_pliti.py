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


DIAMETERS: Tuple[float, ...] = (10, 12, 14, 16, 20)

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

PRE_EX_MOJIBAKE: Tuple[str, ...] = ("5.1.8", "5.1.9", "5.1.10", "5.2.7", "5.2.10")
CHECK_EX_MOJIBAKE: Tuple[str, ...] = ("6.2.7", "8.4 ÑÏ 52-103", "8.5 ÑÏ 52-103", "8.3.4")


def _dispatch(progid: str):
    import win32com.client

    return win32com.client.Dispatch(progid)


def _init_vars():
    """
    Создаёт Vars/Conds и заполняет всё как в VBA (строки/значения).
    Возвращает Vars COM-объект.
    """
    vars_obj = _dispatch("NC_167258518177598E02.Vars")
    conds = vars_obj.Conds

    # Переменные (как в VBA)
    vars_obj[VN("gr_g__b1")].Value = 1
    vars_obj[VN("m__kp")].Value = 1

    # Эти имена в VB идут с кириллицей (низ/верх), в репе они "кракозябрами".
    vars_obj[_s("s__íx")].Value = 0.1
    vars_obj[_s("s__âx")].Value = 0.1
    vars_obj[_s("s__íy")].Value = 0.1
    vars_obj[_s("s__ây")].Value = 0.1
    vars_obj[_s("a__íx")].Value = 0.04
    vars_obj[_s("a__âx")].Value = 0.04
    vars_obj[_s("a__íy")].Value = 0.04
    vars_obj[_s("a__ây")].Value = 0.04
    vars_obj["h"].Value = 0.2
    vars_obj["b"].Value = 1

    # Условия (как в VBA)
    for c in CONDS_MOJIBAKE:
        conds.Add(_s(c))

    # Предварительные вычисления (как в VBA)
    for ex_name in PRE_EX_MOJIBAKE:
        vars_obj.Ex("S_" + VN(_s(ex_name)))

    return vars_obj


def _calc_for_row(vars_obj, forces: RowForces) -> Tuple[Optional[str], float]:
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

    for ix in DIAMETERS:
        for iy in DIAMETERS:
            for jx in DIAMETERS:
                for jy in DIAMETERS:
                    vars_obj[_s("d__síx")].Value = ix
                    vars_obj[_s("d__sâx")].Value = jx
                    vars_obj[_s("d__síy")].Value = iy
                    vars_obj[_s("d__sây")].Value = jy

                    max_r = 0.0
                    # В VBA делается и NCResult=0 и Vars.Result=0
                    try:
                        vars_obj.Result = 0
                    except Exception:
                        pass

                    for ex_name in CHECK_EX_MOJIBAKE:
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

        vars_obj = _init_vars()

        row = 0
        while True:
            row += 1
            cell_text = ws.Cells(row, 1).Value
            if cell_text is None or str(cell_text).strip() == "":
                break

            forces = RowForces(
                mx=float(ws.Cells(row, 1).Value or 0),
                my=float(ws.Cells(row, 2).Value or 0),
                mxy=float(ws.Cells(row, 3).Value or 0),
                qx=float(ws.Cells(row, 4).Value or 0),
                qy=float(ws.Cells(row, 5).Value or 0),
            )
            combo, nc_result = _calc_for_row(vars_obj, forces)
            if combo is not None:
                ws.Cells(row, 6).Value = combo
            ws.Cells(row, 7).Value = float(nc_result)

        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


def _run_csv(path: str) -> None:
    vars_obj = _init_vars()
    rows = _iter_csv_rows(path)

    out_path = os.path.splitext(path)[0] + "_out.csv"
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Mx", "My", "Mxy", "Qx", "Qy", "Arm", "NCResult"])
        for forces in rows:
            combo, nc_result = _calc_for_row(vars_obj, forces)
            w.writerow([forces.mx, forces.my, forces.mxy, forces.qx, forces.qy, combo or "", nc_result])

    print(f"Wrote: {out_path}")


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Порт armirovanie_pliti.vb на Python (NormCAD COM)")
    ap.add_argument("--excel", help="Путь к .xlsx/.xlsm. Читает A:E, пишет F:G (как VBA).")
    ap.add_argument("--sheet", help="Имя листа (если не указано — ActiveSheet).")
    ap.add_argument("--csv", help="Путь к CSV (колонки Mx,My,Mxy,Qx,Qy или 5 колонок без заголовка).")
    args = ap.parse_args(argv)

    if not _is_32bit_python():
        print("ERROR: Требуется 32-bit Python для COM-компонентов NormCAD/NormFEM.", file=sys.stderr)
        return 2

    if not args.excel and not args.csv:
        ap.print_help()
        return 2

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

