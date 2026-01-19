#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Создаёт Excel-файлы для `official_examples/armirovanie_pliti.py`:
- чистый шаблон только A:E
- тестовый файл с несколькими строками усилий (A:E)

Формат как у VBA/скрипта:
- вход: A:E = M__x, M__y, M__xy, Q__x, Q__y
- выход: F:G = Arm, NCResult

Запуск (важно: 32-bit Python):
  C:\\Users\\servuser\\Desktop\\test_normcad\\env_32\\Scripts\\python.exe official_examples\\create_armirovanie_pliti_test_excel.py
"""

from __future__ import annotations

import os


def main() -> int:
    import win32com.client

    template_path = os.path.abspath(r"official_examples\armirovanie_pliti_template_A-E.xlsx")
    test_path = os.path.abspath(r"official_examples\armirovanie_pliti_test.xlsx")
    os.makedirs(os.path.dirname(template_path), exist_ok=True)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # 1) Чистый шаблон только A:E
        wb = excel.Workbooks.Add()
        ws = wb.ActiveSheet
        ws.Name = "Forces"
        headers = ["M__x", "M__y", "M__xy", "Q__x", "Q__y"]
        for i, h in enumerate(headers, start=1):
            ws.Cells(1, i).Value = h
        ws.Columns("A:E").AutoFit()
        wb.SaveAs(template_path, FileFormat=51)  # .xlsx
        wb.Close(SaveChanges=True)

        # 2) Тестовый файл (тоже только A:E; F:G заполнит скрипт расчёта)
        wb = excel.Workbooks.Add()
        ws = wb.ActiveSheet
        ws.Name = "Forces"
        for i, h in enumerate(headers, start=1):
            ws.Cells(1, i).Value = h

        # Тестовые строки усилий (A:E)
        ws.Cells(2, 1).Value = 0
        ws.Cells(2, 2).Value = 0
        ws.Cells(2, 3).Value = 0
        ws.Cells(2, 4).Value = 0
        ws.Cells(2, 5).Value = 0

        ws.Cells(3, 1).Value = 5
        ws.Cells(3, 2).Value = 3
        ws.Cells(3, 3).Value = 0.5
        ws.Cells(3, 4).Value = 2
        ws.Cells(3, 5).Value = 1

        ws.Cells(4, 1).Value = 10
        ws.Cells(4, 2).Value = 8
        ws.Cells(4, 3).Value = 1
        ws.Cells(4, 4).Value = 4
        ws.Cells(4, 5).Value = 3

        ws.Columns("A:E").AutoFit()
        wb.SaveAs(test_path, FileFormat=51)  # .xlsx
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()

    print(template_path)
    print(test_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

