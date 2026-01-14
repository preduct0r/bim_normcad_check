#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт усилий для ферм через API NormFEM.

Перевод с VB кода (normfem.vb) на Python.
Использует библиотеку ncfemapi.dll через COM.

Согласно документации API NormFEM:
- nfApi.SetPath - настройка путей
- nfApi.Calc - запуск расчёта
- nfApi.GetArrZ - получение перемещений
- nfApi.GetArrNM - получение нормальных сил и моментов
- nfApi.GetArrQ - получение поперечных сил
"""

from typing import Optional, Tuple, List, Any

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("ВНИМАНИЕ: win32com не установлен. Установите: pip install pywin32")


class NormFEMCalculator:
    """
    Расчёт усилий через API NormFEM.
    
    Использует библиотеку ncfemapi.dll согласно документации:
    https://normcad.ru/s/API_NormFEM.htm
    """
    
    def __init__(self, 
                 app_path: str = "C:\\Program Files (x86)\\NormCAD",
                 temp_dir: str = "C:\\Temp",
                 project_name: str = "Project"):
        """
        Инициализация API NormFEM.
        
        Args:
            app_path: Путь к папке NormCAD
            temp_dir: Путь к временной директории
            project_name: Имя проекта
        """
        self.app_path = app_path
        self.temp_dir = temp_dir
        self.project_name = project_name
        
        self.nfApi = None
        self.report_lines: List[str] = []
        
        # Результаты расчёта
        self.ArrZ: Optional[Any] = None   # перемещения узлов
        self.ArrNM: Optional[Any] = None  # нормальные силы и моменты
        self.ArrQ: Optional[Any] = None   # поперечные силы
        
        if HAS_WIN32COM:
            try:
                # Создаём экземпляр API NormFEM (п. 4.1 документации)
                self.nfApi = win32com.client.Dispatch("ncfemapi.Main")
                
                # Настройка путей (п. 4.1 документации)
                # SetPath(AppPath, TempDir, Project, ParentForm)
                self.nfApi.SetPath(app_path, temp_dir, project_name, None)
                
            except Exception as e:
                self.nfApi = None
                print(f"Ошибка инициализации NormFEM API: {e}")
    
    def report_add(self, message: str):
        """Добавление строки в отчёт."""
        self.report_lines.append(message)
        print(f"[REPORT] {message}")
    
    def add_all_arr(self, Gpkr: float, Gstn: float):
        """
        Заполнение массивов данных для расчёта.
        
        Здесь должна быть передача таблиц через nfApi.SetArr():
        - m00 - материалы
        - g01 - узлы
        - g02 - элементы
        - g03 - связи
        - и т.д.
        
        Args:
            Gpkr: расчётный вес покрытия
            Gstn: расчётный вес стен
        """
        if not self.nfApi:
            return
        
        # Пример передачи массива материалов (п. 4.2 документации)
        # mArr = ["Steel; 2.1e5; 0.3"]
        # self.nfApi.SetArr("m00", mArr)
        
        # Пример передачи узлов, элементов и т.д.
        # self.nfApi.SetArr("g01", nodes_arr)
        # self.nfApi.SetArr("g02", elements_arr)
        
        self.report_add(f"Загружены данные: Gpkr={Gpkr}, Gstn={Gstn}")
    
    def get_comb(self):
        """Расчёт усилий от сочетаний нагрузок."""
        self.report_add("Расчёт сочетаний нагрузок")
        # Здесь логика расчёта комбинаций
    
    def max_min_n(self):
        """Определение максимальных и минимальных усилий в элементах ферм."""
        self.report_add("Определение max/min усилий")
        # Здесь логика поиска экстремумов
    
    def calc(self, Gp: float, Gs: float) -> bool:
        """
        Функция расчёта усилий для ферм.
        
        Перевод VB функции Calc() на Python.
        
        Args:
            Gp: расчётный вес покрытия
            Gs: расчётный вес стен
        
        Returns:
            True если расчёт успешен, False иначе
        """
        self.report_add("Расчёт усилий")
        
        Gpkr = Gp  # расчётный вес покрытия
        Gstn = Gs  # расчётный вес стен
        
        # Заполнение массивов данных для расчёта
        self.add_all_arr(Gpkr, Gstn)
        
        if not self.nfApi:
            self.report_add("NormFEM API недоступен")
            return False
        
        try:
            # Запуск расчёта через API NormFEM (п. 4.3 документации)
            self.nfApi.Calc()
            
            # Проверка успешности расчёта
            result = self.nfApi.Result
            
            if not result:
                self.report_add("Расчёт не выполнен")
                return False
            
            # Получение массивов результатов из API NormFEM (п. 4.4 документации)
            self.ArrZ = self.nfApi.GetArrZ()    # перемещения узлов
            self.ArrNM = self.nfApi.GetArrNM()  # нормальные силы и моменты
            self.ArrQ = self.nfApi.GetArrQ()    # поперечные силы
            
            self.report_add(f"Получены результаты: ArrZ, ArrNM, ArrQ")
            
            # Расчёт усилий от сочетаний нагрузок
            self.get_comb()
            
            # Определение максимальных и минимальных усилий в элементах ферм
            self.max_min_n()
            
            return True
            
        except Exception as e:
            self.report_add(f"Ошибка расчёта: {e}")
            return False
        
        finally:
            # Освобождение объекта API
            self.nfApi = None
    
    def get_report(self) -> str:
        """Получение полного отчёта."""
        return "\n".join(self.report_lines)


# ============== ТЕСТОВЫЙ ЗАПУСК ==============

if __name__ == "__main__":
    print("=" * 70)
    print("ТЕСТ API NormFEM (перевод с VB)")
    print("=" * 70)
    
    # Фейковые параметры для теста
    Gp_test = 1.5  # расчётный вес покрытия, кН/м2
    Gs_test = 2.0  # расчётный вес стен, кН/м2
    
    print(f"\nВходные параметры:")
    print(f"  Gp (вес покрытия): {Gp_test} кН/м2")
    print(f"  Gs (вес стен): {Gs_test} кН/м2")
    print("-" * 70)
    
    # Создание калькулятора
    calculator = NormFEMCalculator(
        app_path="C:\\Program Files (x86)\\NormCAD",
        temp_dir="C:\\Temp",
        project_name="TestProject"
    )
    
    # Запуск расчёта
    print("\nЗапуск расчёта...")
    success = calculator.calc(Gp_test, Gs_test)
    
    print("-" * 70)
    print(f"Результат: {'УСПЕХ' if success else 'ОШИБКА'}")
    
    if calculator.ArrZ is not None:
        print(f"  ArrZ (перемещения): получены")
    if calculator.ArrNM is not None:
        print(f"  ArrNM (силы/моменты): получены")
    if calculator.ArrQ is not None:
        print(f"  ArrQ (поперечные силы): получены")
    
    print("\n" + "=" * 70)
    print("ОТЧЁТ:")
    print("=" * 70)
    print(calculator.get_report())
    
    print("\n" + "=" * 70)
    print("ПРИМЕЧАНИЕ:")
    print("  Для корректной работы требуется:")
    print("  1. Установленный NormCAD/NormFEM")
    print("  2. Python 32-bit (для 32-bit NormCAD)")
    print("  3. Зарегистрированная ncfemapi.dll")
    print("=" * 70)
