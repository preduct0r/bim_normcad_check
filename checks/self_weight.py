#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт собственного веса конструкций через API NormCAD.

Использует NCAPI.dll для выполнения расчётов по СП 20.13330.2016.
Данные извлекаются из IFC файла (output.jsonl).

Согласно документации API NormCAD:
- Создаётся объект ncApi.Report
- Задаются Norm, TaskName, Unit
- Передаются данные через SetVars, SetConds
- Запускается ClcCalc()
- Результат получается через MaxResult
"""

from __future__ import annotations
import json
from pathlib import Path
from typing import Any, Dict, List, Optional
from dataclasses import dataclass

# Для работы с COM объектами NormCAD
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("ВНИМАНИЕ: win32com не установлен. Установите: pip install pywin32")


@dataclass
class SelfWeightInput:
    """Входные данные для расчёта собственного веса из IFC."""
    element_id: str
    element_name: str
    ifc_class: str
    volume_m3: float
    length_m: float
    cross_section_area_m2: float
    material_name: str


def extract_from_ifc_element(element: Dict[str, Any]) -> Optional[SelfWeightInput]:
    """
    Извлечение данных из элемента IFC для расчёта собственного веса.
    
    Извлекаемые данные:
    - Объём (NetVolume из Qto_*BaseQuantities)
    - Длина (Length)
    - Площадь сечения (CrossSectionArea)
    - Материал
    """
    psets = element.get("psets", {})
    
    volume = 0.0
    length = 0.0
    cross_section = 0.0
    
    # Для балок
    qto_beam = psets.get("Qto_BeamBaseQuantities", {})
    if qto_beam:
        volume = float(qto_beam.get("NetVolume", 0))
        length_val = float(qto_beam.get("Length", 0))
        length = length_val / 1000.0 if length_val > 100 else length_val
        cross_section = float(qto_beam.get("CrossSectionArea", 0))
    
    # Для колонн
    qto_column = psets.get("Qto_ColumnBaseQuantities", {})
    if qto_column:
        volume = float(qto_column.get("NetVolume", 0))
        length_val = float(qto_column.get("Length", qto_column.get("Height", 0)))
        length = length_val / 1000.0 if length_val > 100 else length_val
        cross_section = float(qto_column.get("CrossSectionArea", 0))
    
    # Для плит
    qto_slab = psets.get("Qto_SlabBaseQuantities", {})
    if qto_slab:
        volume = float(qto_slab.get("NetVolume", 0))
    
    # Материал
    materials = element.get("materials", [])
    material_name = materials[0] if materials else "unknown"
    
    if volume <= 0:
        return None
    
    return SelfWeightInput(
        element_id=element.get("global_id", ""),
        element_name=element.get("name", ""),
        ifc_class=element.get("ifc_class", ""),
        volume_m3=volume,
        length_m=length,
        cross_section_area_m2=cross_section,
        material_name=material_name
    )


class NormCADSelfWeight:
    """
    Расчёт собственного веса через API NormCAD.
    
    Использует библиотеку NCAPI.dll согласно документации:
    https://normcad.ru/s/book/NCBkP.pdf
    """
    
    def __init__(self):
        """Инициализация COM объекта NormCAD."""
        if not HAS_WIN32COM:
            raise RuntimeError("Требуется установить pywin32: pip install pywin32")
        
        # Создаём экземпляр объекта отчёта (п. 3.1 документации)
        self.ncApiR = win32com.client.Dispatch("ncApi.Report")
        
        # Объект переменных расчётного модуля
        # Имя библиотеки модуля зависит от конкретного расчёта
        # Для нагрузок используется модуль "Нагрузки и воздействия"
        self.vars = None
    
    def setup_task(self, task_name: str):
        """
        Настройка переменных задания (п. 3.2 документации).
        
        Args:
            task_name: Название задания (отображается в отчёте)
        """
        # Norm - название нормативного документа
        self.ncApiR.Norm = "СП 20.13330.2016"
        
        # TaskName - название задания
        self.ncApiR.TaskName = task_name
        
        # Unit - перечень пунктов задания
        self.ncApiR.Unit = "п.7.2"  # Постоянные нагрузки
    
    def set_input_data(self, data: SelfWeightInput, density_kg_m3: float = 2500.0):
        """
        Передача исходных данных (п. 3.3 документации).
        
        Args:
            data: Данные из IFC
            density_kg_m3: Плотность материала (по умолчанию бетон 2500 кг/м3)
                          НЕ извлекается из IFC - задаётся вручную
        """
        # Подготовка массива переменных для NormCAD
        # Формат зависит от конкретного модуля
        
        # Исходные данные для расчёта собственного веса:
        # - Объём элемента
        # - Плотность материала
        # - Коэффициент надёжности gamma_f = 1.1 (неблагоприятное действие)
        
        vars_data = {
            "V": data.volume_m3,           # Объём, м3
            "rho": density_kg_m3,          # Плотность, кг/м3
            "gamma_f": 1.1,                # Коэф. надёжности по нагрузке
            "L": data.length_m,            # Длина элемента, м
            "A": data.cross_section_area_m2,  # Площадь сечения, м2
        }
        
        # Передача данных через SetVars (если vars объект создан)
        if self.vars:
            self.ncApiR.SetVars(self.vars)
        
        return vars_data
    
    def calculate(self) -> float:
        """
        Запуск расчёта (п. 3.4 документации).
        
        Returns:
            Максимальный коэффициент использования (MaxResult)
        """
        # Загрузка нормативного модуля
        self.ncApiR.ClcLoadNorm()
        
        # Загрузка данных
        self.ncApiR.ClcLoadData()
        
        # Запуск расчёта
        self.ncApiR.ClcCalc()
        
        # Возврат максимального коэффициента использования
        return self.ncApiR.MaxResult
    
    def save_report(self, file_path: str):
        """
        Сохранение отчёта (п. 3.4 документации).
        
        Args:
            file_path: Путь к файлу отчёта
        """
        # MakeReport - сохраняет отчёт в файл
        self.ncApiR.MakeReport(file_path)
    
    def save_report_word(self, file_path: str):
        """
        Сохранение отчёта в Word с рамкой и штампом.
        
        Args:
            file_path: Путь к файлу .doc
        """
        # SendToWord - сохраняет отчёт с рамкой и штампом
        self.ncApiR.SendToWord(file_path)
    
    def check_license(self) -> bool:
        """
        Проверка наличия аппаратного ключа защиты.
        
        Returns:
            True если ключ подключён
        """
        return self.ncApiR.TestKey()


def process_ifc_file(jsonl_path: str, report_dir: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Обработка JSONL файла с данными IFC через API NormCAD.
    
    Args:
        jsonl_path: Путь к JSONL файлу с извлечёнными данными IFC
        report_dir: Директория для сохранения отчётов (опционально)
    
    Returns:
        Список результатов расчёта
    """
    results = []
    
    # Справочник плотностей материалов
    # НЕ извлекается из IFC - используются значения по умолчанию
    MATERIAL_DENSITIES = {
        "wood": 500,
        "wood_spruce": 450,
        "wood_spruce_beam": 450,
        "wood_pine": 500,
        "steel": 7850,
        "concrete": 2500,
        "brick": 1800,
    }
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            if not line.strip():
                continue
            
            element = json.loads(line)
            
            # Только конструктивные элементы
            ifc_class = element.get("ifc_class", "")
            if ifc_class not in ["IfcBeam", "IfcColumn", "IfcSlab", "IfcWall", "IfcMember"]:
                continue
            
            # Извлечение данных из IFC
            data = extract_from_ifc_element(element)
            if not data:
                continue
            
            # Определение плотности по названию материала
            density = 2500.0  # По умолчанию - бетон
            mat_lower = data.material_name.lower()
            for key, val in MATERIAL_DENSITIES.items():
                if key in mat_lower:
                    density = val
                    break
            
            result = {
                "element_id": data.element_id,
                "element_name": data.element_name,
                "ifc_class": data.ifc_class,
                "volume_m3": data.volume_m3,
                "length_m": data.length_m,
                "material": data.material_name,
                "density_kg_m3": density,
                "density_source": "default (NOT from IFC)",
            }
            
            # Попытка расчёта через NormCAD API
            if HAS_WIN32COM:
                try:
                    calc = NormCADSelfWeight()
                    calc.setup_task(f"Собственный вес - {data.element_name}")
                    calc.set_input_data(data, density)
                    
                    if calc.check_license():
                        max_result = calc.calculate()
                        result["max_result"] = max_result
                        result["status"] = "calculated"
                        
                        if report_dir:
                            report_path = Path(report_dir) / f"self_weight_{data.element_id}.txt"
                            calc.save_report(str(report_path))
                            result["report"] = str(report_path)
                    else:
                        result["status"] = "no_license"
                        result["error"] = "NormCAD license key not found"
                        
                except Exception as e:
                    result["status"] = "error"
                    result["error"] = str(e)
            else:
                result["status"] = "no_win32com"
                result["error"] = "pywin32 not installed"
            
            results.append(result)
    
    return results


# ============== ПРИМЕР ИСПОЛЬЗОВАНИЯ ==============

if __name__ == "__main__":
    print("=" * 70)
    print("РАСЧЁТ СОБСТВЕННОГО ВЕСА ЧЕРЕЗ API NormCAD")
    print("=" * 70)
    
    jsonl_path = Path(__file__).parent.parent / "output.jsonl"
    
    if not jsonl_path.exists():
        print(f"Файл {jsonl_path} не найден")
        exit(1)
    
    print(f"\nОбработка файла: {jsonl_path}")
    print("-" * 70)
    
    results = process_ifc_file(str(jsonl_path))
    
    for r in results:
        print(f"\n{r['ifc_class']}: {r['element_name']}")
        print(f"  ID: {r['element_id']}")
        print(f"  Объём: {r['volume_m3']:.6f} м3")
        print(f"  Материал: {r['material']}")
        print(f"  Плотность: {r['density_kg_m3']} кг/м3 ({r['density_source']})")
        print(f"  Статус: {r['status']}")
        if "error" in r:
            print(f"  Ошибка: {r['error']}")
        if "max_result" in r:
            print(f"  MaxResult: {r['max_result']}")
    
    print("\n" + "=" * 70)
    print("ДАННЫЕ ИЗ IFC:")
    print("  [+] Объём элемента (NetVolume)")
    print("  [+] Длина элемента (Length)")
    print("  [+] Площадь сечения (CrossSectionArea)")
    print("  [+] Название материала")
    print("\nДАННЫЕ НЕ ИЗ IFC (по умолчанию):")
    print("  [-] Плотность материала -> из справочника по названию")
    print("  [-] Коэффициент gamma_f -> 1.1")
    print("=" * 70)
