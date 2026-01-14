#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт снеговой нагрузки через API NormCAD.

Использует NCAPI.dll для выполнения расчётов по СП 20.13330.2016.
Модуль "Нагрузки и воздействия" -> Снеговая нагрузка.

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
class SnowLoadInput:
    """
    Входные данные для расчёта снеговой нагрузки.
    
    Данные из IFC:
    - element_id, element_name: идентификация элемента
    - area_m2: площадь покрытия (из Qto)
    
    Данные НЕ из IFC (по умолчанию):
    - snow_region: снеговой район (I-VIII)
    - roof_angle_deg: угол наклона кровли
    - mu: коэффициент перехода (неравномерность)
    - Ce: коэффициент сноса снега
    - Ct: термический коэффициент
    """
    # Из IFC
    element_id: str = ""
    element_name: str = ""
    area_m2: float = 0.0
    
    # НЕ из IFC - значения по умолчанию
    snow_region: str = "III"
    roof_angle_deg: float = 0.0
    mu: float = 1.0
    Ce: float = 1.0
    Ct: float = 1.0


# Таблица 10.1 СП 20 - Вес снегового покрова Sg (кПа) по районам
SNOW_REGIONS_SG = {
    "I": 0.5,
    "II": 1.0,
    "III": 1.5,
    "IV": 2.0,
    "V": 2.5,
    "VI": 3.0,
    "VII": 3.5,
    "VIII": 4.0,
}


def extract_from_ifc_element(element: Dict[str, Any]) -> Optional[SnowLoadInput]:
    """
    Извлечение данных из элемента IFC для расчёта снеговой нагрузки.
    
    Ищем элементы кровли (IfcSlab с типом ROOF или IfcRoof).
    """
    ifc_class = element.get("ifc_class", "")
    predefined_type = element.get("predefined_type", "")
    name = element.get("name", "")
    
    # Фильтруем только элементы кровли
    is_roof = (
        ifc_class == "IfcRoof" or
        predefined_type == "ROOF" or
        "roof" in name.lower() or
        "кровля" in name.lower()
    )
    
    if not is_roof and ifc_class != "IfcSlab":
        return None
    
    psets = element.get("psets", {})
    area = 0.0
    
    # Площадь из Qto
    qto_slab = psets.get("Qto_SlabBaseQuantities", {})
    if "GrossArea" in qto_slab:
        area = float(qto_slab["GrossArea"])
    elif "NetArea" in qto_slab:
        area = float(qto_slab["NetArea"])
    
    # Pset для кровли
    roof_pset = psets.get("Pset_RoofCommon", {})
    if "ProjectedArea" in roof_pset:
        area = float(roof_pset["ProjectedArea"])
    
    if area <= 0:
        return None
    
    return SnowLoadInput(
        element_id=element.get("global_id", ""),
        element_name=element.get("name", ""),
        area_m2=area
    )


class NormCADSnowLoad:
    """
    Расчёт снеговой нагрузки через API NormCAD.
    
    Использует библиотеку NCAPI.dll согласно документации:
    https://normcad.ru/s/book/NCBkP.pdf
    
    Модуль: "Нагрузки и воздействия" -> Снеговая нагрузка
    """
    
    def __init__(self):
        """Инициализация COM объекта NormCAD."""
        if not HAS_WIN32COM:
            raise RuntimeError("Требуется установить pywin32: pip install pywin32")
        
        # Создаём экземпляр объекта отчёта (п. 3.1 документации)
        self.ncApiR = win32com.client.Dispatch("ncApi.Report")
        
        # Объект переменных расчётного модуля
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
        
        # Unit - перечень пунктов задания (раздел 10 - снеговые нагрузки)
        self.ncApiR.Unit = "п.10"
    
    def set_input_data(self, data: SnowLoadInput):
        """
        Передача исходных данных (п. 3.3 документации).
        
        Args:
            data: Данные для расчёта (частично из IFC, частично по умолчанию)
        """
        # Получаем Sg по району
        Sg = SNOW_REGIONS_SG.get(data.snow_region, 1.5)
        
        # Подготовка массива переменных для NormCAD
        # Формат зависит от конкретного модуля "Снеговая нагрузка"
        vars_data = {
            # Исходные данные по СП 20 п.10
            "Sg": Sg,                      # Вес снегового покрова, кПа
            "mu": data.mu,                 # Коэффициент перехода (неравномерность)
            "Ce": data.Ce,                 # Коэффициент сноса
            "Ct": data.Ct,                 # Термический коэффициент
            "gamma_f": 1.4,                # Коэф. надёжности по снеговой нагрузке
            "alpha": data.roof_angle_deg,  # Угол наклона кровли
            "A": data.area_m2,             # Площадь покрытия
            
            # Параметры для учёта неравномерности (прил. Б СП 20)
            "snow_region": data.snow_region,
        }
        
        # Передача данных через SetVars
        if self.vars:
            self.ncApiR.SetVars(self.vars)
        
        return vars_data
    
    def set_conditions(self, conditions: Dict[str, Any]):
        """
        Передача условий расчёта (п. 3.3 документации).
        
        Args:
            conditions: Словарь условий (тип кровли, наличие перепадов и т.д.)
        """
        # ArConds - массив условий
        # Например: тип кровли, наличие снеговых мешков и т.д.
        if conditions:
            self.ncApiR.SetConds(conditions)
            self.ncApiR.ClcLoadConds()
    
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
        self.ncApiR.MakeReport(file_path)
    
    def save_report_word(self, file_path: str):
        """
        Сохранение отчёта в Word с рамкой и штампом.
        
        Args:
            file_path: Путь к файлу .doc
        """
        self.ncApiR.SendToWord(file_path)
    
    def check_license(self) -> bool:
        """
        Проверка наличия аппаратного ключа защиты.
        
        Returns:
            True если ключ подключён
        """
        return self.ncApiR.TestKey()


def process_ifc_file(jsonl_path: str, 
                     snow_region: str = "III",
                     report_dir: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Обработка JSONL файла с данными IFC через API NormCAD.
    
    Args:
        jsonl_path: Путь к JSONL файлу с извлечёнными данными IFC
        snow_region: Снеговой район (НЕ из IFC, задаётся вручную)
        report_dir: Директория для сохранения отчётов (опционально)
    
    Returns:
        Список результатов расчёта
    """
    results = []
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            if not line.strip():
                continue
            
            element = json.loads(line)
            
            # Извлечение данных из IFC
            data = extract_from_ifc_element(element)
            if not data:
                continue
            
            # Применяем параметры по умолчанию (НЕ из IFC)
            data.snow_region = snow_region
            
            Sg = SNOW_REGIONS_SG.get(snow_region, 1.5)
            
            result = {
                "element_id": data.element_id,
                "element_name": data.element_name,
                "area_m2": data.area_m2,
                "area_source": "from IFC (Qto)",
                "snow_region": snow_region,
                "snow_region_source": "default (NOT from IFC)",
                "Sg_kPa": Sg,
                "mu": data.mu,
                "mu_source": "default (NOT from IFC)",
                "Ce": data.Ce,
                "Ct": data.Ct,
            }
            
            # Попытка расчёта через NormCAD API
            if HAS_WIN32COM:
                try:
                    calc = NormCADSnowLoad()
                    calc.setup_task(f"Снеговая нагрузка - {data.element_name}")
                    calc.set_input_data(data)
                    
                    if calc.check_license():
                        max_result = calc.calculate()
                        result["max_result"] = max_result
                        result["status"] = "calculated"
                        
                        if report_dir:
                            report_path = Path(report_dir) / f"snow_load_{data.element_id}.txt"
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


def calculate_single(
    area_m2: float,
    snow_region: str = "III",
    mu: float = 1.0,
    Ce: float = 1.0,
    Ct: float = 1.0,
    roof_angle_deg: float = 0.0,
    task_name: str = "Снеговая нагрузка",
    report_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    Расчёт снеговой нагрузки для одного элемента через API NormCAD.
    
    Args:
        area_m2: Площадь покрытия (может быть из IFC)
        snow_region: Снеговой район I-VIII (НЕ из IFC)
        mu: Коэффициент перехода (НЕ из IFC)
        Ce: Коэффициент сноса (НЕ из IFC)
        Ct: Термический коэффициент (НЕ из IFC)
        roof_angle_deg: Угол наклона кровли (НЕ из IFC)
        task_name: Название задания для отчёта
        report_path: Путь для сохранения отчёта
    
    Returns:
        Словарь с результатами
    """
    data = SnowLoadInput(
        area_m2=area_m2,
        snow_region=snow_region,
        mu=mu,
        Ce=Ce,
        Ct=Ct,
        roof_angle_deg=roof_angle_deg
    )
    
    Sg = SNOW_REGIONS_SG.get(snow_region, 1.5)
    
    result = {
        "area_m2": area_m2,
        "snow_region": snow_region,
        "Sg_kPa": Sg,
        "mu": mu,
        "Ce": Ce,
        "Ct": Ct,
        "roof_angle_deg": roof_angle_deg,
    }
    
    if HAS_WIN32COM:
        try:
            calc = NormCADSnowLoad()
            calc.setup_task(task_name)
            calc.set_input_data(data)
            
            if calc.check_license():
                max_result = calc.calculate()
                result["max_result"] = max_result
                result["status"] = "calculated"
                
                if report_path:
                    calc.save_report(report_path)
                    result["report"] = report_path
            else:
                result["status"] = "no_license"
                result["error"] = "NormCAD license key not found"
                
        except Exception as e:
            result["status"] = "error"
            result["error"] = str(e)
    else:
        result["status"] = "no_win32com"
        result["error"] = "pywin32 not installed"
    
    return result


# ============== ПРИМЕР ИСПОЛЬЗОВАНИЯ ==============

if __name__ == "__main__":
    print("=" * 70)
    print("РАСЧЁТ СНЕГОВОЙ НАГРУЗКИ ЧЕРЕЗ API NormCAD")
    print("=" * 70)
    
    # Пример 1: Расчёт для одного элемента
    print("\n--- Пример: расчёт для площади 100 м2 ---")
    
    result = calculate_single(
        area_m2=100.0,           # Площадь (может быть из IFC)
        snow_region="III",       # НЕ из IFC
        mu=1.0,                  # НЕ из IFC
        Ce=1.0,                  # НЕ из IFC
        Ct=1.0,                  # НЕ из IFC
        roof_angle_deg=0.0,      # НЕ из IFC
    )
    
    print(f"\nИсходные данные:")
    print(f"  Площадь: {result['area_m2']} м2")
    print(f"  Снеговой район: {result['snow_region']} (Sg = {result['Sg_kPa']} кПа)")
    print(f"  Коэффициент mu: {result['mu']}")
    print(f"  Коэффициент Ce: {result['Ce']}")
    print(f"  Коэффициент Ct: {result['Ct']}")
    print(f"\nСтатус: {result['status']}")
    if "error" in result:
        print(f"Ошибка: {result['error']}")
    if "max_result" in result:
        print(f"MaxResult: {result['max_result']}")
    
    # Пример 2: Обработка IFC файла
    jsonl_path = Path(__file__).parent.parent / "output.jsonl"
    
    if jsonl_path.exists():
        print(f"\n--- Обработка IFC файла: {jsonl_path} ---")
        results = process_ifc_file(str(jsonl_path), snow_region="III")
        
        if results:
            for r in results:
                print(f"\nЭлемент: {r['element_name']}")
                print(f"  Площадь: {r['area_m2']} м2 ({r['area_source']})")
                print(f"  Статус: {r['status']}")
        else:
            print("Элементы кровли не найдены в IFC файле")
    
    print("\n" + "=" * 70)
    print("ДАННЫЕ ИЗ IFC:")
    print("  [+] Площадь покрытия (GrossArea/NetArea из Qto)")
    print("  [+] Идентификация элемента (global_id, name)")
    print("\nДАННЫЕ НЕ ИЗ IFC (задаются вручную):")
    print("  [-] snow_region - Снеговой район (I-VIII)")
    print("  [-] mu - Коэффициент перехода (неравномерность)")
    print("  [-] Ce - Коэффициент сноса снега ветром")
    print("  [-] Ct - Термический коэффициент")
    print("  [-] roof_angle_deg - Угол наклона кровли")
    print("=" * 70)
