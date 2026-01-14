#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт снеговой нагрузки по СП 20.13330.2016 (Нагрузки и воздействия).

Модуль вычисляет снеговую нагрузку с учётом:
- Снегового района (Sg)
- Коэффициента mu (учёт неравномерности распределения снега)
- Коэффициентов надёжности по нагрузке

ВАЖНО: Следующие параметры НЕ извлекаются из IFC и заданы по умолчанию:
- snow_region: снеговой район (по умолчанию III)
- roof_angle: угол наклона кровли (по умолчанию 0deg)
- roof_type: тип кровли для определения mu (по умолчанию 'flat')
- Ce: коэффициент, учитывающий снос снега с покрытий (по умолчанию 1.0)
- Ct: термический коэффициент (по умолчанию 1.0)
"""

from __future__ import annotations
import json
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional
from pathlib import Path


# ============== НОРМАТИВНЫЕ ДАННЫЕ ПО СП 20.13330.2016 ==============

# Таблица 10.1 - Вес снегового покрова Sg (кПа) по снеговым районам
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

# Коэффициент надёжности по снеговой нагрузке (п. 10.12)
GAMMA_F_SNOW = 1.4  # для расчёта по предельным состояниям первой группы


@dataclass
class SnowLoadParams:
    """Параметры для расчёта снеговой нагрузки."""
    
    # Извлекаемые из IFC (через psets или геометрию)
    element_id: str = ""
    element_name: str = ""
    area_m2: float = 0.0  # Площадь покрытия (из Qto или геометрии)
    
    # НЕ извлекаются из IFC - значения по умолчанию
    snow_region: str = "III"  # Снеговой район (I-VIII)
    roof_angle_deg: float = 0.0  # Угол наклона кровли, градусы
    roof_type: str = "flat"  # Тип кровли: 'flat', 'single_slope', 'gable', 'arch'
    Ce: float = 1.0  # Коэффициент сноса снега ветром
    Ct: float = 1.0  # Термический коэффициент
    
    # Для учёта неравномерности (Приложение Б)
    parapet_height_m: float = 0.0  # Высота парапета
    adjacent_roof: bool = False  # Наличие перепада высот
    height_difference_m: float = 0.0  # Перепад высот


@dataclass
class SnowLoadResult:
    """Результаты расчёта снеговой нагрузки."""
    
    element_id: str
    element_name: str
    
    # Нормативные значения
    Sg: float  # Вес снегового покрова по району, кПа
    mu: float  # Коэффициент перехода (учёт неравномерности)
    Ce: float  # Коэффициент сноса
    Ct: float  # Термический коэффициент
    
    # Расчётные нагрузки
    S0: float  # Нормативное значение снеговой нагрузки, кПа
    S: float   # Расчётное значение снеговой нагрузки, кПа
    gamma_f: float  # Коэффициент надёжности
    
    # Полная нагрузка на элемент
    area_m2: float  # Площадь
    total_load_kN: float  # Полная нагрузка, кН
    
    # Примечания
    notes: List[str] = field(default_factory=list)


def calculate_mu(params: SnowLoadParams) -> tuple[float, List[str]]:
    """
    Определение коэффициента mu по СП 20 (Приложение Б).
    
    Возвращает (mu, notes) - коэффициент и примечания о его определении.
    """
    notes = []
    alpha = params.roof_angle_deg
    
    # Таблица Б.1 - Коэффициент mu для покрытий без перепада высот
    if params.roof_type == "flat":
        # Плоская кровля (alpha <= 25deg)
        if alpha <= 25:
            mu = 1.0
            notes.append(f"Плоская кровля (alpha={alpha}deg <= 25deg): mu=1.0 по табл. Б.1")
        elif alpha <= 60:
            # Линейная интерполяция mu = (60 - alpha) / 35
            mu = (60 - alpha) / 35
            notes.append(f"Скатная кровля (25deg < alpha={alpha}deg <= 60deg): mu={(60-alpha)/35:.2f}")
        else:
            mu = 0.0
            notes.append(f"Крутая кровля (alpha={alpha}deg > 60deg): mu=0 (снег не задерживается)")
    
    elif params.roof_type == "single_slope":
        # Односкатное покрытие
        if alpha <= 30:
            mu = 1.0
        elif alpha <= 60:
            mu = (60 - alpha) / 30
        else:
            mu = 0.0
        notes.append(f"Односкатное покрытие: mu={mu:.2f}")
    
    elif params.roof_type == "gable":
        # Двускатное покрытие (возможна неравномерность)
        if alpha <= 30:
            mu = 1.0  # Равномерная нагрузка
            notes.append(f"Двускатное покрытие (alpha={alpha}deg <= 30deg): равномерная нагрузка, mu=1.0")
        else:
            mu = (60 - alpha) / 30 if alpha <= 60 else 0.0
            notes.append(f"Двускатное покрытие: mu={mu:.2f}")
        
        # Неравномерный вариант (п. Б.4)
        notes.append("ВНИМАНИЕ: Для двускатных кровель требуется дополнительная проверка "
                    "неравномерного распределения (mu_max до 1.25 для alpha<=30deg)")
    
    elif params.roof_type == "arch":
        # Сводчатое покрытие - приблизительно
        mu = 1.0
        notes.append("Сводчатое покрытие: принято mu=1.0 (требуется детальный расчёт по Б.8)")
    
    else:
        mu = 1.0
        notes.append(f"Неизвестный тип кровли '{params.roof_type}': принято mu=1.0")
    
    # Учёт перепада высот (Приложение Б, схемы Б.11-Б.14)
    if params.adjacent_roof and params.height_difference_m > 0:
        # Упрощённый учёт снегового мешка
        h = params.height_difference_m
        # mu может достигать 2.5-4.0 в зоне мешка
        mu_local = min(2.0 * h / 1.0 + 1.0, 4.0)  # Упрощённая формула
        notes.append(f"ВНИМАНИЕ: Перепад высот {h:.1f}м - локальный mu может достигать {mu_local:.1f}")
        notes.append("Требуется детальный расчёт снегового мешка по прил. Б")
    
    # Учёт парапета
    if params.parapet_height_m > 0.6:
        notes.append(f"Парапет h={params.parapet_height_m:.2f}м - возможно образование снегового мешка")
    
    return mu, notes


def calculate_snow_load(params: SnowLoadParams) -> SnowLoadResult:
    """
    Расчёт снеговой нагрузки по СП 20.13330.2016.
    
    Формула: S = Sg x mu x Ce x Ct x gamma_f
    где:
        S - расчётное значение снеговой нагрузки
        Sg - вес снегового покрова (по району)
        mu - коэффициент перехода
        Ce - коэффициент сноса
        Ct - термический коэффициент
        gamma_f - коэффициент надёжности
    """
    notes = []
    
    # Получаем Sg по снеговому району
    Sg = SNOW_REGIONS_SG.get(params.snow_region, 1.5)
    if params.snow_region not in SNOW_REGIONS_SG:
        notes.append(f"Неизвестный снеговой район '{params.snow_region}', принято Sg={Sg} кПа (район III)")
    else:
        notes.append(f"Снеговой район {params.snow_region}: Sg={Sg} кПа")
    
    # Определяем коэффициент mu
    mu, mu_notes = calculate_mu(params)
    notes.extend(mu_notes)
    
    # Коэффициенты
    Ce = params.Ce
    Ct = params.Ct
    
    if Ce != 1.0:
        notes.append(f"Коэффициент сноса Ce={Ce} (учтён снос снега ветром)")
    if Ct != 1.0:
        notes.append(f"Термический коэффициент Ct={Ct}")
    
    # Нормативное значение снеговой нагрузки
    S0 = Sg * mu * Ce * Ct
    notes.append(f"S0 = SgxmuxCexCt = {Sg}x{mu:.2f}x{Ce}x{Ct} = {S0:.3f} кПа")
    
    # Расчётное значение
    gamma_f = GAMMA_F_SNOW
    S = S0 * gamma_f
    notes.append(f"S = S0xgamma_f = {S0:.3f}x{gamma_f} = {S:.3f} кПа")
    
    # Полная нагрузка на элемент
    area = params.area_m2 if params.area_m2 > 0 else 0.0
    total_load = S * area  # кПа x м2 = кН
    if area > 0:
        notes.append(f"Полная снеговая нагрузка: {S:.3f} x {area:.2f} = {total_load:.2f} кН")
    
    return SnowLoadResult(
        element_id=params.element_id,
        element_name=params.element_name,
        Sg=Sg,
        mu=mu,
        Ce=Ce,
        Ct=Ct,
        S0=S0,
        S=S,
        gamma_f=gamma_f,
        area_m2=area,
        total_load_kN=total_load,
        notes=notes
    )


def extract_snow_params_from_ifc_element(element: Dict[str, Any]) -> SnowLoadParams:
    """
    Извлечение параметров из элемента IFC.
    
    ВАЖНО: Большинство параметров для расчёта снега НЕ содержатся в типичных IFC файлах.
    Функция извлекает то, что возможно, и использует значения по умолчанию для остального.
    """
    params = SnowLoadParams()
    
    # Базовая идентификация
    params.element_id = element.get("global_id", "")
    params.element_name = element.get("name", "")
    
    # Попытка извлечь площадь из Qto
    psets = element.get("psets", {})
    
    # Для плит (IfcSlab) - площадь покрытия
    qto = psets.get("Qto_SlabBaseQuantities", {})
    if "GrossArea" in qto:
        params.area_m2 = float(qto["GrossArea"])
    elif "NetArea" in qto:
        params.area_m2 = float(qto["NetArea"])
    
    # Попытка определить угол наклона (редко в IFC)
    # Обычно требуется анализ геометрии
    
    # Pset для кровли (если есть)
    roof_pset = psets.get("Pset_RoofCommon", {})
    if "ProjectedArea" in roof_pset:
        params.area_m2 = float(roof_pset["ProjectedArea"])
    
    return params


def process_ifc_data(jsonl_path: str, output_path: Optional[str] = None,
                     default_params: Optional[Dict[str, Any]] = None) -> List[SnowLoadResult]:
    """
    Обработка JSONL файла с данными IFC и расчёт снеговых нагрузок.
    
    Args:
        jsonl_path: Путь к JSONL файлу с извлечёнными данными IFC
        output_path: Путь для сохранения результатов (опционально)
        default_params: Параметры по умолчанию для расчёта
    
    Returns:
        Список результатов расчёта
    """
    results = []
    defaults = default_params or {}
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            if not line.strip():
                continue
            
            element = json.loads(line)
            
            # Обрабатываем только элементы покрытия (плиты кровли)
            ifc_class = element.get("ifc_class", "")
            if ifc_class not in ["IfcSlab", "IfcRoof"]:
                continue
            
            # Проверяем, что это кровля
            predefined_type = element.get("predefined_type", "")
            psets = element.get("psets", {})
            is_roof = (predefined_type == "ROOF" or 
                      "Roof" in element.get("name", "") or
                      "Кровля" in element.get("name", ""))
            
            if not is_roof and ifc_class != "IfcRoof":
                continue
            
            params = extract_snow_params_from_ifc_element(element)
            
            # Применяем значения по умолчанию
            for key, value in defaults.items():
                if hasattr(params, key):
                    setattr(params, key, value)
            
            result = calculate_snow_load(params)
            results.append(result)
    
    # Сохранение результатов
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            for r in results:
                f.write(json.dumps(r.__dict__, ensure_ascii=False, indent=2))
                f.write("\n")
    
    return results


# ============== ИНТЕГРАЦИЯ С NormCAD API ==============

def prepare_normcad_snow_data(result: SnowLoadResult) -> Dict[str, Any]:
    """
    Подготовка данных для передачи в NormCAD API.
    
    Формирует словарь переменных для модуля "Нагрузки и воздействия".
    """
    return {
        "Norm": "СП 20.13330.2016",
        "TaskName": f"Снеговая нагрузка - {result.element_name}",
        "Unit": "п.10",
        
        # Исходные данные
        "SnowRegion": result.Sg,  # Вес снегового покрова
        "Mu": result.mu,          # Коэффициент mu
        "Ce": result.Ce,          # Коэффициент сноса
        "Ct": result.Ct,          # Термический коэффициент
        "GammaF": result.gamma_f, # Коэффициент надёжности
        
        # Результаты
        "S0": result.S0,          # Нормативное значение
        "S": result.S,            # Расчётное значение
        "Area": result.area_m2,
        "TotalLoad": result.total_load_kN,
    }


# ============== ПРИМЕР ИСПОЛЬЗОВАНИЯ ==============

if __name__ == "__main__":
    import sys
    
    # Пример расчёта с параметрами по умолчанию
    print("=" * 60)
    print("РАСЧЁТ СНЕГОВОЙ НАГРУЗКИ ПО СП 20.13330.2016")
    print("=" * 60)
    
    # Демонстрационный расчёт
    demo_params = SnowLoadParams(
        element_id="demo_roof_001",
        element_name="Плита покрытия",
        area_m2=100.0,  # 100 м2
        snow_region="III",  # Снеговой район III (Sg=1.5 кПа)
        roof_angle_deg=10.0,  # Угол наклона 10deg
        roof_type="flat",
        Ce=1.0,
        Ct=1.0,
    )
    
    result = calculate_snow_load(demo_params)
    
    print(f"\nЭлемент: {result.element_name} ({result.element_id})")
    print(f"\nИсходные данные:")
    print(f"  Снеговой район: III (Sg = {result.Sg} кПа)")
    print(f"  Коэффициент mu = {result.mu:.2f}")
    print(f"  Коэффициент Ce = {result.Ce}")
    print(f"  Коэффициент Ct = {result.Ct}")
    print(f"  Площадь покрытия = {result.area_m2} м2")
    
    print(f"\nРезультаты:")
    print(f"  Нормативная снеговая нагрузка S0 = {result.S0:.3f} кПа")
    print(f"  Расчётная снеговая нагрузка S = {result.S:.3f} кПа")
    print(f"  Полная нагрузка на покрытие = {result.total_load_kN:.2f} кН")
    
    print(f"\nПримечания:")
    for note in result.notes:
        print(f"  • {note}")
    
    print("\n" + "=" * 60)
    print("НЕДОСТАЮЩИЕ ДАННЫЕ ИЗ IFC:")
    print("=" * 60)
    print("""
Следующие параметры НЕ извлекаются из типичного IFC файла
и должны быть заданы вручную или получены из других источников:

1. snow_region - Снеговой район строительства
   -> Определяется по карте районирования РФ (рис. 10.1 СП 20)
   -> По умолчанию: III район (Sg = 1.5 кПа)

2. roof_angle_deg - Угол наклона кровли
   -> Может быть вычислен из геометрии IFC
   -> По умолчанию: 0deg (плоская кровля)

3. roof_type - Тип кровли
   -> Определяет схему распределения снега
   -> По умолчанию: 'flat'

4. Ce - Коэффициент сноса снега ветром
   -> Зависит от скорости ветра и рельефа
   -> По умолчанию: 1.0

5. Ct - Термический коэффициент
   -> Учитывает теплопотери через покрытие
   -> По умолчанию: 1.0

6. Параметры неравномерности (снеговые мешки):
   - parapet_height_m - высота парапета
   - adjacent_roof - наличие перепада высот
   - height_difference_m - величина перепада
""")
