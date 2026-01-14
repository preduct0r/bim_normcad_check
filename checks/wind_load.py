#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт ветровой нагрузки по СП 20.13330.2016 (Нагрузки и воздействия).

Модуль вычисляет ветровую нагрузку с учётом:
- Ветрового района (w0)
- Типа местности (A, B, C)
- Высотного коэффициента k(ze)
- Аэродинамических коэффициентов (c)
- Пульсационной составляющей

ВАЖНО: Следующие параметры НЕ извлекаются из IFC и заданы по умолчанию:
- wind_region: ветровой район (по умолчанию III)
- terrain_type: тип местности (по умолчанию B)
- building_height: высота здания (может быть извлечена из геометрии)
- aero_coefficients: аэродинамические коэффициенты (требуют CFD или справочники)
"""

from __future__ import annotations
import json
import math
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple
from pathlib import Path


# ============== НОРМАТИВНЫЕ ДАННЫЕ ПО СП 20.13330.2016 ==============

# Таблица 11.1 - Нормативное значение ветрового давления w0 (кПа) по районам
WIND_REGIONS_W0 = {
    "Ia": 0.17,
    "I": 0.23,
    "II": 0.30,
    "III": 0.38,
    "IV": 0.48,
    "V": 0.60,
    "VI": 0.73,
    "VII": 0.85,
}

# Таблица 11.2 - Коэффициент k для типов местности
# k(ze) - коэффициент, учитывающий изменение ветрового давления по высоте
# Формат: {тип_местности: [(ze_max, k, zeta), ...]}
# ze - эквивалентная высота, k - коэффициент, zeta - коэф. пульсации

TERRAIN_PARAMS = {
    "A": {  # Открытые побережья морей, озёр, водохранилищ, пустыни, степи, тундра
        "alpha": 0.15,
        "k10": 1.0,
        "zeta10": 0.76,
        "z0": 5.0,
    },
    "B": {  # Городские территории, лесные массивы, местность с препятствиями h > 10м
        "alpha": 0.20,
        "k10": 0.65,
        "zeta10": 1.06,
        "z0": 10.0,
    },
    "C": {  # Городские районы с застройкой h > 25м
        "alpha": 0.25,
        "k10": 0.40,
        "zeta10": 1.78,
        "z0": 20.0,
    },
}

# Коэффициент надёжности по ветровой нагрузке (п. 11.1.12)
GAMMA_F_WIND = 1.4


# Аэродинамические коэффициенты для типовых конструкций (Приложение В)
# Упрощённые значения для прямоугольных зданий
AERO_COEFFICIENTS = {
    "windward": 0.8,     # Наветренная сторона
    "leeward": -0.5,     # Подветренная сторона  
    "side": -0.6,        # Боковые стены
    "roof_flat": -0.7,   # Плоская кровля (разрежение)
    "roof_windward": -0.6,  # Скатная кровля, наветренный скат
    "roof_leeward": -0.4,   # Скатная кровля, подветренный скат
}


@dataclass
class WindLoadParams:
    """Параметры для расчёта ветровой нагрузки."""
    
    # Извлекаемые из IFC (через psets или геометрию)
    element_id: str = ""
    element_name: str = ""
    element_height_m: float = 0.0  # Высота элемента над землёй
    element_width_m: float = 0.0   # Ширина элемента
    element_area_m2: float = 0.0   # Площадь воздействия
    
    # НЕ извлекаются из IFC - значения по умолчанию
    wind_region: str = "III"       # Ветровой район
    terrain_type: str = "B"        # Тип местности (A, B, C)
    building_height_m: float = 10.0  # Полная высота здания
    building_width_m: float = 20.0   # Ширина здания (поперёк ветра)
    building_length_m: float = 30.0  # Длина здания (вдоль ветра)
    
    # Аэродинамика
    surface_type: str = "windward"  # Тип поверхности для c
    custom_aero_coef: Optional[float] = None  # Пользовательский коэф.
    
    # Учёт пульсации
    include_pulsation: bool = True
    correlation_coef_nu: float = 0.8  # Коэф. пространственной корреляции


@dataclass  
class WindLoadResult:
    """Результаты расчёта ветровой нагрузки."""
    
    element_id: str
    element_name: str
    
    # Нормативные значения
    w0: float           # Базовое ветровое давление, кПа
    k_ze: float         # Высотный коэффициент
    c: float            # Аэродинамический коэффициент
    
    # Средняя составляющая
    wm: float           # Средняя ветровая нагрузка, кПа
    
    # Пульсационная составляющая
    zeta: float         # Коэффициент пульсации
    nu: float           # Коэффициент корреляции
    wp: float           # Пульсационная составляющая, кПа
    
    # Полная нагрузка
    w_norm: float       # Нормативная полная нагрузка, кПа
    w_calc: float       # Расчётная нагрузка, кПа
    gamma_f: float      # Коэффициент надёжности
    
    # Нагрузка на элемент
    area_m2: float
    total_load_kN: float
    
    # Направления ветра
    wind_directions: List[str] = field(default_factory=list)
    
    # Примечания
    notes: List[str] = field(default_factory=list)


def calculate_k_ze(ze: float, terrain_type: str) -> Tuple[float, float]:
    """
    Расчёт коэффициента k(ze) и zeta(ze) по формуле 11.4 СП 20.
    
    k(ze) = k10 × (ze/10)^(2α)  для ze ≥ z0
    k(ze) = k10 × (z0/10)^(2α)  для ze < z0
    
    Returns:
        (k, zeta) - высотный коэффициент и коэффициент пульсации
    """
    params = TERRAIN_PARAMS.get(terrain_type, TERRAIN_PARAMS["B"])
    
    alpha = params["alpha"]
    k10 = params["k10"]
    zeta10 = params["zeta10"]
    z0 = params["z0"]
    
    # Эквивалентная высота не менее z0
    ze_calc = max(ze, z0)
    
    # Коэффициент k
    k = k10 * (ze_calc / 10) ** (2 * alpha)
    
    # Коэффициент пульсации zeta
    zeta = zeta10 * (ze_calc / 10) ** (-alpha)
    
    return k, zeta


def calculate_correlation_coefficient(
    rho: float, 
    chi: float,
    terrain_type: str = "B"
) -> float:
    """
    Расчёт коэффициента пространственной корреляции ν по п. 11.1.11 СП 20.
    
    Args:
        rho: Параметр, учитывающий размер здания поперёк ветра
        chi: Параметр, учитывающий размер здания вдоль ветра
    """
    # Упрощённая формула для прямоугольных зданий
    # ν = 1 / (1 + 0.8 × (ρ + χ))
    nu = 1.0 / (1.0 + 0.8 * (rho + chi))
    return max(0.4, min(1.0, nu))


def get_aero_coefficient(surface_type: str, custom: Optional[float] = None) -> float:
    """Получение аэродинамического коэффициента."""
    if custom is not None:
        return custom
    return AERO_COEFFICIENTS.get(surface_type, 0.8)


def calculate_wind_load(params: WindLoadParams) -> WindLoadResult:
    """
    Расчёт ветровой нагрузки по СП 20.13330.2016.
    
    Формулы:
    wm = w0 × k(ze) × c  - средняя составляющая
    wp = wm × ζ × ν      - пульсационная составляющая
    w = wm + wp          - полная нагрузка
    """
    notes = []
    
    # Базовое ветровое давление
    w0 = WIND_REGIONS_W0.get(params.wind_region, 0.38)
    if params.wind_region not in WIND_REGIONS_W0:
        notes.append(f"Неизвестный ветровой район '{params.wind_region}', принято w0={w0} кПа (район III)")
    else:
        notes.append(f"Ветровой район {params.wind_region}: w0={w0} кПа")
    
    # Тип местности
    if params.terrain_type not in TERRAIN_PARAMS:
        notes.append(f"Неизвестный тип местности '{params.terrain_type}', принят тип B")
        terrain = "B"
    else:
        terrain = params.terrain_type
        notes.append(f"Тип местности: {terrain}")
    
    # Определение эквивалентной высоты
    # Для отдельного элемента используем его высоту
    ze = params.element_height_m if params.element_height_m > 0 else params.building_height_m
    notes.append(f"Эквивалентная высота ze = {ze:.1f} м")
    
    # Высотный коэффициент и коэффициент пульсации
    k_ze, zeta = calculate_k_ze(ze, terrain)
    notes.append(f"Коэффициент k(ze) = {k_ze:.3f}")
    notes.append(f"Коэффициент пульсации ζ(ze) = {zeta:.3f}")
    
    # Аэродинамический коэффициент
    c = get_aero_coefficient(params.surface_type, params.custom_aero_coef)
    notes.append(f"Аэродинамический коэффициент c = {c:.2f} ({params.surface_type})")
    
    # Средняя составляющая ветровой нагрузки
    wm = w0 * k_ze * c
    notes.append(f"Средняя составляющая wm = w0×k×c = {w0}×{k_ze:.3f}×{c:.2f} = {wm:.4f} кПа")
    
    # Пульсационная составляющая
    wp = 0.0
    nu = params.correlation_coef_nu
    
    if params.include_pulsation:
        # Расчёт коэффициента корреляции
        if params.building_width_m > 0 and params.building_height_m > 0:
            # Упрощённый расчёт
            rho = params.building_width_m / (10 * params.building_height_m)
            chi = params.building_length_m / (10 * params.building_height_m)
            nu = calculate_correlation_coefficient(rho, chi, terrain)
        
        wp = abs(wm) * zeta * nu
        notes.append(f"Коэффициент корреляции ν = {nu:.3f}")
        notes.append(f"Пульсационная составляющая wp = wm×ζ×ν = {abs(wm):.4f}×{zeta:.3f}×{nu:.3f} = {wp:.4f} кПа")
    else:
        notes.append("Пульсационная составляющая не учитывается")
    
    # Полная нормативная нагрузка
    # Знак wm сохраняется (отсос имеет отрицательный знак)
    if wm >= 0:
        w_norm = wm + wp
    else:
        w_norm = wm - wp  # Для отсоса wp усиливает эффект
    
    notes.append(f"Нормативная нагрузка w = wm + wp = {w_norm:.4f} кПа")
    
    # Расчётная нагрузка
    gamma_f = GAMMA_F_WIND
    w_calc = w_norm * gamma_f
    notes.append(f"Расчётная нагрузка w×γf = {w_norm:.4f}×{gamma_f} = {w_calc:.4f} кПа")
    
    # Полная нагрузка на элемент
    area = params.element_area_m2 if params.element_area_m2 > 0 else 0.0
    total_load = abs(w_calc) * area
    
    if area > 0:
        notes.append(f"Полная ветровая нагрузка на элемент: {abs(w_calc):.4f} × {area:.2f} = {total_load:.2f} кН")
    
    # Учёт направлений ветра
    wind_directions = ["0°", "90°", "180°", "270°"]
    notes.append("Расчёт выполнен для одного направления ветра. "
                "Требуется проверка для всех направлений.")
    
    return WindLoadResult(
        element_id=params.element_id,
        element_name=params.element_name,
        w0=w0,
        k_ze=k_ze,
        c=c,
        wm=wm,
        zeta=zeta,
        nu=nu,
        wp=wp,
        w_norm=w_norm,
        w_calc=w_calc,
        gamma_f=gamma_f,
        area_m2=area,
        total_load_kN=total_load,
        wind_directions=wind_directions,
        notes=notes
    )


def extract_wind_params_from_ifc_element(element: Dict[str, Any]) -> WindLoadParams:
    """
    Извлечение параметров из элемента IFC для расчёта ветровой нагрузки.
    
    ВАЖНО: Большинство параметров НЕ содержатся в типичных IFC файлах.
    """
    params = WindLoadParams()
    
    # Базовая идентификация
    params.element_id = element.get("global_id", "")
    params.element_name = element.get("name", "")
    
    # Попытка извлечь координаты (высоту)
    placement = element.get("placement", {})
    xyz = placement.get("xyz", [0, 0, 0])
    if len(xyz) >= 3:
        params.element_height_m = xyz[2] / 1000.0  # Предполагаем мм -> м
    
    # Попытка извлечь размеры из section_approx
    section = element.get("section_approx", {})
    if section:
        params.element_width_m = section.get("b", 0) / 1000.0
        
    # Попытка извлечь площадь
    psets = element.get("psets", {})
    
    # Для стен
    qto_wall = psets.get("Qto_WallBaseQuantities", {})
    if "GrossSideArea" in qto_wall:
        params.element_area_m2 = float(qto_wall["GrossSideArea"])
    
    return params


def calculate_wind_for_building(
    building_height: float,
    building_width: float,
    building_length: float,
    wind_region: str = "III",
    terrain_type: str = "B",
) -> Dict[str, WindLoadResult]:
    """
    Расчёт ветровых нагрузок на все поверхности здания.
    
    Returns:
        Словарь с результатами для каждой поверхности
    """
    results = {}
    
    surfaces = [
        ("windward", building_width * building_height, building_height / 2),
        ("leeward", building_width * building_height, building_height / 2),
        ("side", building_length * building_height, building_height / 2),
        ("roof_flat", building_width * building_length, building_height),
    ]
    
    for surface_type, area, height in surfaces:
        params = WindLoadParams(
            element_id=f"surface_{surface_type}",
            element_name=f"Поверхность: {surface_type}",
            element_height_m=height,
            element_area_m2=area,
            wind_region=wind_region,
            terrain_type=terrain_type,
            building_height_m=building_height,
            building_width_m=building_width,
            building_length_m=building_length,
            surface_type=surface_type,
        )
        
        results[surface_type] = calculate_wind_load(params)
    
    return results


# ============== ИНТЕГРАЦИЯ С NormCAD API ==============

def prepare_normcad_wind_data(result: WindLoadResult) -> Dict[str, Any]:
    """
    Подготовка данных для передачи в NormCAD API.
    """
    return {
        "Norm": "СП 20.13330.2016",
        "TaskName": f"Ветровая нагрузка - {result.element_name}",
        "Unit": "п.11",
        
        # Исходные данные
        "W0": result.w0,          # Базовое давление
        "K_ze": result.k_ze,      # Высотный коэффициент
        "C": result.c,            # Аэродинамический коэффициент
        "Zeta": result.zeta,      # Коэффициент пульсации
        "Nu": result.nu,          # Коэффициент корреляции
        "GammaF": result.gamma_f, # Коэффициент надёжности
        
        # Результаты
        "Wm": result.wm,          # Средняя составляющая
        "Wp": result.wp,          # Пульсационная составляющая
        "W_norm": result.w_norm,  # Нормативная нагрузка
        "W_calc": result.w_calc,  # Расчётная нагрузка
        "Area": result.area_m2,
        "TotalLoad": result.total_load_kN,
    }


# ============== ПРИМЕР ИСПОЛЬЗОВАНИЯ ==============

if __name__ == "__main__":
    print("=" * 70)
    print("РАСЧЁТ ВЕТРОВОЙ НАГРУЗКИ ПО СП 20.13330.2016")
    print("=" * 70)
    
    # Демонстрационный расчёт для здания
    building_h = 15.0  # м
    building_w = 20.0  # м
    building_l = 30.0  # м
    
    print(f"\nПараметры здания:")
    print(f"  Высота: {building_h} м")
    print(f"  Ширина: {building_w} м")
    print(f"  Длина: {building_l} м")
    print(f"  Ветровой район: III (w0 = 0.38 кПа)")
    print(f"  Тип местности: B (городская территория)")
    
    results = calculate_wind_for_building(
        building_height=building_h,
        building_width=building_w,
        building_length=building_l,
        wind_region="III",
        terrain_type="B"
    )
    
    print("\n" + "-" * 70)
    print("РЕЗУЛЬТАТЫ РАСЧЁТА:")
    print("-" * 70)
    
    for surface, result in results.items():
        print(f"\n{surface.upper()}:")
        print(f"  Площадь: {result.area_m2:.1f} м²")
        print(f"  Аэродинамический коэффициент c = {result.c:.2f}")
        print(f"  Средняя составляющая wm = {result.wm:.4f} кПа")
        print(f"  Пульсационная составляющая wp = {result.wp:.4f} кПа")
        print(f"  Нормативная нагрузка w = {result.w_norm:.4f} кПа")
        print(f"  Расчётная нагрузка = {result.w_calc:.4f} кПа")
        print(f"  Полная нагрузка = {result.total_load_kN:.2f} кН")
    
    print("\n" + "=" * 70)
    print("НЕДОСТАЮЩИЕ ДАННЫЕ ИЗ IFC:")
    print("=" * 70)
    print("""
Следующие параметры НЕ извлекаются из типичного IFC файла:

1. wind_region - Ветровой район строительства
   → Определяется по карте районирования РФ (рис. 11.1 СП 20)
   → По умолчанию: III район (w0 = 0.38 кПа)

2. terrain_type - Тип местности (A, B, C)
   → A - открытая местность (побережья, степи)
   → B - городская территория, лес
   → C - городские районы с высотной застройкой
   → По умолчанию: B

3. building_height/width/length - Габариты здания
   → Могут быть вычислены из bbox всех элементов IFC
   → По умолчанию: требуется задать вручную

4. surface_type - Тип поверхности для аэродинамического коэффициента
   → Определяет знак и величину c
   → По умолчанию: 'windward' (наветренная сторона)

5. Аэродинамические коэффициенты (c)
   → Для сложных форм требуется CFD моделирование
   → Или справочные данные из Приложения В СП 20
   → По умолчанию: типовые значения для прямоугольного здания

6. Динамические характеристики (для зданий h > 40м)
   → Собственные частоты, логарифмический декремент
   → Требуется специальный расчёт
""")
