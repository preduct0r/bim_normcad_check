#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт собственного веса конструкций по СП 20.13330.2016.

Модуль вычисляет нагрузку от собственного веса конструкций на основе:
- Объёма элементов (из Qto)
- Плотности материалов
- Коэффициентов надёжности по нагрузке

Данные, которые ИЗВЛЕКАЮТСЯ из IFC:
- Объём элементов (NetVolume из Qto_*BaseQuantities)
- Наименование материала
- Геометрические размеры (длина, площадь сечения)

Данные, которые НЕ извлекаются из IFC (значения по умолчанию):
- Плотность материала (используется справочная по названию)
- Коэффициент надёжности по нагрузке
"""

from __future__ import annotations
import json
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple
from pathlib import Path


# ============== СПРАВОЧНЫЕ ДАННЫЕ ПО ПЛОТНОСТИ МАТЕРИАЛОВ ==============

# Плотности материалов, кг/м3 (из СП 20, Приложение Г и справочников)
MATERIAL_DENSITIES = {
    # Древесина
    "wood": 500,
    "wood_spruce": 450,
    "wood_spruce_beam": 450,  # Ель для балок
    "wood_pine": 500,
    "wood_oak": 700,
    "timber": 500,
    "glulam": 500,  # Клеёный брус
    "lvl": 480,     # LVL брус
    
    # Сталь
    "steel": 7850,
    "steel_s235": 7850,
    "steel_s345": 7850,
    "steel_structural": 7850,
    
    # Бетон
    "concrete": 2500,
    "concrete_b25": 2500,
    "concrete_b30": 2500,
    "reinforced_concrete": 2500,
    "lightweight_concrete": 1800,
    
    # Кирпич
    "brick": 1800,
    "brick_ceramic": 1700,
    "brick_silicate": 1800,
    
    # Прочие
    "aluminum": 2700,
    "glass": 2500,
    "gypsum": 1200,
    "insulation": 100,
}

# Коэффициент надёжности по постоянным нагрузкам (п. 7.2 СП 20)
# gamma_f = 1.1 при неблагоприятном действии (увеличение веса)
# gamma_f = 0.9 при благоприятном действии (уменьшение веса)
GAMMA_F_PERMANENT_UNFAV = 1.1
GAMMA_F_PERMANENT_FAV = 0.9


@dataclass
class SelfWeightParams:
    """Параметры для расчёта собственного веса."""
    
    # Извлекаемые из IFC
    element_id: str = ""
    element_name: str = ""
    ifc_class: str = ""
    
    # Геометрия (из Qto или section_approx)
    volume_m3: float = 0.0          # Объём, м3
    length_m: float = 0.0           # Длина (для линейных элементов), м
    cross_section_area_m2: float = 0.0  # Площадь сечения, м2
    
    # Материал
    material_name: str = ""         # Название материала из IFC
    density_kg_m3: Optional[float] = None  # Плотность (если известна)
    
    # Коэффициенты
    gamma_f: float = GAMMA_F_PERMANENT_UNFAV


@dataclass
class SelfWeightResult:
    """Результаты расчёта собственного веса."""
    
    element_id: str
    element_name: str
    ifc_class: str
    
    # Геометрия
    volume_m3: float
    length_m: float
    cross_section_area_m2: float
    
    # Материал
    material_name: str
    density_kg_m3: float
    density_source: str  # "ifc" или "default"
    
    # Нагрузки
    mass_kg: float              # Масса, кг
    weight_norm_kN: float       # Нормативный вес, кН
    weight_calc_kN: float       # Расчётный вес, кН
    gamma_f: float
    
    # Распределённая нагрузка (для линейных элементов)
    linear_load_kN_m: Optional[float] = None  # кН/м
    
    # Примечания
    notes: List[str] = field(default_factory=list)


def normalize_material_name(name: str) -> str:
    """Нормализация названия материала для поиска в справочнике."""
    if not name:
        return ""
    
    # Приведение к нижнему регистру и замена разделителей
    normalized = name.lower().strip()
    normalized = re.sub(r'[\s\-_]+', '_', normalized)
    
    return normalized


def get_material_density(material_name: str, custom_density: Optional[float] = None) -> Tuple[float, str]:
    """
    Определение плотности материала.
    
    Returns:
        (density, source) - плотность и источник ("ifc", "lookup", "default")
    """
    # Если задана пользовательская плотность
    if custom_density is not None and custom_density > 0:
        return custom_density, "custom"
    
    if not material_name:
        return 2500.0, "default"  # Бетон по умолчанию
    
    normalized = normalize_material_name(material_name)
    
    # Точное совпадение
    if normalized in MATERIAL_DENSITIES:
        return MATERIAL_DENSITIES[normalized], "lookup"
    
    # Поиск по ключевым словам
    keywords_density = [
        (["steel", "сталь"], 7850),
        (["wood", "timber", "древ", "брус", "дерев", "spruce", "pine", "ель", "сосна"], 500),
        (["concrete", "бетон"], 2500),
        (["brick", "кирпич"], 1800),
        (["aluminum", "алюмин"], 2700),
        (["glass", "стекло"], 2500),
        (["gypsum", "гипс"], 1200),
    ]
    
    for keywords, density in keywords_density:
        for kw in keywords:
            if kw in normalized:
                return density, "lookup"
    
    # По умолчанию - бетон
    return 2500.0, "default"


def calculate_self_weight(params: SelfWeightParams) -> SelfWeightResult:
    """
    Расчёт собственного веса элемента.
    
    Формулы:
    G = ro x V x g  - вес элемента
    где:
        ro - плотность материала, кг/м3
        V - объём элемента, м3
        g ≈ 10 м/с² (или 9.81 для точного расчёта)
    """
    notes = []
    g = 9.81 / 1000  # м/с² -> кН/кг
    
    # Определение плотности
    density, density_source = get_material_density(params.material_name, params.density_kg_m3)
    
    if density_source == "default":
        notes.append(f"ВНИМАНИЕ: Плотность материала '{params.material_name}' не найдена, "
                    f"принято ro={density} кг/м3 (бетон)")
    elif density_source == "lookup":
        notes.append(f"Материал: {params.material_name}, плотность ro={density} кг/м3 (из справочника)")
    else:
        notes.append(f"Материал: {params.material_name}, плотность ro={density} кг/м3 (задано)")
    
    # Определение объёма
    volume = params.volume_m3
    
    if volume <= 0 and params.length_m > 0 and params.cross_section_area_m2 > 0:
        # Расчёт объёма по длине и площади сечения
        volume = params.length_m * params.cross_section_area_m2
        notes.append(f"Объём рассчитан: V = L x A = {params.length_m:.3f} x {params.cross_section_area_m2:.6f} = {volume:.6f} м3")
    elif volume > 0:
        notes.append(f"Объём из IFC: V = {volume:.6f} м3")
    else:
        notes.append("ВНИМАНИЕ: Объём элемента не определён!")
    
    # Расчёт массы и веса
    mass = density * volume  # кг
    weight_norm = mass * g   # кН
    
    notes.append(f"Масса: m = ro x V = {density} x {volume:.6f} = {mass:.2f} кг")
    notes.append(f"Нормативный вес: G0 = m x g = {mass:.2f} x {g:.4f} = {weight_norm:.4f} кН")
    
    # Расчётный вес
    gamma_f = params.gamma_f
    weight_calc = weight_norm * gamma_f
    notes.append(f"Расчётный вес: G = G0 x gamma_f = {weight_norm:.4f} x {gamma_f} = {weight_calc:.4f} кН")
    
    # Линейная нагрузка для балок/колонн
    linear_load = None
    if params.length_m > 0:
        linear_load = weight_calc / params.length_m
        notes.append(f"Погонная нагрузка: q = G / L = {weight_calc:.4f} / {params.length_m:.3f} = {linear_load:.4f} кН/м")
    
    return SelfWeightResult(
        element_id=params.element_id,
        element_name=params.element_name,
        ifc_class=params.ifc_class,
        volume_m3=volume,
        length_m=params.length_m,
        cross_section_area_m2=params.cross_section_area_m2,
        material_name=params.material_name,
        density_kg_m3=density,
        density_source=density_source,
        mass_kg=mass,
        weight_norm_kN=weight_norm,
        weight_calc_kN=weight_calc,
        gamma_f=gamma_f,
        linear_load_kN_m=linear_load,
        notes=notes
    )


def extract_self_weight_params_from_ifc_element(element: Dict[str, Any]) -> SelfWeightParams:
    """
    Извлечение параметров из элемента IFC для расчёта собственного веса.
    """
    params = SelfWeightParams()
    
    # Базовая идентификация
    params.element_id = element.get("global_id", "")
    params.element_name = element.get("name", "")
    params.ifc_class = element.get("ifc_class", "")
    
    # Материал
    materials = element.get("materials", [])
    if materials:
        params.material_name = materials[0]
    
    # Геометрия из psets
    psets = element.get("psets", {})
    
    # Для балок
    qto_beam = psets.get("Qto_BeamBaseQuantities", {})
    if qto_beam:
        if "NetVolume" in qto_beam:
            params.volume_m3 = float(qto_beam["NetVolume"])
        if "Length" in qto_beam:
            # Длина может быть в мм или м, проверяем порядок величины
            length_val = float(qto_beam["Length"])
            params.length_m = length_val / 1000.0 if length_val > 100 else length_val
        if "CrossSectionArea" in qto_beam:
            params.cross_section_area_m2 = float(qto_beam["CrossSectionArea"])
    
    # Для колонн
    qto_column = psets.get("Qto_ColumnBaseQuantities", {})
    if qto_column:
        if "NetVolume" in qto_column:
            params.volume_m3 = float(qto_column["NetVolume"])
        if "Length" in qto_column or "Height" in qto_column:
            length_val = float(qto_column.get("Length", qto_column.get("Height", 0)))
            params.length_m = length_val / 1000.0 if length_val > 100 else length_val
        if "CrossSectionArea" in qto_column:
            params.cross_section_area_m2 = float(qto_column["CrossSectionArea"])
    
    # Для плит
    qto_slab = psets.get("Qto_SlabBaseQuantities", {})
    if qto_slab:
        if "NetVolume" in qto_slab:
            params.volume_m3 = float(qto_slab["NetVolume"])
    
    # Для стен
    qto_wall = psets.get("Qto_WallBaseQuantities", {})
    if qto_wall:
        if "NetVolume" in qto_wall:
            params.volume_m3 = float(qto_wall["NetVolume"])
    
    # Дополнительно из section_approx
    section = element.get("section_approx", {})
    if section:
        if params.length_m == 0 and "length" in section:
            length_val = float(section["length"])
            params.length_m = length_val / 1000.0 if length_val > 100 else length_val
        
        if params.cross_section_area_m2 == 0:
            # Расчёт площади из b x h
            b = section.get("b", 0)
            h = section.get("h", 0)
            if b > 0 and h > 0:
                # Переводим из мм в м
                b_m = b / 1000.0 if b > 1 else b
                h_m = h / 1000.0 if h > 1 else h
                params.cross_section_area_m2 = b_m * h_m
    
    return params


def process_ifc_data(jsonl_path: str, output_path: Optional[str] = None) -> List[SelfWeightResult]:
    """
    Обработка JSONL файла с данными IFC и расчёт собственного веса.
    """
    results = []
    total_weight_norm = 0.0
    total_weight_calc = 0.0
    
    with open(jsonl_path, 'r', encoding='utf-8') as f:
        for line in f:
            if not line.strip():
                continue
            
            element = json.loads(line)
            
            # Обрабатываем конструктивные элементы
            ifc_class = element.get("ifc_class", "")
            if ifc_class not in ["IfcBeam", "IfcColumn", "IfcSlab", "IfcWall", "IfcMember", "IfcFooting"]:
                continue
            
            params = extract_self_weight_params_from_ifc_element(element)
            result = calculate_self_weight(params)
            results.append(result)
            
            total_weight_norm += result.weight_norm_kN
            total_weight_calc += result.weight_calc_kN
    
    # Добавляем итоговую информацию
    if results:
        results[0].notes.append(f"\n=== ИТОГО по всем элементам ===")
        results[0].notes.append(f"Суммарный нормативный вес: {total_weight_norm:.2f} кН")
        results[0].notes.append(f"Суммарный расчётный вес: {total_weight_calc:.2f} кН")
    
    # Сохранение результатов
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            for r in results:
                f.write(json.dumps(r.__dict__, ensure_ascii=False, indent=2))
                f.write("\n")
    
    return results


# ============== ИНТЕГРАЦИЯ С NormCAD API ==============

def prepare_normcad_self_weight_data(result: SelfWeightResult) -> Dict[str, Any]:
    """
    Подготовка данных для передачи в NormCAD API.
    """
    return {
        "Norm": "СП 20.13330.2016",
        "TaskName": f"Собственный вес - {result.element_name}",
        "Unit": "п.7",
        
        # Исходные данные
        "Material": result.material_name,
        "Density": result.density_kg_m3,
        "Volume": result.volume_m3,
        "Length": result.length_m,
        "CrossSection": result.cross_section_area_m2,
        "GammaF": result.gamma_f,
        
        # Результаты
        "Mass": result.mass_kg,
        "WeightNorm": result.weight_norm_kN,
        "WeightCalc": result.weight_calc_kN,
        "LinearLoad": result.linear_load_kN_m,
    }


# ============== ПРИМЕР ИСПОЛЬЗОВАНИЯ ==============

if __name__ == "__main__":
    import sys
    
    print("=" * 70)
    print("РАСЧЁТ СОБСТВЕННОГО ВЕСА КОНСТРУКЦИЙ ПО СП 20.13330.2016")
    print("=" * 70)
    
    # Проверяем наличие входного файла
    jsonl_path = Path(__file__).parent.parent / "output.jsonl"
    
    if jsonl_path.exists():
        print(f"\nОбработка файла: {jsonl_path}")
        print("-" * 70)
        
        results = process_ifc_data(str(jsonl_path))
        
        total_weight = 0.0
        for result in results:
            print(f"\n{result.ifc_class}: {result.element_name}")
            print(f"  ID: {result.element_id}")
            print(f"  Материал: {result.material_name} (ro={result.density_kg_m3} кг/м3)")
            print(f"  Объём: {result.volume_m3:.6f} м3")
            print(f"  Длина: {result.length_m:.3f} м")
            print(f"  Масса: {result.mass_kg:.2f} кг")
            print(f"  Нормативный вес: {result.weight_norm_kN:.4f} кН")
            print(f"  Расчётный вес: {result.weight_calc_kN:.4f} кН")
            if result.linear_load_kN_m:
                print(f"  Погонная нагрузка: {result.linear_load_kN_m:.4f} кН/м")
            total_weight += result.weight_calc_kN
        
        print("\n" + "=" * 70)
        print(f"ИТОГО расчётный вес конструкций: {total_weight:.2f} кН")
        print("=" * 70)
    
    else:
        # Демонстрационный расчёт
        print("\nФайл output.jsonl не найден. Демонстрационный расчёт:")
        print("-" * 70)
        
        demo_params = SelfWeightParams(
            element_id="demo_beam_001",
            element_name="Балка деревянная 100x200",
            ifc_class="IfcBeam",
            volume_m3=0.054,  # 2.7м x 0.1м x 0.2м
            length_m=2.7,
            cross_section_area_m2=0.02,  # 100мм x 200мм
            material_name="wood_spruce_beam",
        )
        
        result = calculate_self_weight(demo_params)
        
        print(f"\nЭлемент: {result.element_name}")
        print(f"Класс IFC: {result.ifc_class}")
        print(f"\nИсходные данные:")
        print(f"  Материал: {result.material_name}")
        print(f"  Плотность: {result.density_kg_m3} кг/м3 ({result.density_source})")
        print(f"  Объём: {result.volume_m3:.6f} м3")
        print(f"  Длина: {result.length_m:.3f} м")
        print(f"  Площадь сечения: {result.cross_section_area_m2:.6f} м2")
        
        print(f"\nРезультаты:")
        print(f"  Масса: {result.mass_kg:.2f} кг")
        print(f"  Нормативный вес: {result.weight_norm_kN:.4f} кН")
        print(f"  Расчётный вес (gamma_f={result.gamma_f}): {result.weight_calc_kN:.4f} кН")
        print(f"  Погонная нагрузка: {result.linear_load_kN_m:.4f} кН/м")
    
    print("\n" + "=" * 70)
    print("ДАННЫЕ ИЗ IFC:")
    print("=" * 70)
    print("""
Данные, которые ИЗВЛЕКАЮТСЯ из IFC:
[+] Объём элемента (NetVolume из Qto_*BaseQuantities)
[+] Длина элемента (Length из Qto или section_approx)
[+] Площадь сечения (CrossSectionArea из Qto или bxh из section_approx)
[+] Название материала (из materials)

Данные, которые НЕ извлекаются из IFC (используются справочные):
[-] Плотность материала -> определяется по названию из справочника
[-] Коэффициент надёжности gamma_f -> по умолчанию 1.1 (неблагоприятное действие)

Справочник плотностей материалов:
  - Сталь: 7850 кг/м3
  - Бетон: 2500 кг/м3
  - Древесина (ель): 450 кг/м3
  - Древесина (сосна): 500 кг/м3
  - Кирпич: 1800 кг/м3
""")
