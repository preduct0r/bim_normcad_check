#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# python ifc_extract.py --url "https://raw.githubusercontent.com/buildingSMART/Sample-Test-Files/main/IFC%204.0.2.1%20%28IFC%204%29/PCERT-Sample-Scene/Building-Structural.ifc" --types IfcBeam,IfcColumn,IfcSlab --geom --verbose

"""
Universal IFC extractor for downstream mapping (e.g., to NormCAD inputs).
- Works with local IFC path or HTTP(S) URL
- Extracts: ids, class, name/tag, placement, materials, psets/qto
- Optional geometry bbox via ifcopenshell.geom (needs OCC in ifcopenshell build)
Output: JSONL (one element per line) or pretty print.

Usage examples:
  python ifc_extract.py --url RAW_URL --types IfcBeam,IfcColumn --geom --limit 50
  python ifc_extract.py model.ifc --json out.jsonl --pick "Column" --verbose
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import tempfile
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests
import numpy as np

import ifcopenshell
from ifcopenshell.util.element import get_psets, get_material, get_type
from ifcopenshell.util.placement import get_local_placement


LOG = logging.getLogger("ifc_extract")

PROFILE_PARAM_KEYS = (
    # common parameterized profiles
    "XDim",
    "YDim",
    "Radius",
    "SemiAxis1",
    "SemiAxis2",
    # steel shapes, etc.
    "OverallWidth",
    "OverallDepth",
    "WebThickness",
    "FlangeThickness",
    "FilletRadius",
    "EdgeRadius",
)


# --------------------------- IO ---------------------------

def is_url(s: str) -> bool:
    return s.startswith("http://") or s.startswith("https://")


def download_to_temp(url: str) -> str:
    LOG.info("Downloading IFC from URL: %s", url)
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    fd, path = tempfile.mkstemp(suffix=".ifc", prefix="ifc_")
    with os.fdopen(fd, "wb") as f:
        f.write(r.content)
    LOG.info("Saved to: %s (%d bytes)", path, len(r.content))
    return path


def load_ifc(source: str) -> Tuple[Any, str]:
    """Returns (ifc_model, local_path_used)."""
    path = download_to_temp(source) if is_url(source) else source
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    LOG.info("Opening IFC: %s", path)
    model = ifcopenshell.open(path)
    return model, path


# ------------------------- Extractors -------------------------

def safe_str(x: Any) -> Optional[str]:
    return None if x is None else str(x)


def safe_getattr(obj: Any, name: str, default: Any = None) -> Any:
    """
    ifcopenshell entity_instance.__getattr__ can raise RuntimeError for some optional
    attributes in certain exports. Treat such cases as "missing" instead of crashing.
    """
    try:
        return getattr(obj, name)
    except Exception:
        return default


def extract_model_units(model) -> Dict[str, Any]:
    """
    Best-effort extraction of model length units.
    Returns a dict with scale factors useful for downstream conversion.
    """
    out: Dict[str, Any] = {
        "length_unit": None,
        "length_prefix": None,
        "length_scale_to_m": 1.0,
        "length_scale_to_mm": 1000.0,
    }
    try:
        projs = model.by_type("IfcProject") or []
        if not projs:
            return out
        units = getattr(projs[0], "UnitsInContext", None)
        unit_assign = getattr(units, "Units", None) if units else None
        if not unit_assign:
            return out
        for u in unit_assign:
            if not u or not hasattr(u, "is_a"):
                continue
            if u.is_a("IfcSIUnit") and getattr(u, "UnitType", None) == "LENGTHUNIT":
                name = getattr(u, "Name", None)
                prefix = getattr(u, "Prefix", None)  # e.g. MILLI
                out["length_unit"] = safe_str(name)
                out["length_prefix"] = safe_str(prefix)
                # IFC SI unit is always METRE with optional prefix for LENGTHUNIT
                scale_to_m = 1.0
                if prefix:
                    p = str(prefix).upper()
                    scale_to_m = {
                        "MILLI": 1e-3,
                        "CENTI": 1e-2,
                        "DECI": 1e-1,
                        "KILO": 1e3,
                        "MICRO": 1e-6,
                    }.get(p, 1.0)
                out["length_scale_to_m"] = float(scale_to_m)
                out["length_scale_to_mm"] = float(scale_to_m * 1000.0)
                break
    except Exception as e:
        LOG.debug("Units extract failed: %s", e)
    return out


def _coords_bbox_2d(coords: List[List[float]]) -> Optional[Dict[str, Any]]:
    if not coords:
        return None
    xs = [c[0] for c in coords if len(c) >= 2]
    ys = [c[1] for c in coords if len(c) >= 2]
    if not xs or not ys:
        return None
    xmin, xmax = min(xs), max(xs)
    ymin, ymax = min(ys), max(ys)
    return {"min": [xmin, ymin], "max": [xmax, ymax], "size": [xmax - xmin, ymax - ymin]}


def _curve_coords_2d(curve) -> Optional[List[List[float]]]:
    """
    Extract 2D coordinates from the most common curve types used in profile defs.
    We only need bounding box for downstream mapping (e.g., approximate b/h).
    """
    try:
        if curve is None:
            return None
        t = curve.is_a()
        if t == "IfcPolyline":
            coords = []
            for p in getattr(curve, "Points", None) or []:
                c = getattr(p, "Coordinates", None) or []
                if len(c) >= 2:
                    coords.append([float(c[0]), float(c[1])])
            return coords or None
        if t == "IfcIndexedPolyCurve":
            pts = getattr(curve, "Points", None)
            # Points is typically IfcCartesianPointList2D with CoordList
            cl = getattr(pts, "CoordList", None) if pts else None
            if cl:
                coords = [[float(x), float(y)] for (x, y, *_) in cl]
                return coords or None
        return None
    except Exception:
        return None


def profile_record(profile, units: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if profile is None or not hasattr(profile, "is_a"):
        return None
    out: Dict[str, Any] = {
        "ifc_class": profile.is_a(),
        "profile_name": safe_str(getattr(profile, "ProfileName", None)),
        "profile_type": safe_str(getattr(profile, "ProfileType", None)),
    }
    # parameters (if parameterized)
    for k in PROFILE_PARAM_KEYS:
        if hasattr(profile, k):
            try:
                out[k] = float(getattr(profile, k))
            except Exception:
                out[k] = safe_str(getattr(profile, k))

    # bbox for arbitrary profiles (or anything where we can read points)
    outer = getattr(profile, "OuterCurve", None)
    if outer is not None and hasattr(outer, "is_a"):
        out["outer_curve_class"] = outer.is_a()
        coords = _curve_coords_2d(outer)
        bbox2d = _coords_bbox_2d(coords) if coords else None
        if bbox2d:
            out["bbox2d"] = bbox2d
            # Convenience: approximate width/height in model units and in mm
            sx, sy = bbox2d["size"]
            out["approx_width"] = float(sx)
            out["approx_height"] = float(sy)
            scale_to_mm = float(units.get("length_scale_to_mm", 1000.0))
            out["approx_width_mm"] = float(sx * scale_to_mm)
            out["approx_height_mm"] = float(sy * scale_to_mm)
    return out


def extract_section_from_representation(el, units: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Try to get cross-section profile + extrusion depth from geometric representation.
    Works well for IfcExtrudedAreaSolid-based elements.
    """
    try:
        rep = getattr(el, "Representation", None)
        if not rep:
            return None
        reps = getattr(rep, "Representations", None) or []
        for r in reps:
            items = getattr(r, "Items", None) or []
            for it in items:
                if not it or not hasattr(it, "is_a"):
                    continue
                t = it.is_a()
                if t == "IfcExtrudedAreaSolid":
                    prof = profile_record(getattr(it, "SweptArea", None), units)
                    depth = getattr(it, "Depth", None)
                    sec: Dict[str, Any] = {"solid_class": t, "profile": prof}
                    if depth is not None:
                        sec["extrusion_depth"] = float(depth)
                        sec["extrusion_depth_mm"] = float(depth * float(units.get("length_scale_to_mm", 1000.0)))
                    return sec
                if t == "IfcSweptAreaSolid":
                    prof = profile_record(getattr(it, "SweptArea", None), units)
                    return {"solid_class": t, "profile": prof}
        return None
    except Exception as e:
        LOG.debug("Section-from-repr failed for %s: %s", getattr(el, "GlobalId", "?"), e)
        return None


def extract_material_profile(el, units: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Try to extract profile from material associations (IfcMaterialProfileSetUsage etc.).
    Some IFC exports store section/profile here rather than in geometry.
    """
    try:
        mat = get_material(el, should_inherit=True)
        if mat is None or not hasattr(mat, "is_a"):
            return None
        t = mat.is_a()
        if t == "IfcMaterialProfileSetUsage":
            mps = getattr(mat, "ForProfileSet", None)
            t = mps.is_a() if mps else t
            profiles = getattr(mps, "MaterialProfiles", None) if mps else None
            if profiles:
                mp = profiles[0]
                prof = profile_record(getattr(mp, "Profile", None), units)
                return {"material_class": "IfcMaterialProfileSetUsage", "profile": prof}
        if t == "IfcMaterialProfileSet":
            profiles = getattr(mat, "MaterialProfiles", None) or []
            if profiles:
                mp = profiles[0]
                prof = profile_record(getattr(mp, "Profile", None), units)
                return {"material_class": "IfcMaterialProfileSet", "profile": prof}
        return None
    except Exception as e:
        LOG.debug("Material profile extract failed for %s: %s", getattr(el, "GlobalId", "?"), e)
        return None


def extract_placement_xyz(el) -> Optional[Dict[str, Any]]:
    """Get element placement as matrix + xyz (world)."""
    try:
        if not safe_getattr(el, "ObjectPlacement", None):
            return None
        m = get_local_placement(el.ObjectPlacement)  # 4x4 matrix (nested list or numpy array)
        # Convert numpy arrays to lists for JSON serialization
        if isinstance(m, np.ndarray):
            m = m.tolist()
        elif isinstance(m, (list, tuple)) and len(m) > 0:
            # Handle nested numpy arrays
            m = [[float(x) if isinstance(x, (np.integer, np.floating)) else float(x) for x in row] if isinstance(row, (list, tuple, np.ndarray)) else row for row in m]
        return {
            "matrix_4x4": m,
            "xyz": [float(m[0][3]), float(m[1][3]), float(m[2][3])]
        }
    except Exception as e:
        LOG.debug("Placement extract failed for %s: %s", getattr(el, "GlobalId", "?"), e)
        return None


def extract_materials(el) -> List[str]:
    """Return flattened material names if present."""
    mats: List[str] = []
    try:
        mat = get_material(el, should_inherit=True)
        if mat is None:
            return mats

        # get_material may return IfcMaterial, IfcMaterialLayerSet, IfcMaterialProfileSet, etc.
        t = mat.is_a()
        if t == "IfcMaterial":
            if safe_getattr(mat, "Name", None):
                mats.append(str(mat.Name))
        else:
            # Walk common containers
            for attr in ("MaterialLayers", "MaterialProfiles", "Materials"):
                if hasattr(mat, attr):
                    items = getattr(mat, attr) or []
                    for it in items:
                        # layers: it.Material, profiles: it.Material
                        m = getattr(it, "Material", None) or it
                        name = safe_getattr(m, "Name", None)
                        if name:
                            mats.append(str(name))
        # unique while keeping order
        seen = set()
        out = []
        for x in mats:
            if x not in seen:
                seen.add(x)
                out.append(x)
        return out
    except Exception as e:
        LOG.debug("Material extract failed for %s: %s", getattr(el, "GlobalId", "?"), e)
        return []


def extract_psets(el, include_qto: bool = True) -> Dict[str, Any]:
    """
    Property Sets & Quantities as nested dict.
    include_qto=True usually returns both Psets and Qto (Quantities) if present.
    """
    try:
        return get_psets(el, psets_only=not include_qto) or {}
    except Exception as e:
        LOG.debug("Pset extract failed for %s: %s", getattr(el, "GlobalId", "?"), e)
        return {}


def compute_bbox_from_shape(shape) -> Optional[Dict[str, Any]]:
    """Compute bbox from mesh vertices (no extra util funcs)."""
    try:
        verts = shape.geometry.verts  # flat list [x0,y0,z0,x1,y1,z1,...]
        if not verts:
            return None
        xs = verts[0::3]
        ys = verts[1::3]
        zs = verts[2::3]
        xmin, xmax = min(xs), max(xs)
        ymin, ymax = min(ys), max(ys)
        zmin, zmax = min(zs), max(zs)
        dx, dy, dz = (xmax - xmin), (ymax - ymin), (zmax - zmin)
        return {
            "min": [xmin, ymin, zmin],
            "max": [xmax, ymax, zmax],
            "size": [dx, dy, dz],
        }
    except Exception as e:
        LOG.debug("BBox compute failed: %s", e)
        return None


def extract_geometry_bbox(model, el) -> Optional[Dict[str, Any]]:
    """
    Optional geometry extraction using ifcopenshell.geom.
    Note: requires ifcopenshell build with geometry engine (OCC).
    """
    try:
        import ifcopenshell.geom
        settings = ifcopenshell.geom.settings()
        settings.set(settings.USE_WORLD_COORDS, True)
        shape = ifcopenshell.geom.create_shape(settings, el)
        bbox = compute_bbox_from_shape(shape)
        return bbox
    except Exception as e:
        LOG.debug("Geometry extract failed for %s: %s", getattr(el, "GlobalId", "?"), e)
        return None


def element_record(model, el, geom: bool, include_qto: bool) -> Dict[str, Any]:
    units = extract_model_units(model)
    typ = None
    try:
        typ = get_type(el)
    except Exception:
        typ = None
    rec: Dict[str, Any] = {
        "global_id": safe_str(safe_getattr(el, "GlobalId", None)),
        "ifc_class": el.is_a(),
        "name": safe_str(safe_getattr(el, "Name", None)),
        "tag": safe_str(safe_getattr(el, "Tag", None)),
        "predefined_type": safe_str(safe_getattr(el, "PredefinedType", None)),
        "type_name": safe_str(getattr(typ, "Name", None)) if typ else None,
        "type_global_id": safe_str(getattr(typ, "GlobalId", None)) if typ else None,
        "placement": extract_placement_xyz(el),
        "materials": extract_materials(el),
        "psets": extract_psets(el, include_qto=include_qto),
        "model_units": units,
    }

    # section / profile: prefer material profile set, else representation
    rec["section"] = extract_material_profile(el, units) or extract_section_from_representation(el, units)

    if geom:
        rec["bbox"] = extract_geometry_bbox(model, el)

        # handy heuristic length for linear members (beam/column/member): max bbox dimension
        bbox = rec.get("bbox") or {}
        size = bbox.get("size")
        if size:
            rec["heuristic_length"] = float(max(size))

    return rec


def iter_elements(model, classes: Optional[List[str]]) -> Iterable[Any]:
    if not classes:
        # reasonable default set for structural extraction
        classes = ["IfcBeam", "IfcColumn", "IfcMember", "IfcSlab", "IfcWall", "IfcFooting"]
    for c in classes:
        for el in model.by_type(c):
            yield el


# --------------------------- Output ---------------------------

def json_serialize(obj: Any) -> Any:
    """Convert numpy types and other non-serializable objects to JSON-compatible types."""
    if isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, (np.integer, np.intc, np.intp, np.int8, np.int16, np.int32, np.int64)):
        return int(obj)
    elif isinstance(obj, (np.floating, np.float16, np.float32, np.float64)):
        return float(obj)
    elif isinstance(obj, np.bool_):
        return bool(obj)
    elif isinstance(obj, dict):
        return {key: json_serialize(value) for key, value in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [json_serialize(item) for item in obj]
    return obj


def pretty_print(rec: Dict[str, Any]) -> None:
    gid = rec.get("global_id")
    cls = rec.get("ifc_class")
    name = rec.get("name")
    print(f"- {cls}  {gid}  name={name!r}")
    if rec.get("materials"):
        print(f"  materials: {rec['materials']}")
    if rec.get("placement", {}).get("xyz"):
        print(f"  xyz: {rec['placement']['xyz']}")
    if rec.get("bbox"):
        print(f"  bbox.size: {rec['bbox'].get('size')}")
    if rec.get("heuristic_length") is not None:
        print(f"  heuristic_length: {rec['heuristic_length']}")
    # show only top-level pset names for readability
    psets = rec.get("psets") or {}
    if psets:
        keys = sorted(psets.keys())
        print(f"  psets/qto: {keys[:12]}{' ...' if len(keys) > 12 else ''}")


# --------------------------- Main ---------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description="Universal IFC extractor (debug-friendly)")
    ap.add_argument("source", nargs="?", help="Path to .ifc or URL")
    ap.add_argument("--url", help="IFC URL (alternative to positional source)")
    ap.add_argument("--types", default="", help="Comma-separated IFC classes (e.g., IfcBeam,IfcColumn). Default: structural set.")
    ap.add_argument("--pick", default="", help="Filter by substring in name/tag/global_id (case-insensitive)")
    ap.add_argument("--limit", type=int, default=0, help="Max number of elements to output (0 = no limit)")
    ap.add_argument("--geom", action="store_true", help="Try to compute bbox from geometry (needs geom engine)")
    ap.add_argument("--no-qto", action="store_true", help="Do not include Qto (quantities), only Psets")
    ap.add_argument("--json", default="", help="Write JSONL to file (one element per line)")
    ap.add_argument("--verbose", action="store_true", help="Verbose logging")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s:%(name)s:%(message)s"
    )

    source = args.url or args.source
    if not source:
        ap.error("Provide IFC source path or --url")

    classes = [x.strip() for x in args.types.split(",") if x.strip()] or None
    pick = args.pick.lower().strip()
    include_qto = not args.no_qto

    model, used_path = load_ifc(source)
    LOG.info("Schema: %s", model.schema)

    out_f = open(args.json, "w", encoding="utf-8") if args.json else None
    count = 0

    try:
        for el in iter_elements(model, classes):
            rec = element_record(model, el, geom=args.geom, include_qto=include_qto)

            if pick:
                hay = " ".join([
                    (rec.get("global_id") or ""),
                    (rec.get("name") or ""),
                    (rec.get("tag") or ""),
                    (rec.get("ifc_class") or ""),
                ]).lower()
                if pick not in hay:
                    continue

            if out_f:
                # Serialize numpy arrays and other non-JSON types
                rec_serialized = json_serialize(rec)
                out_f.write(json.dumps(rec_serialized, ensure_ascii=False) + "\n")
            else:
                pretty_print(rec)

            count += 1
            if args.limit and count >= args.limit:
                break

        LOG.info("Done. Extracted: %d elements", count)
        LOG.info("IFC file used: %s", used_path)
        return 0
    finally:
        if out_f:
            out_f.close()


if __name__ == "__main__":
    raise SystemExit(main())
