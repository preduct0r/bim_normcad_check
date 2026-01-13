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

import ifcopenshell
from ifcopenshell.util.element import get_psets, get_material
from ifcopenshell.util.placement import get_local_placement


LOG = logging.getLogger("ifc_extract")


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


def extract_placement_xyz(el) -> Optional[Dict[str, Any]]:
    """Get element placement as matrix + xyz (world)."""
    try:
        if not getattr(el, "ObjectPlacement", None):
            return None
        m = get_local_placement(el.ObjectPlacement)  # 4x4 matrix (nested list)
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
            if getattr(mat, "Name", None):
                mats.append(str(mat.Name))
        else:
            # Walk common containers
            for attr in ("MaterialLayers", "MaterialProfiles", "Materials"):
                if hasattr(mat, attr):
                    items = getattr(mat, attr) or []
                    for it in items:
                        # layers: it.Material, profiles: it.Material
                        m = getattr(it, "Material", None) or it
                        name = getattr(m, "Name", None)
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
    rec: Dict[str, Any] = {
        "global_id": safe_str(getattr(el, "GlobalId", None)),
        "ifc_class": el.is_a(),
        "name": safe_str(getattr(el, "Name", None)),
        "tag": safe_str(getattr(el, "Tag", None)),
        "predefined_type": safe_str(getattr(el, "PredefinedType", None)),
        "placement": extract_placement_xyz(el),
        "materials": extract_materials(el),
        "psets": extract_psets(el, include_qto=include_qto),
    }

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
                out_f.write(json.dumps(rec, ensure_ascii=False) + "\n")
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
