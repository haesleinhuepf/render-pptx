"""Evaluate DrawingML preset (prstGeom) and custom (custGeom) shape geometry into polylines."""
import math
import re
import xml.etree.ElementTree as ET
from functools import lru_cache
from pathlib import Path

_DEFS_FILE = Path(__file__).parent / "data" / "presetShapeDefinitions.xml"
_ANGLE = 60000.0
_ANGLE_CONSTANTS = {"cd2": 10800000, "cd4": 5400000, "cd8": 2700000, "3cd4": 16200000,
                    "3cd8": 8100000, "5cd8": 13500000, "7cd8": 18900000}
_DIVIDED = re.compile(r"^(w|h|ss|ls)d(\d+)$")


def _local(el):
    return el.tag.rpartition("}")[2]


def _children(el, name):
    return [c for c in el if _local(c) == name]


def _child(el, name):
    for c in el:
        if _local(c) == name:
            return c
    return None


@lru_cache(maxsize=1)
def _preset_definitions():
    root = ET.parse(_DEFS_FILE).getroot()
    return {el.tag: el for el in root}


def preset_names():
    return set(_preset_definitions())


class _Guides:
    """Lazily evaluated guide/adjust-value namespace for one shape."""

    def __init__(self, geom, w, h, overrides):
        self.cache = {}
        self.formulas = {}
        self.cache.update({"w": w, "h": h, "l": 0.0, "t": 0.0, "r": w, "b": h, "hc": w / 2.0,
                           "vc": h / 2.0, "ss": min(w, h), "ls": max(w, h)})
        self.cache.update({k: float(v) for k, v in _ANGLE_CONSTANTS.items()})
        for lst in ("avLst", "gdLst"):
            node = _child(geom, lst)
            if node is None:
                continue
            for gd in _children(node, "gd"):
                self.formulas[gd.get("name")] = gd.get("fmla", "")
        for name, val in overrides.items():
            if name in self.formulas:
                self.formulas[name] = "val %s" % val
                self.cache.pop(name, None)

    def get(self, token):
        try:
            return float(token)
        except ValueError:
            pass
        if token in self.cache:
            return self.cache[token]
        m = _DIVIDED.match(token)
        if m:
            base = self.get(m.group(1))
            self.cache[token] = base / float(m.group(2))
            return self.cache[token]
        fmla = self.formulas.get(token)
        if fmla is None:
            return 0.0
        self.cache[token] = 0.0  # guards against cyclic definitions
        self.cache[token] = self._compute(fmla)
        return self.cache[token]

    def _compute(self, fmla):
        parts = fmla.split()
        op, args = parts[0], parts[1:]
        v = [self.get(a) for a in args] + [0.0, 0.0, 0.0]
        x, y, z = v[0], v[1], v[2]
        rad = lambda a: math.radians(a / _ANGLE)
        if op == "val":
            return x
        if op == "*/":
            return x * y / z if z else 0.0
        if op == "+-":
            return x + y - z
        if op == "+/":
            return (x + y) / z if z else 0.0
        if op == "?:":
            return y if x > 0 else z
        if op == "abs":
            return abs(x)
        if op == "sqrt":
            return math.sqrt(x) if x > 0 else 0.0
        if op == "max":
            return max(x, y)
        if op == "min":
            return min(x, y)
        if op == "mod":
            return math.sqrt(x * x + y * y + z * z)
        if op == "pin":
            return x if y < x else z if y > z else y
        if op == "sin":
            return x * math.sin(rad(y))
        if op == "cos":
            return x * math.cos(rad(y))
        if op == "tan":
            return x * math.tan(rad(y))
        if op == "at2":
            return math.degrees(math.atan2(y, x)) * _ANGLE
        if op == "cat2":
            return x * math.cos(math.atan2(z, y))
        if op == "sat2":
            return x * math.sin(math.atan2(z, y))
        return 0.0


def _arc_points(cx, cy, wr, hr, t0, dt):
    n = max(6, int(abs(dt) / (math.pi / 36)) + 1)
    return [(cx + wr * math.cos(t0 + dt * i / n), cy + hr * math.sin(t0 + dt * i / n)) for i in range(1, n + 1)]


def _bezier(pts, n=20):
    out = []
    for i in range(1, n + 1):
        t = i / n
        u = 1 - t
        if len(pts) == 4:
            c = (u ** 3, 3 * u * u * t, 3 * u * t * t, t ** 3)
        else:
            c = (u * u, 2 * u * t, t * t)
        out.append((sum(k * p[0] for k, p in zip(c, pts)), sum(k * p[1] for k, p in zip(c, pts))))
    return out


def _build_path(path, guides, w, h):
    pw, ph = float(path.get("w") or 0), float(path.get("h") or 0)
    sx = w / pw if pw else 1.0
    sy = h / ph if ph else 1.0
    pt = lambda el: (guides.get(el.get("x")) * sx, guides.get(el.get("y")) * sy)
    subpaths = []
    pts = None
    cur = (0.0, 0.0)
    start = cur
    for op in path:
        name = _local(op)
        if name == "moveTo":
            if pts and len(pts) > 1:
                subpaths.append((pts, False))
            cur = start = pt(_child(op, "pt"))
            pts = [cur]
            continue
        if name == "close":
            if pts and len(pts) > 1:
                subpaths.append((pts, True))
            pts, cur = None, start
            continue
        if pts is None:
            pts = [cur]
        if name == "lnTo":
            cur = pt(_child(op, "pt"))
            pts.append(cur)
        elif name in ("cubicBezTo", "quadBezTo"):
            ctrl = [cur] + [pt(p) for p in _children(op, "pt")]
            new = _bezier(ctrl)
            pts.extend(new)
            cur = new[-1]
        elif name == "arcTo":
            wr = guides.get(op.get("wR")) * sx
            hr = guides.get(op.get("hR")) * sy
            st = math.radians(guides.get(op.get("stAng")) / _ANGLE)
            sw = math.radians(guides.get(op.get("swAng")) / _ANGLE)
            if wr <= 0 or hr <= 0:
                continue
            # visual angle -> ellipse parameter
            t0 = math.atan2(wr * math.sin(st), hr * math.cos(st))
            if abs(sw) >= 2 * math.pi - 1e-6:
                dt = sw
            else:
                t1 = math.atan2(wr * math.sin(st + sw), hr * math.cos(st + sw))
                dt = t1 - t0
                while sw > 0 and dt < 0:
                    dt += 2 * math.pi
                while sw < 0 and dt > 0:
                    dt -= 2 * math.pi
            cx, cy = cur[0] - wr * math.cos(t0), cur[1] - hr * math.sin(t0)
            new = _arc_points(cx, cy, wr, hr, t0, dt)
            pts.extend(new)
            cur = new[-1]
    if pts and len(pts) > 1:
        subpaths.append((pts, False))
    return subpaths


def evaluate(geom, w, h, overrides=None):
    """Return [{"fill": mode, "stroke": bool, "subpaths": [(points, closed), ...]}] for a geometry element.

    geom is a preset definition (see ``preset_geometry``) or an a:custGeom element.
    """
    guides = _Guides(geom, w, h, overrides or {})
    path_lst = _child(geom, "pathLst")
    result = []
    for path_index, path in enumerate(_children(path_lst, "path") if path_lst is not None else []):
        stroke = path.get("stroke", "true") not in ("0", "false")
        if _local(geom) == "cloud" and path_index == 1:
            # This auxiliary cloud path is made of disconnected accent arcs.
            stroke = False
        result.append({
            "fill": path.get("fill", "norm"),
            "stroke": stroke,
            "subpaths": _build_path(path, guides, w, h),
        })
    return result


def preset_geometry(name):
    return _preset_definitions().get(name)


def shape_geometry(shape):
    """Return (geometry element, adjust overrides) for a shape, or (None, {})."""
    try:
        sppr = shape._element.spPr
    except Exception:
        return None, {}
    prst = cust = None
    for c in sppr:
        if _local(c) == "prstGeom":
            prst = c
        elif _local(c) == "custGeom":
            cust = c
    if cust is not None:
        return cust, {}
    if prst is None:
        return None, {}
    geom = preset_geometry(prst.get("prst"))
    overrides = {}
    av = _child(prst, "avLst")
    if av is not None:
        for gd in _children(av, "gd"):
            fmla = gd.get("fmla", "")
            if fmla.startswith("val "):
                overrides[gd.get("name")] = fmla[4:].strip()
    return geom, overrides
