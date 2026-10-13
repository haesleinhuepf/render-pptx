"""Stroke polylines with width, dash patterns, caps and joins onto a Pillow RGBA layer."""
import math

from PIL import Image, ImageDraw

DASHES = {
    "solid": None,
    "dot": (1, 3), "sysDot": (1, 1), "dash": (4, 3), "sysDash": (3, 1),
    "lgDash": (8, 3), "dashDot": (4, 3, 1, 3), "sysDashDot": (3, 1, 1, 1),
    "lgDashDot": (8, 3, 1, 3), "lgDashDotDot": (8, 3, 1, 3, 1, 3),
    "sysDashDotDot": (3, 1, 1, 1, 1, 1),
}


def dash_polyline(pts, pattern, width):
    """Split a polyline into the visible pieces of a dash pattern (lengths in multiples of width)."""
    pieces, cur = [], None
    idx, remaining = 0, max(pattern[0] * width, 1e-6)
    on = True
    for p0, p1 in zip(pts, pts[1:]):
        seg = math.hypot(p1[0] - p0[0], p1[1] - p0[1])
        if seg == 0:
            continue
        pos = 0.0
        while pos < seg:
            step = min(remaining, seg - pos)
            a = (p0[0] + (p1[0] - p0[0]) * pos / seg, p0[1] + (p1[1] - p0[1]) * pos / seg)
            b = (p0[0] + (p1[0] - p0[0]) * (pos + step) / seg, p0[1] + (p1[1] - p0[1]) * (pos + step) / seg)
            if on:
                if cur is None:
                    cur = [a]
                    pieces.append(cur)
                cur.append(b)
            pos += step
            remaining -= step
            if remaining <= 1e-9:
                idx += 1
                remaining = max(pattern[idx % len(pattern)] * width, 1e-6)
                on = idx % 2 == 0
                if not on:
                    cur = None
    return pieces


def _unit(a, b):
    dx, dy = b[0] - a[0], b[1] - a[1]
    n = math.hypot(dx, dy)
    return (dx / n, dy / n) if n else None


def _stroke_piece(draw, pts, closed, color, width, cap, join, miter_limit):
    hw = width / 2.0
    clean = [pts[0]]
    for p in pts[1:]:
        if math.hypot(p[0] - clean[-1][0], p[1] - clean[-1][1]) > 1e-6:
            clean.append(p)
    if closed and len(clean) > 1 and math.hypot(clean[0][0] - clean[-1][0], clean[0][1] - clean[-1][1]) <= 1e-6:
        clean.pop()
    pts = clean
    if len(pts) == 1:
        if cap == "rnd":
            draw.ellipse((pts[0][0] - hw, pts[0][1] - hw, pts[0][0] + hw, pts[0][1] + hw), fill=color)
        return
    n = len(pts)
    segs = n if closed else n - 1
    dirs = [_unit(pts[i], pts[(i + 1) % n]) for i in range(segs)]
    for i in range(segs):
        a, b, u = pts[i], pts[(i + 1) % n], dirs[i]
        a2, b2 = a, b
        if not closed and cap == "sq":
            if i == 0:
                a2 = (a[0] - u[0] * hw, a[1] - u[1] * hw)
            if i == segs - 1:
                b2 = (b[0] + u[0] * hw, b[1] + u[1] * hw)
        nx, ny = -u[1] * hw, u[0] * hw
        draw.polygon([(a2[0] + nx, a2[1] + ny), (b2[0] + nx, b2[1] + ny),
                      (b2[0] - nx, b2[1] - ny), (a2[0] - nx, a2[1] - ny)], fill=color)
    joins = range(n) if closed else range(1, n - 1)
    for i in joins:
        u1, u2 = dirs[(i - 1) % segs], dirs[i % segs]
        v = pts[i]
        if join == "round":
            if width >= 3:
                draw.ellipse((v[0] - hw, v[1] - hw, v[0] + hw, v[1] + hw), fill=color)
            continue
        cross = u1[0] * u2[1] - u1[1] * u2[0]
        if abs(cross) < 1e-9:
            continue
        side = -1.0 if cross > 0 else 1.0  # outer side of the turn
        p1 = (v[0] + side * -u1[1] * hw, v[1] + side * u1[0] * hw)
        p2 = (v[0] + side * -u2[1] * hw, v[1] + side * u2[0] * hw)
        poly = [v, p1, p2]
        if join == "miter":
            dot = u1[0] * u2[0] + u1[1] * u2[1]
            half = math.sqrt(max((1 + dot) / 2.0, 1e-9))  # cos of half the turning angle
            ratio = 1.0 / half
            if ratio <= miter_limit:
                mx, my = (p1[0] + p2[0]) / 2 - v[0], (p1[1] + p2[1]) / 2 - v[1]
                ml = math.hypot(mx, my)
                if ml:
                    k = hw * ratio / ml
                    poly = [v, p1, (v[0] + mx * k, v[1] + my * k), p2]
        draw.polygon(poly, fill=color)
    if not closed and cap == "rnd":
        for v in (pts[0], pts[-1]):
            draw.ellipse((v[0] - hw, v[1] - hw, v[0] + hw, v[1] + hw), fill=color)


def stroke(layer, pts, closed, color, width, dash=None, cap="flat", join="round", miter_limit=8.0):
    """Draw a polyline on an RGBA layer; width in layer pixels, dash as multiples of width."""
    if len(pts) < 2 or width <= 0:
        return
    # draw on an opaque mask so overlapping parts of a translucent stroke don't double up
    mask = Image.new("L", layer.size, 0)
    draw = ImageDraw.Draw(mask)
    if dash:
        for piece in dash_polyline(pts + [pts[0]] if closed else pts, dash, width):
            _stroke_piece(draw, piece, False, 255, width, cap, join, miter_limit)
    else:
        _stroke_piece(draw, pts, closed, 255, width, cap, join, miter_limit)
    alpha = color[3] if len(color) > 3 else 255
    if alpha < 255:
        mask = mask.point(lambda v: v * alpha // 255)
    stroke_img = Image.new("RGBA", layer.size, tuple(color[:3]) + (0,))
    stroke_img.putalpha(mask)
    layer.alpha_composite(stroke_img)
