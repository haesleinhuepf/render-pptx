"""Render pptx slides onto a PIL canvas."""
import io
import os

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.dml import MSO_FILL
from pptx.util import Emu

_FONTS = ["DejaVuSans.ttf", "Arial.ttf", "arial.ttf", "LiberationSans-Regular.ttf"]
_FONTS_BOLD = ["DejaVuSans-Bold.ttf", "Arial Bold.ttf", "arialbd.ttf", "LiberationSans-Bold.ttf"]


def _font(size, bold=False):
    size = max(int(size), 1)
    for name in (_FONTS_BOLD if bold else []) + _FONTS:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default(size)


def _rgb(color_format):
    """Return an RGB tuple for a python-pptx ColorFormat, or None."""
    try:
        if color_format.type is None:
            return None
        return tuple(color_format.rgb)
    except (AttributeError, KeyError, TypeError, ValueError):
        pass
    try:  # theme colors have no direct rgb
        from pptx.enum.dml import MSO_THEME_COLOR as T
        return {
            T.BACKGROUND_1: (255, 255, 255), T.LIGHT_1: (255, 255, 255),
            T.TEXT_1: (0, 0, 0), T.DARK_1: (0, 0, 0),
        }.get(color_format.theme_color)
    except Exception:
        return None


def _fill_color(fill):
    try:
        if fill.type == MSO_FILL.SOLID:
            return _rgb(fill.fore_color)
        if fill.type == MSO_FILL.GRADIENT:
            return _rgb(fill.gradient_stops[0].color)
    except Exception:
        pass
    return None


class _Canvas:
    def __init__(self, width_emu, height_emu, width_px):
        self.scale = width_px / width_emu
        self.size = (width_px, max(1, round(height_emu * self.scale)))
        self.image = Image.new("RGB", self.size, (255, 255, 255))

    def px(self, emu):
        return emu * self.scale

    def box(self, shape, ox=0, oy=0, sx=1.0, sy=1.0):
        """Pixel bounding box (x0, y0, x1, y1) of a shape, with group transform."""
        l, t, w, h = shape.left, shape.top, shape.width, shape.height
        if None in (l, t, w, h):
            return None
        x0 = ox + l * sx
        y0 = oy + t * sy
        return (self.px(x0), self.px(y0), self.px(x0 + w * sx), self.px(y0 + h * sy))

    def paste_rotated(self, layer, box, rotation):
        """Composite an RGBA layer (sized to box) onto the canvas, rotated about its center."""
        if rotation:
            layer = layer.rotate(-rotation, expand=True, resample=Image.BICUBIC)
        cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
        pos = (round(cx - layer.width / 2), round(cy - layer.height / 2))
        self.image.paste(layer, pos, layer)


def _draw_text(canvas, text_frame, box, default_color=(0, 0, 0)):
    x0, y0, x1, y1 = box
    pad_l = canvas.px(text_frame.margin_left)
    pad_r = canvas.px(text_frame.margin_right)
    pad_t = canvas.px(text_frame.margin_top)
    pad_b = canvas.px(text_frame.margin_bottom)
    avail = max(x1 - x0 - pad_l - pad_r, 1)
    draw = ImageDraw.Draw(canvas.image)

    lines = []  # (text, font, color, align, height)
    for para in text_frame.paragraphs:
        runs = [r for r in para.runs if r.text]
        if not runs:
            lines.append(("", _font(canvas.px(Emu(127000))), default_color, None, canvas.px(Emu(127000)) * 1.2))
            continue
        # one font/color per paragraph line chunk; wrap per run words
        words = []
        for r in runs:
            size = r.font.size or (para.font.size if para.font.size else Emu(18 * 12700))
            fpx = canvas.px(size)
            font = _font(fpx, bool(r.font.bold))
            color = default_color
            try:
                color = _rgb(r.font.color) or default_color
            except Exception:
                pass
            for i, wd in enumerate(r.text.replace("\v", "\n").split(" ")):
                words.append((wd, font, color, fpx))
        cur, cur_w = [], 0
        def flush():
            nonlocal cur, cur_w
            if cur:
                h = max(w[3] for w in cur) * 1.2
                lines.append((cur, None, None, para.alignment, h))
            cur, cur_w = [], 0
        for wd, font, color, fpx in words:
            ww = draw.textlength(wd + " ", font=font)
            if cur and cur_w + ww > avail + draw.textlength(" ", font=font):
                flush()
            cur.append((wd, font, color, fpx))
            cur_w += ww
        flush()

    total = sum(l[4] for l in lines)
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    anchor = text_frame.vertical_anchor
    inner_h = y1 - y0 - pad_t - pad_b
    y = y0 + pad_t
    if anchor == MSO_ANCHOR.MIDDLE:
        y += (inner_h - total) / 2
    elif anchor == MSO_ANCHOR.BOTTOM:
        y += inner_h - total
    for content, _f, _c, align, h in lines:
        if isinstance(content, str):
            y += h
            continue
        line_w = sum(draw.textlength(w[0] + " ", font=w[1]) for w in content) - draw.textlength(" ", font=content[-1][1])
        x = x0 + pad_l
        if align == PP_ALIGN.CENTER:
            x += (avail - line_w) / 2
        elif align == PP_ALIGN.RIGHT:
            x += avail - line_w
        for wd, font, color, fpx in content:
            draw.text((x, y), wd, font=font, fill=color)
            x += draw.textlength(wd + " ", font=font)
        y += h


def _picture_image(shape):
    """Return the picture cropped according to its crop information."""
    img = Image.open(io.BytesIO(shape.image.blob))
    img.load()
    img = img.convert("RGBA")
    w, h = img.size
    l, t, r, b = shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom
    # negative crop values pad the image with transparency
    x0, y0, x1, y1 = round(l * w), round(t * h), round(w - r * w), round(h - b * h)
    if x1 <= x0 or y1 <= y0:
        return None
    if min(x0, y0) >= 0 and x1 <= w and y1 <= h:
        return img.crop((x0, y0, x1, y1))
    out = Image.new("RGBA", (x1 - x0, y1 - y0), (0, 0, 0, 0))
    out.paste(img, (-x0, -y0))
    return out


def _draw_picture(canvas, shape, box):
    img = _picture_image(shape)
    if img is None:
        return
    bw, bh = max(round(box[2] - box[0]), 1), max(round(box[3] - box[1]), 1)
    img = img.resize((bw, bh), Image.LANCZOS)
    canvas.paste_rotated(img, box, getattr(shape, "rotation", 0))


def _draw_shape_geometry(canvas, shape, box, fill, line, line_w):
    from pptx.enum.shapes import MSO_SHAPE
    layer = Image.new("RGBA", (max(round(box[2] - box[0]), 1), max(round(box[3] - box[1]), 1)), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    rect = (0, 0, layer.width - 1, layer.height - 1)
    kind = None
    try:
        kind = shape.auto_shape_type
    except Exception:
        pass
    kw = dict(fill=fill, outline=line, width=line_w)
    if kind == MSO_SHAPE.OVAL:
        d.ellipse(rect, **kw)
    elif kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        d.rounded_rectangle(rect, radius=min(layer.size) // 6, **kw)
    elif kind in (MSO_SHAPE.ISOSCELES_TRIANGLE,):
        d.polygon([(rect[2] / 2, 0), (rect[2], rect[3]), (0, rect[3])], **kw)
    else:
        d.rectangle(rect, **kw)
    canvas.paste_rotated(layer, box, getattr(shape, "rotation", 0))


def _draw_table(canvas, shape, box):
    table = shape.table
    draw = ImageDraw.Draw(canvas.image)
    y = box[1]
    for ri, row in enumerate(table.rows):
        rh = canvas.px(row.height)
        x = box[0]
        for ci, col in enumerate(table.columns):
            cw = canvas.px(col.width)
            cell = table.cell(ri, ci)
            if cell.is_spanned:
                x += cw
                continue
            c = _fill_color(cell.fill)
            draw.rectangle((x, y, x + cw, y + rh), fill=c, outline=(128, 128, 128))
            _draw_text(canvas, cell.text_frame, (x, y, x + cw, y + rh))
            x += cw
        y += rh


def _draw_shapes(canvas, shapes, ox=0, oy=0, sx=1.0, sy=1.0, skip_placeholders=False):
    for shape in shapes:
        if skip_placeholders and shape.is_placeholder:
            continue
        try:
            _draw_shape(canvas, shape, ox, oy, sx, sy)
        except Exception:  # never fail a whole slide for one odd shape
            continue


def _draw_shape(canvas, shape, ox, oy, sx, sy):
    st = shape.shape_type
    if st == MSO_SHAPE_TYPE.GROUP:
        # map child coordinate space onto the group's extents
        xfrm = shape._element.grpSpPr.xfrm
        ch_off, ch_ext = xfrm.chOff, xfrm.chExt
        gx = ox + shape.left * sx
        gy = oy + shape.top * sy
        nsx = shape.width * sx / ch_ext.cx if ch_ext.cx else sx
        nsy = shape.height * sy / ch_ext.cy if ch_ext.cy else sy
        _draw_shapes(canvas, shape.shapes, gx - ch_off.x * nsx, gy - ch_off.y * nsy, nsx, nsy)
        return
    box = canvas.box(shape, ox, oy, sx, sy)
    if box is None:
        return
    if st == MSO_SHAPE_TYPE.PICTURE or (shape.is_placeholder and hasattr(shape, "image")):
        _draw_picture(canvas, shape, box)
        return
    if getattr(shape, "has_table", False) and shape.has_table:
        _draw_table(canvas, shape, box)
        return
    if st == MSO_SHAPE_TYPE.LINE or shape.__class__.__name__ == "Connector":
        color = _rgb(shape.line.color) or (0, 0, 0)
        width = max(round(canvas.px(shape.line.width or Emu(12700))), 1)
        ImageDraw.Draw(canvas.image).line((box[0], box[1], box[2], box[3]), fill=color, width=width)
        return
    fill = line = None
    line_w = 0
    try:
        fill = _fill_color(shape.fill)
        if shape.line.fill.type == MSO_FILL.SOLID:
            line = _rgb(shape.line.color)
            line_w = max(round(canvas.px(shape.line.width or Emu(12700))), 1)
    except Exception:
        pass
    if fill or line:
        _draw_shape_geometry(canvas, shape, box, fill, line, line_w)
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        _draw_text(canvas, shape.text_frame, box)


def _background(canvas, slide):
    for obj in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        try:
            fill = obj.background.fill
            c = _fill_color(fill)
            if c:
                canvas.image.paste(c, (0, 0, *canvas.size))
                return
        except Exception:
            continue


def _render(prs, slide, width):
    canvas = _Canvas(prs.slide_width, prs.slide_height, width)
    _background(canvas, slide)
    layout = slide.slide_layout
    # master elements first, then layout, then the slide itself
    _draw_shapes(canvas, layout.slide_master.shapes, skip_placeholders=True)
    _draw_shapes(canvas, layout.shapes, skip_placeholders=True)
    _draw_shapes(canvas, slide.shapes)
    return canvas.image


def render_slide(pptx, index=0, width=1920):
    """Render slide ``index`` (0-based) of a pptx path or Presentation as a PIL image."""
    prs = pptx if hasattr(pptx, "slides") else Presentation(pptx)
    return _render(prs, prs.slides[index], width)


def render_presentation(pptx, output_dir, width=1920):
    """Save every slide as ``slide_<n>.png`` in ``output_dir``; returns the file paths."""
    prs = pptx if hasattr(pptx, "slides") else Presentation(pptx)
    os.makedirs(output_dir, exist_ok=True)
    paths = []
    for i, slide in enumerate(prs.slides, start=1):
        path = os.path.join(output_dir, f"slide_{i}.png")
        _render(prs, slide, width).save(path)
        paths.append(path)
    return paths
