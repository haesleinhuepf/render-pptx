import io

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from render_pptx import render_presentation, render_slide


def _picture_bytes():
    img = Image.new("RGB", (100, 100), (255, 0, 0))
    img.paste((0, 0, 255), (50, 0, 100, 100))  # right half blue
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def _deck(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(2), Inches(2))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0, 255, 0)
    pic = slide.shapes.add_picture(_picture_bytes(), Inches(4), Inches(4), Inches(2), Inches(2))
    pic.crop_left = 0.5  # keep only blue half
    path = tmp_path / "t.pptx"
    prs.save(path)
    return path


def test_shape_and_cropped_picture(tmp_path):
    img = render_slide(_deck(tmp_path), 0, width=1000)
    assert img.size == (1000, 750)
    assert img.getpixel((50, 50)) == (0, 255, 0)
    assert img.getpixel((500, 500)) == (0, 0, 255)  # no red from cropped half
    assert img.getpixel((900, 700)) == (255, 255, 255)


def test_export_png(tmp_path):
    paths = render_presentation(_deck(tmp_path), tmp_path / "out", width=500)
    assert len(paths) == 1 and Image.open(paths[0]).format == "PNG"


def _text_deck(tmp_path, text, rotation=0, tabs=()):
    from lxml import etree
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1))
    tb.rotation = rotation
    para = tb.text_frame.paragraphs[0]
    para.add_run().text = text
    if tabs:
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        pPr = para._p.get_or_add_pPr()
        lst = etree.SubElement(pPr, "{%s}tabLst" % ns)
        for pos in tabs:
            etree.SubElement(lst, "{%s}tab" % ns, pos=str(int(Inches(pos))), algn="l")
    path = tmp_path / "txt.pptx"
    prs.save(path)
    return path


def _dark_columns(img):
    return [x for x in range(img.width)
            if any(img.getpixel((x, y))[0] < 100 for y in range(img.height))]


def test_tab_uses_ruler_position(tmp_path):
    img = render_slide(_text_deck(tmp_path, "A\tB", tabs=(4,)), 0, width=1000)
    cols = _dark_columns(img)
    # "B" starts at 1in margin + 0.1in inset + 4in tab = 5.1in -> 510px of 1000px (10in)
    assert any(500 <= x <= 530 for x in cols)
    assert not any(130 < x < 500 for x in cols)


def test_text_rotation(tmp_path):
    img = render_slide(_text_deck(tmp_path, "HELLO WORLD", rotation=90), 0, width=1000)
    cols = _dark_columns(img)
    # horizontal text spans wide; rotated by 90 it becomes narrow
    assert max(cols) - min(cols) < 100


def test_indentation_without_tab(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1))
    para = tb.text_frame.paragraphs[0]
    para.add_run().text = "Indented"
    para._p.get_or_add_pPr().set("marL", str(int(Inches(2))))
    path = tmp_path / "ind.pptx"
    prs.save(path)
    cols = _dark_columns(render_slide(path, 0, width=1000))
    assert min(cols) >= 300  # 1in box + 0.1in inset + 2in indent -> 310px


def test_font_one_point_smaller(tmp_path):
    from render_pptx import renderer
    sizes = []
    orig = renderer._font
    renderer._font = lambda size, bold=False, name=None, italic=False: sizes.append(size) or orig(size, bold, name, italic)
    try:
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        run = tb.text_frame.paragraphs[0].add_run()
        run.text = "x"
        from pptx.util import Pt
        run.font.size = Pt(20)
        path = tmp_path / "f.pptx"
        prs.save(path)
        render_slide(path, 0, width=720)  # 720px / 10in = 72 px per in = 1px per pt
    finally:
        renderer._font = orig
    assert sizes == [19]


def test_theme_text_color(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = "X"
    run.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
    path = tmp_path / "theme-color.pptx"
    prs.save(path)

    img = render_slide(path, 0, width=720)
    pixels = img.load()
    assert any(pixels[x, y] == (79, 129, 189)
               for y in range(img.height) for x in range(img.width))


def test_theme_text_color_tint(tmp_path):
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = "X"
    run.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
    color = run._r.rPr.find("{%s}solidFill/{%s}schemeClr" % (ns, ns))
    etree.SubElement(color, "{%s}tint" % ns, val="50000")
    path = tmp_path / "theme-color-tint.pptx"
    prs.save(path)

    img = render_slide(path, 0, width=720)
    pixels = img.load()
    assert any(pixels[x, y] == (167, 192, 222)
               for y in range(img.height) for x in range(img.width))


def test_inherited_theme_text_color(tmp_path):
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb.text_frame.paragraphs[0].add_run().text = "X"
    lst = tb.text_frame._txBody.find("{%s}lstStyle" % ns)
    lvl = etree.SubElement(lst, "{%s}lvl1pPr" % ns)
    rpr = etree.SubElement(lvl, "{%s}defRPr" % ns)
    fill = etree.SubElement(rpr, "{%s}solidFill" % ns)
    etree.SubElement(fill, "{%s}schemeClr" % ns, val="accent1")
    path = tmp_path / "inherited-theme-color.pptx"
    prs.save(path)

    img = render_slide(path, 0, width=720)
    pixels = img.load()
    assert any(pixels[x, y] == (79, 129, 189)
               for y in range(img.height) for x in range(img.width))


def test_inherited_size_and_font_name(tmp_path):
    from render_pptx import renderer
    seen = []
    orig = renderer._font
    renderer._font = lambda size, bold=False, name=None, italic=False: seen.append((size, name)) or orig(size, bold, name, italic)
    try:
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])  # title slide
        slide.shapes.title.text = "T"  # size from master title style (44pt), theme font
        tb = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(4), Inches(1))
        run = tb.text_frame.paragraphs[0].add_run()
        run.text = "x"
        run.font.name = "DejaVu Sans"
        path = tmp_path / "i.pptx"
        prs.save(path)
        render_slide(path, 0, width=720)
    finally:
        renderer._font = orig
    assert seen[0][0] > 19  # title is not rendered with the 18pt fallback
    assert seen[0][1] == "Calibri"
    assert seen[-1] == (17, "DejaVu Sans")


def test_inherited_indentation_and_tabs(tmp_path):
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1))
    tf = tb.text_frame
    tf.paragraphs[0].add_run().text = "A"
    tf.add_paragraph().add_run().text = "B"
    # indentation and tab stops defined once in the shape's list style, not per paragraph
    lst = tf._txBody.find("{%s}lstStyle" % ns)
    lvl = etree.SubElement(lst, "{%s}lvl1pPr" % ns, marL=str(int(Inches(2))))
    path = tmp_path / "inh.pptx"
    prs.save(path)  # (list style is inserted before save so python-pptx keeps it)
    img = render_slide(path, 0, width=1000)
    assert min(_dark_columns(img)) >= 300


def test_table_theme_fill_modifiers(tmp_path):
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(1, 1, Inches(1), Inches(1), Inches(2), Inches(2)).table
    cell = table.cell(0, 0)
    cell.fill.solid()
    cell.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_6
    scheme_color = cell._tc.tcPr.find("{%s}solidFill/{%s}schemeClr" % (ns, ns))
    etree.SubElement(scheme_color, "{%s}lumMod" % ns, val="60000")
    etree.SubElement(scheme_color, "{%s}lumOff" % ns, val="40000")
    path = tmp_path / "table-theme-fill.pptx"
    prs.save(path)

    img = render_slide(path, 0, width=720)
    assert img.getpixel((144, 144)) == (250, 192, 144)


def test_table_style_alternating_fills_and_cell_override(tmp_path):
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(3, 1, Inches(1), Inches(1), Inches(2), Inches(3)).table
    cell = table.cell(2, 0)
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(0, 255, 0)
    path = tmp_path / "table-style-fills.pptx"
    prs.save(path)

    prs = Presentation(path)
    table = next(shape.table for shape in prs.slides[0].shapes if shape.has_table)
    style_id = "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"
    table._tbl.tblPr.find("{%s}tableStyleId" % ns).text = style_id
    table._tbl.tblPr.set("firstRow", "0")
    table._tbl.tblPr.set("bandRow", "1")

    style_part = prs.part.part_related_by(RT.TABLE_STYLES)
    styles = etree.fromstring(style_part.blob)
    style = etree.SubElement(
        styles, "{%s}tblStyle" % ns,
        styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
        styleName="Medium Style 2 - Accent 1")
    whole_fill = etree.SubElement(etree.SubElement(
        etree.SubElement(style, "{%s}wholeTbl" % ns), "{%s}tcStyle" % ns),
        "{%s}fill" % ns)
    etree.SubElement(etree.SubElement(whole_fill, "{%s}solidFill" % ns),
                     "{%s}schemeClr" % ns, val="accent1")
    band_fill = etree.SubElement(etree.SubElement(
        etree.SubElement(style, "{%s}band1H" % ns), "{%s}tcStyle" % ns),
        "{%s}fill" % ns)
    band_color = etree.SubElement(etree.SubElement(
        band_fill, "{%s}solidFill" % ns), "{%s}schemeClr" % ns, val="accent1")
    etree.SubElement(band_color, "{%s}tint" % ns, val="50000")
    style_part._blob = etree.tostring(styles, xml_declaration=True, encoding="UTF-8")

    prs.save(path)

    img = render_slide(path, 0, width=720)
    assert img.getpixel((144, 108)) == (79, 129, 189)
    assert img.getpixel((144, 180)) == (79, 129, 189)
    assert img.getpixel((144, 252)) == (0, 255, 0)


def test_table_style_uses_declared_theme_color(tmp_path):
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(1, 1, Inches(1), Inches(1), Inches(2), Inches(2)).table
    table._tbl.tblPr.set("firstRow", "0")
    table._tbl.tblPr.set("bandRow", "1")
    style_id = "{8A0DE7A6-89E2-40E7-9C1E-5849B6C87C28}"
    table._tbl.tblPr.find("{%s}tableStyleId" % ns).text = style_id

    style_part = prs.part.part_related_by(RT.TABLE_STYLES)
    styles = etree.fromstring(style_part.blob)
    style = etree.SubElement(styles, "{%s}tblStyle" % ns, styleId=style_id)
    fill = etree.SubElement(
        etree.SubElement(
            etree.SubElement(
                etree.SubElement(style, "{%s}band1H" % ns), "{%s}tcStyle" % ns),
            "{%s}fill" % ns),
        "{%s}solidFill" % ns)
    color = etree.SubElement(fill, "{%s}schemeClr" % ns, val="accent4")
    etree.SubElement(color, "{%s}alpha" % ns, val="40000")
    style_part._blob = etree.tostring(styles, xml_declaration=True, encoding="UTF-8")

    path = tmp_path / "table-style-alpha-fill.pptx"
    prs.save(path)
    img = render_slide(path, 0, width=720)
    from render_pptx.renderer import _theme_colors
    theme = _theme_colors(Presentation(path).slide_masters[0])
    assert img.getpixel((144, 144)) == theme["accent4"]


def test_preset_shapes_are_not_boxes(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for kind in (MSO_SHAPE.RIGHT_ARROW, MSO_SHAPE.DONUT, MSO_SHAPE.UP_ARROW):
        slide.shapes.add_shape(kind, Inches(1), Inches(1), Inches(2), Inches(2))
        img = render_slide_from(prs, tmp_path)
        w, h = img.size
        # corners of the shape's bounding box stay empty for arrows/donuts; donut centre is a hole
        x0, y0 = int(w * 1 / 10), int(h * 1 / 7.5)
        assert img.getpixel((x0 + 2, y0 + 2)) == (255, 255, 255)
        slide.shapes._spTree.remove(slide.shapes[-1]._element)


def render_slide_from(prs, tmp_path):
    path = tmp_path / "s.pptx"
    prs.save(path)
    return render_slide(path, 0, 1000).convert("RGB")


def test_outline_color_weight_and_dash(tmp_path):
    from pptx.enum.dml import MSO_LINE
    from pptx.util import Pt
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(4), Inches(2))
    shape.fill.background()
    shape.line.color.rgb = RGBColor(255, 0, 0)
    shape.line.width = Pt(6)
    shape.line.dash_style = MSO_LINE.DASH
    img = render_slide_from(prs, tmp_path)
    scale = img.width / prs.slide_width
    y = round(Inches(1) * scale)
    row = [img.getpixel((x, y)) for x in range(round(Inches(1) * scale) + 12, round(Inches(5) * scale) - 12)]
    red = [px == (255, 0, 0) for px in row]
    assert any(red) and not all(red)  # dashes with gaps
    assert all(px in ((255, 0, 0), (255, 255, 255)) or px[1] == px[2] for px in row)
    # thick line: pixels well above the nominal edge are still red inside a dash
    x_dash = next(x for x, r in zip(range(round(Inches(1) * scale) + 12, 10 ** 6), red) if r)
    px = img.getpixel((x_dash, y - 3))
    assert px[0] == 255 and px[1] < 30
