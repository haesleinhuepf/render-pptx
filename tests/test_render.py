import io

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
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
