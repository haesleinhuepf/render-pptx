# render-pptx

Export slides of `.pptx` files as `.png` images, using `python-pptx` and `Pillow`.
All shapes of a slide are drawn on a canvas, including shapes inherited from the
slide layout and slide master. Picture cropping is respected.

## Install

```
pip install .
```

## Usage

```python
from render_pptx import render_slide, render_presentation

img = render_slide("deck.pptx", 0)           # PIL.Image of first slide
paths = render_presentation("deck.pptx", "out/", width=1920)  # slide_1.png, ...
```

Command line:

```
render-pptx deck.pptx out_folder --width 1920
```

Supported: backgrounds (solid), auto shapes (rectangles, ellipses, others as
rectangles) with solid fill/line, text, pictures (with crop and rotation),
tables, groups, connectors/lines. Gradients and other effects are approximated or ignored.
