# render-pptx

Export slides of `.pptx` files as `.png` images, using `python-pptx` and `Pillow`.
All shapes of a slide are drawn on a canvas, including shapes inherited from the
slide layout and slide master. Picture cropping is respected.

## Install

```
pip install .
```

## Usage

You can convert a `Presentation` object into a list of PIL `Image`s like this:

```python
from pptx import Presentation
from render_pptx import render_slide, render_slides, render_presentation

prs = Presentation("deck.pptx")
all_images = render_slides(prs) # returns list of PIL Images
```

Or you can convert a `.pptx` file to multiple `.png` files using the command line:

```
render-pptx deck.pptx .
```

