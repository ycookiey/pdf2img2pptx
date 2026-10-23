# pdf2img2pptx

Converts a PDF to a 16:9 PowerPoint (PPTX) file by rendering each page to a PNG image and placing one image on each slide. Because each slide is a picture, the layout of the PDF is kept, but the text cannot be edited or searched. The two steps can also be run separately: PDF to images, and images to PPTX.

## Installation

Requires Python 3.8 or later. No other program is needed: PyMuPDF renders the pages.

```bash
git clone https://github.com/ycookiey/pdf2img2pptx.git
cd pdf2img2pptx
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
# PDF to PPTX
python pdf2img2pptx.py pdf2pptx input.pdf --output output.pptx

# PDF to PNG images (page_001.png, page_002.png, ...)
python pdf2img2pptx.py pdf2img input.pdf --output ./images

# Images to PPTX: one slide per image, in file name order
python pdf2img2pptx.py img2pptx ./images --output output.pptx
```

| Option | Command | Description |
|---|---|---|
| `--output` | all | Output PPTX file (required), or for `pdf2img` the output directory (default `./images`) |
| `--dpi` | `pdf2pptx`, `pdf2img` | Rendering resolution. Default 600 |
| `--keep-images` | `pdf2pptx` | Keep the temporary directory with the page images. Its path is printed at the end |
| `--image-dir` | `pdf2pptx` | Write the page images to this directory instead of a temporary one, and keep them |

`img2pptx` takes a single image or a directory. Supported formats: PNG, JPEG, BMP, GIF.

## Limitations

- Slides are always 16:9. Each image is scaled to cover the whole slide and centered, so pages with another shape are cut off. On an A4 portrait page, only the middle 40% is visible.
- At the default 600 dpi, a 16:9 page of 10 x 5.625 inches becomes a 6000 x 3375 pixel image. Lower `--dpi` for smaller files.
- With `--image-dir`, images already in that directory are added as slides too.
- The `--help` text and the progress and error messages are in Japanese.

## Uninstall

Delete the cloned folder. The virtual environment is inside it.

## License

[MIT](LICENSE). PyMuPDF is licensed under AGPL-3.0 or a commercial license from Artifex. Check its terms if you distribute this tool together with PyMuPDF.
