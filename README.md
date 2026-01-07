# pdf2img2pptx

PDFを画像化し、16:9のPPTXファイルを自動生成するツール。
画像を経由することでレイアウトの崩れを防ぎます。

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### PDF to PPTX

```bash
python pdf2img2pptx.py pdf2pptx input.pdf --output output.pptx
```

### PDF to Images

```bash
python pdf2img2pptx.py pdf2img input.pdf --output ./images
```

### Images to PPTX

```bash
python pdf2img2pptx.py img2pptx ./images --output output.pptx
```

## Options

- `--dpi`: 解像度（デフォルト: 600）
- `--keep-images`: 中間画像を保持
- `--image-dir`: 中間画像の保存先

## Specifications

- Slide size: 16:9
- Image format: PNG (default 600dpi)
- Supported images: PNG, JPEG, JPG, BMP, GIF