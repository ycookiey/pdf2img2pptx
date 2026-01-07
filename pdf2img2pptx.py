#!/usr/bin/env python3
"""
PDF to Image to PPTX Converter
PDFを画像化し、画像群をスライドサイズいっぱいに配置したPPTXファイルを自動生成するツール
"""

import argparse
import os
import sys
import tempfile
import shutil
from pathlib import Path
from typing import List, Optional

try:
    import fitz  # PyMuPDF
except ImportError:
    print("Error: PyMuPDF is not installed. Run: pip install PyMuPDF", file=sys.stderr)
    sys.exit(1)

try:
    from pptx import Presentation
    from pptx.util import Inches
except ImportError:
    print("Error: python-pptx is not installed. Run: pip install python-pptx", file=sys.stderr)
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("Error: Pillow is not installed. Run: pip install Pillow", file=sys.stderr)
    sys.exit(1)


# サポートされる画像形式
SUPPORTED_IMAGE_FORMATS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')


def pdf_to_images(pdf_path: str, output_dir: str, dpi: int = 600) -> List[str]:
    """
    PDFファイルを画像に変換

    Args:
        pdf_path: 入力PDFファイルパス
        output_dir: 出力ディレクトリ
        dpi: 解像度（デフォルト: 600）

    Returns:
        生成された画像ファイルパスのリスト

    Raises:
        FileNotFoundError: PDFファイルが見つからない場合
        Exception: PDF処理でエラーが発生した場合
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDFファイルが見つかりません: {pdf_path}")

    # 出力ディレクトリを作成
    os.makedirs(output_dir, exist_ok=True)

    print(f"PDFを開いています: {pdf_path}")

    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        print(f"総ページ数: {total_pages}")

        image_paths = []

        for page_num in range(total_pages):
            page = doc[page_num]

            # DPIに基づいてスケーリング行列を作成（72dpiが基準）
            mat = fitz.Matrix(dpi / 72, dpi / 72)

            # ページを画像化
            pix = page.get_pixmap(matrix=mat)

            # ファイル名を3桁のゼロパディングで生成
            output_filename = f"page_{page_num + 1:03d}.png"
            output_path = os.path.join(output_dir, output_filename)

            # PNG形式で保存
            pix.save(output_path)
            image_paths.append(output_path)

            print(f"  ページ {page_num + 1}/{total_pages} を変換しました: {output_filename}")

        doc.close()
        print(f"全ページの変換が完了しました。出力先: {output_dir}")

        return image_paths

    except Exception as e:
        raise Exception(f"PDF処理中にエラーが発生しました: {str(e)}")


def images_to_pptx(image_source: str, output_pptx: str) -> None:
    """
    画像ファイルまたはディレクトリからPPTXファイルを生成

    Args:
        image_source: 画像ファイルまたはディレクトリのパス
        output_pptx: 出力PPTXファイルパス

    Raises:
        FileNotFoundError: 画像ファイルまたはディレクトリが見つからない場合
        ValueError: 有効な画像ファイルが見つからない場合
    """
    if not os.path.exists(image_source):
        raise FileNotFoundError(f"画像ソースが見つかりません: {image_source}")

    # 画像ファイルリストを取得
    image_files = []

    if os.path.isfile(image_source):
        # 単一ファイルの場合
        if image_source.lower().endswith(SUPPORTED_IMAGE_FORMATS):
            image_files = [image_source]
    elif os.path.isdir(image_source):
        # ディレクトリの場合
        for filename in sorted(os.listdir(image_source)):
            if filename.lower().endswith(SUPPORTED_IMAGE_FORMATS):
                image_files.append(os.path.join(image_source, filename))

    if not image_files:
        raise ValueError(f"有効な画像ファイルが見つかりません: {image_source}")

    print(f"画像ファイル数: {len(image_files)}")

    # プレゼンテーションを作成
    prs = Presentation()

    # スライドサイズを16:9に設定
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    # 空白レイアウトを使用（インデックス6）
    blank_layout = prs.slide_layouts[6]

    for idx, image_path in enumerate(image_files, start=1):
        print(f"  スライド {idx}/{len(image_files)} を作成中: {os.path.basename(image_path)}")

        # 画像のサイズを取得
        with Image.open(image_path) as img:
            img_width, img_height = img.size

        # 画像のアスペクト比
        img_aspect = img_width / img_height

        # スライドのアスペクト比
        slide_aspect = prs.slide_width / prs.slide_height

        # スライドサイズいっぱいに画像を配置
        # アスペクト比を維持しながら、はみ出す部分をクロップ
        if img_aspect > slide_aspect:
            # 画像が横長の場合、高さを基準にする
            height = prs.slide_height
            width = int(height * img_aspect)
            left = int((prs.slide_width - width) / 2)
            top = 0
        else:
            # 画像が縦長の場合、幅を基準にする
            width = prs.slide_width
            height = int(width / img_aspect)
            left = 0
            top = int((prs.slide_height - height) / 2)

        # スライドを追加
        slide = prs.slides.add_slide(blank_layout)

        # 画像を追加
        slide.shapes.add_picture(image_path, left, top, width, height)

    # PPTXファイルを保存
    prs.save(output_pptx)
    print(f"PPTXファイルを保存しました: {output_pptx}")


def pdf_to_pptx(pdf_path: str, output_pptx: str, dpi: int = 600,
                keep_images: bool = False, image_dir: Optional[str] = None) -> None:
    """
    PDFから直接PPTXファイルを生成

    Args:
        pdf_path: 入力PDFファイルパス
        output_pptx: 出力PPTXファイルパス
        dpi: 解像度（デフォルト: 600）
        keep_images: 中間画像を保持するか（デフォルト: False）
        image_dir: 中間画像の保存先（デフォルト: 一時ディレクトリ）
    """
    # 一時ディレクトリまたは指定されたディレクトリを使用
    if image_dir:
        temp_dir = image_dir
        os.makedirs(temp_dir, exist_ok=True)
        is_temp = False
    else:
        temp_dir = tempfile.mkdtemp(prefix="pdf2pptx_")
        is_temp = True

    try:
        # PDF → 画像
        print("ステップ1: PDFを画像に変換中...")
        pdf_to_images(pdf_path, temp_dir, dpi)

        # 画像 → PPTX
        print("\nステップ2: 画像からPPTXを生成中...")
        images_to_pptx(temp_dir, output_pptx)

        print("\n変換が完了しました！")

        if keep_images:
            print(f"中間画像は保持されています: {temp_dir}")

    finally:
        # 一時ディレクトリをクリーンアップ
        if is_temp and not keep_images:
            shutil.rmtree(temp_dir)
            print(f"一時ファイルを削除しました: {temp_dir}")


def main():
    """メインエントリーポイント"""
    parser = argparse.ArgumentParser(
        description='PDF to Image to PPTX Converter',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # PDFを画像に変換
  %(prog)s pdf2img input.pdf --output ./images --dpi 600

  # 画像からPPTXを生成
  %(prog)s img2pptx ./images --output output.pptx

  # PDFから直接PPTXを生成
  %(prog)s pdf2pptx input.pdf --output output.pptx --dpi 600
        """
    )

    subparsers = parser.add_subparsers(dest='command', help='実行するコマンド')

    # pdf2img サブコマンド
    parser_pdf2img = subparsers.add_parser('pdf2img', help='PDFを画像に変換')
    parser_pdf2img.add_argument('input', help='入力PDFファイルパス')
    parser_pdf2img.add_argument('--output', default='./images', help='出力ディレクトリ（デフォルト: ./images）')
    parser_pdf2img.add_argument('--dpi', type=int, default=600, help='解像度（デフォルト: 600）')

    # img2pptx サブコマンド
    parser_img2pptx = subparsers.add_parser('img2pptx', help='画像からPPTXを生成')
    parser_img2pptx.add_argument('input', help='画像ファイルまたはディレクトリパス')
    parser_img2pptx.add_argument('--output', required=True, help='出力PPTXファイルパス')

    # pdf2pptx サブコマンド
    parser_pdf2pptx = subparsers.add_parser('pdf2pptx', help='PDFから直接PPTXを生成')
    parser_pdf2pptx.add_argument('input', help='入力PDFファイルパス')
    parser_pdf2pptx.add_argument('--output', required=True, help='出力PPTXファイルパス')
    parser_pdf2pptx.add_argument('--dpi', type=int, default=600, help='解像度（デフォルト: 600）')
    parser_pdf2pptx.add_argument('--keep-images', action='store_true', help='中間画像を削除せずに保持')
    parser_pdf2pptx.add_argument('--image-dir', help='中間画像の保存先（デフォルト: 一時ディレクトリ）')

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    try:
        if args.command == 'pdf2img':
            pdf_to_images(args.input, args.output, args.dpi)

        elif args.command == 'img2pptx':
            images_to_pptx(args.input, args.output)

        elif args.command == 'pdf2pptx':
            pdf_to_pptx(args.input, args.output, args.dpi,
                       args.keep_images, args.image_dir)

    except Exception as e:
        print(f"エラー: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
