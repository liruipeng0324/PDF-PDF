from __future__ import annotations

import argparse
import concurrent.futures
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image


def render_page(input_pdf: str, output_dir: str, page_index: int, dpi: int, image_format: str, jpeg_quality: int) -> str:
    pdf = pdfium.PdfDocument(input_pdf)
    page = pdf[page_index]
    scale = dpi / 72.0
    bitmap = page.render(scale=scale)
    image = bitmap.to_pil()
    extension = "jpg" if image_format == "jpeg" else "png"
    output_path = Path(output_dir) / f"page-{page_index + 1:06d}.{extension}"
    if image_format == "jpeg":
        if image.mode == "RGBA":
            background = Image.new("RGB", image.size, "white")
            background.paste(image, mask=image.getchannel("A"))
            image = background
        elif image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
        image.save(output_path, quality=jpeg_quality, optimize=True, dpi=(dpi, dpi))
    else:
        image.save(output_path, dpi=(dpi, dpi))
    page.close()
    pdf.close()
    return str(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Render PDF pages to images with PDFium.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--format", choices=("png", "jpeg"), default="png")
    parser.add_argument("--jpeg-quality", type=int, default=75)
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(args.input)
    page_count = len(pdf)
    pdf.close()

    workers = max(1, min(args.workers, page_count))
    print(f"PDFIUM_TOTAL_PAGES {page_count}", flush=True)

    if workers == 1:
        for page_index in range(page_count):
            render_page(args.input, args.output_dir, page_index, args.dpi, args.format, args.jpeg_quality)
            print(f"IMAGE_PAGE {page_index + 1}/{page_count}", flush=True)
        return 0

    completed = 0
    with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
        futures = [
            executor.submit(render_page, args.input, args.output_dir, page_index, args.dpi, args.format, args.jpeg_quality)
            for page_index in range(page_count)
        ]
        for future in concurrent.futures.as_completed(futures):
            future.result()
            completed += 1
            print(f"IMAGE_PAGE {completed}/{page_count}", flush=True)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
