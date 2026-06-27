from __future__ import annotations

import argparse
import concurrent.futures
from pathlib import Path

import pypdfium2 as pdfium


def render_page(input_pdf: str, output_dir: str, page_index: int, dpi: int) -> str:
    pdf = pdfium.PdfDocument(input_pdf)
    page = pdf[page_index]
    scale = dpi / 72.0
    bitmap = page.render(scale=scale)
    image = bitmap.to_pil()
    output_path = Path(output_dir) / f"page-{page_index + 1:06d}.png"
    image.save(output_path)
    page.close()
    pdf.close()
    return str(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Render PDF pages to PNG with PDFium.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--workers", type=int, default=4)
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
            render_page(args.input, args.output_dir, page_index, args.dpi)
            print(f"PNG_PAGE {page_index + 1}/{page_count}", flush=True)
        return 0

    completed = 0
    with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
        futures = [
            executor.submit(render_page, args.input, args.output_dir, page_index, args.dpi)
            for page_index in range(page_count)
        ]
        for future in concurrent.futures.as_completed(futures):
            future.result()
            completed += 1
            print(f"PNG_PAGE {completed}/{page_count}", flush=True)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
