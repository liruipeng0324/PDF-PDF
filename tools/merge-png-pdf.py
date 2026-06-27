from __future__ import annotations

import argparse
from pathlib import Path

import img2pdf
import pikepdf


def read_pages(list_path: Path) -> list[str]:
    pages = []
    for raw_line in list_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if line:
            pages.append(line)
    if not pages:
        raise SystemExit("Image page list is empty.")
    return pages


def read_page_sizes(source_pdf: Path) -> list[tuple[float, float]]:
    sizes = []
    with pikepdf.open(source_pdf) as pdf:
        for page in pdf.pages:
            media_box = page.MediaBox
            width = float(media_box[2]) - float(media_box[0])
            height = float(media_box[3]) - float(media_box[1])
            sizes.append((width, height))
    return sizes


def source_page_layout_fun(page_sizes: list[tuple[float, float]]):
    index = {"value": 0}

    def layout_fun(imgwidthpx, imgheightpx, ndpi):
        page_index = index["value"]
        index["value"] += 1
        if page_index >= len(page_sizes):
            raise ValueError("More images than source PDF pages.")
        page_width, page_height = page_sizes[page_index]
        return page_width, page_height, page_width, page_height

    return layout_fun


def main() -> None:
    parser = argparse.ArgumentParser(description="Merge rendered page images into a PDF using a page list file.")
    parser.add_argument("--list", required=True, dest="list_path", help="UTF-8 text file containing one image path per line.")
    parser.add_argument("--source-pdf", required=True, help="Source PDF used to preserve exact page sizes.")
    parser.add_argument("--output", required=True, help="Target PDF path.")
    args = parser.parse_args()

    pages = read_pages(Path(args.list_path))
    page_sizes = read_page_sizes(Path(args.source_pdf))
    if len(page_sizes) != len(pages):
        raise SystemExit(f"Page count mismatch: {len(pages)} images, {len(page_sizes)} source PDF pages.")

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("wb") as handle:
        handle.write(img2pdf.convert(pages, layout_fun=source_page_layout_fun(page_sizes)))


if __name__ == "__main__":
    main()
