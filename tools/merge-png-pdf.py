from __future__ import annotations

import argparse
from pathlib import Path

import img2pdf


def read_pages(list_path: Path) -> list[str]:
    pages = []
    for raw_line in list_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if line:
            pages.append(line)
    if not pages:
        raise SystemExit("PNG page list is empty.")
    return pages


def main() -> None:
    parser = argparse.ArgumentParser(description="Merge PNG pages into a PDF using a page list file.")
    parser.add_argument("--list", required=True, dest="list_path", help="UTF-8 text file containing one PNG path per line.")
    parser.add_argument("--output", required=True, help="Target PDF path.")
    args = parser.parse_args()

    pages = read_pages(Path(args.list_path))
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("wb") as handle:
        handle.write(img2pdf.convert(pages))


if __name__ == "__main__":
    main()
