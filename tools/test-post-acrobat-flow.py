import argparse
import subprocess
import sys
from pathlib import Path

import pikepdf

ROOT = Path(__file__).resolve().parents[1]
RESTORE_SCRIPT = ROOT / "tools" / "restore-bookmarks.py"


def count_outline_items(pdf_path: Path) -> int:
    count = 0
    with pikepdf.open(pdf_path) as pdf:
        outlines = pdf.Root.get("/Outlines")
        if outlines is None:
            return 0

        def walk(node):
            nonlocal count
            while node is not None:
                count += 1
                child = node.get("/First")
                if child is not None:
                    walk(child)
                node = node.get("/Next")

        first = outlines.get("/First")
        if first is not None:
            walk(first)
    return count


def text_marker_count(pdf_path: Path, max_pages: int = 20) -> int:
    total = 0
    with pikepdf.open(pdf_path) as pdf:
        for index, page in enumerate(pdf.pages):
            if index >= max_pages:
                break
            contents = page.get("/Contents", None)
            if contents is None:
                continue
            try:
                data = contents.read_bytes() if hasattr(contents, "read_bytes") else b""
            except Exception:
                continue
            total += data.count(b"Tj") + data.count(b"TJ") + data.count(b"'") + data.count(b'"')
    return total


def main() -> int:
    parser = argparse.ArgumentParser(description="Test the post-Acrobat OCR flow without running Acrobat.")
    parser.add_argument("--tagged", required=True, help="Tagged/bookmarked PDF generated before image rendering.")
    parser.add_argument("--scanned", required=True, help="Acrobat OCR output PDF.")
    parser.add_argument("--output", required=True, help="Final PDF path to write.")
    args = parser.parse_args()

    tagged = Path(args.tagged)
    scanned = Path(args.scanned)
    output = Path(args.output)

    print(f"POST_TEST_TAGGED {tagged}")
    print(f"POST_TEST_SCANNED {scanned}")
    print(f"POST_TEST_OUTPUT {output}")

    if not tagged.exists():
        print("POST_TEST_FAILED tagged PDF missing")
        return 1
    if not scanned.exists():
        print("POST_TEST_FAILED scanned PDF missing")
        return 1

    markers = text_marker_count(scanned)
    print(f"POST_TEST_SCANNED_TEXT_MARKERS {markers}")
    if markers <= 0:
        print("POST_TEST_FAILED scanned PDF has no text layer markers")
        return 1

    output.parent.mkdir(parents=True, exist_ok=True)
    command = [
        sys.executable,
        str(RESTORE_SCRIPT),
        "--source",
        str(tagged),
        "--target",
        str(scanned),
        "--output",
        str(output),
    ]
    result = subprocess.run(command, cwd=str(ROOT))
    print(f"POST_TEST_RESTORE_EXIT {result.returncode}")
    if result.returncode != 0:
        return result.returncode

    if not output.exists() or output.stat().st_size < 1024:
        print("POST_TEST_FAILED final PDF missing or too small")
        return 1

    final_markers = text_marker_count(output)
    outline_count = count_outline_items(output)
    print(f"POST_TEST_FINAL_SIZE {output.stat().st_size}")
    print(f"POST_TEST_FINAL_TEXT_MARKERS {final_markers}")
    print(f"POST_TEST_FINAL_BOOKMARKS {outline_count}")
    if final_markers <= 0:
        print("POST_TEST_FAILED final PDF has no text layer markers")
        return 1
    if outline_count <= 0:
        print("POST_TEST_FAILED final PDF has no bookmarks")
        return 1

    print("POST_TEST_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
