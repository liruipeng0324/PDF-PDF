import argparse
import os

import pikepdf


def _extract_text_length(pdf_path: str, max_pages: int = 20) -> tuple[int, str]:
    try:
        from pypdf import PdfReader
    except Exception as exc:
        return 0, f"pypdf unavailable: {exc}"

    try:
        reader = PdfReader(pdf_path)
        total = 0
        for page in reader.pages[:max_pages]:
            total += len(page.extract_text() or "")
        return total, "pypdf"
    except Exception as exc:
        return 0, f"pypdf failed: {exc}"


def _text_length(pdf_path: str, max_pages: int = 20) -> int:
    total = 0
    with pikepdf.open(pdf_path) as pdf:
        for index, page in enumerate(pdf.pages):
            if index >= max_pages:
                break
            try:
                contents = page.get("/Contents", None)
                if contents is None:
                    continue
                data = contents.read_bytes() if hasattr(contents, "read_bytes") else b""
                total += data.count(b"Tj") + data.count(b"TJ") + data.count(b"'") + data.count(b'"')
            except Exception:
                continue
    return total


def verify_dual_pdf(pdf_path: str, before_pdf: str = "") -> bool:
    if not os.path.exists(pdf_path):
        print(f"VERIFY_FAILED file missing: {pdf_path}")
        return False
    size = os.path.getsize(pdf_path)
    if size < 1024:
        print(f"VERIFY_FAILED file too small: {size}")
        return False

    after_text = _text_length(pdf_path)
    before_text = _text_length(before_pdf) if before_pdf and os.path.exists(before_pdf) else 0
    extracted_text_len, extraction_method = _extract_text_length(pdf_path)
    print(f"VERIFY_FILE_SIZE {size}")
    print(f"VERIFY_TEXT_MARKERS_BEFORE {before_text}")
    print(f"VERIFY_TEXT_MARKERS_AFTER {after_text}")
    print(f"VERIFY_EXTRACTED_TEXT_LENGTH {extracted_text_len}")
    print(f"VERIFY_TEXT_EXTRACTION_METHOD {extraction_method}")

    if extracted_text_len <= 0 and after_text <= 0:
        print("VERIFY_FAILED no searchable text layer detected")
        return False
    if extracted_text_len <= 0:
        if before_text and after_text <= before_text:
            print("VERIFY_WARNING text marker count did not increase; Acrobat may have saved OCR back to the source PDF before verification")
        print("VERIFY_WARNING text extraction returned empty text; PDF stream markers suggest text may exist")
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify generated searchable-image PDF.")
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--before", default="")
    args = parser.parse_args()
    return 0 if verify_dual_pdf(args.pdf, args.before) else 1


if __name__ == "__main__":
    raise SystemExit(main())
