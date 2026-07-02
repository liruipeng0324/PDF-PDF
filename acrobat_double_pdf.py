import argparse
import os
import shutil
import subprocess
import time
import traceback
from pathlib import Path

import win32com.client

from acrobat_ui_ocr import AcrobatUiOcrError, run_ocr_by_ui
from verify_dual_pdf import verify_dual_pdf, _text_length


DEFAULT_ACROBAT_EXE = r"D:\Acrobat\Acrobat\Acrobat.exe"


def _kill_acrobat() -> None:
    subprocess.run("taskkill /f /im Acrobat.exe 2>nul", shell=True)
    subprocess.run("taskkill /f /im AcroCEF.exe 2>nul", shell=True)


def _find_acrobat_exe() -> str:
    candidates = [
        DEFAULT_ACROBAT_EXE,
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files\Adobe\Acrobat\Acrobat\Acrobat.exe",
        r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
    ]
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return "Acrobat.exe"


def _active_window_title() -> str:
    try:
        from pywinauto import Desktop

        windows = Desktop(backend="uia").windows(title_re=".*(Adobe Acrobat|Acrobat|Save|Error|Replace|Protected).*")
        return " | ".join(window.window_text() for window in windows if window.window_text())
    except Exception as exc:
        return f"unavailable: {exc}"


def _print_open_diagnostics(src_pdf: str, acrobat_exe: str) -> None:
    print(f"ACROBAT_OPEN_DIAG_INPUT {src_pdf}")
    print(f"ACROBAT_OPEN_DIAG_EXISTS {os.path.exists(src_pdf)}")
    print(f"ACROBAT_OPEN_DIAG_SIZE {os.path.getsize(src_pdf) if os.path.exists(src_pdf) else -1}")
    print(f"ACROBAT_OPEN_DIAG_CWD {os.getcwd()}")
    print(f"ACROBAT_OPEN_DIAG_EXE {acrobat_exe}")
    print(f"ACROBAT_OPEN_DIAG_WINDOWS {_active_window_title()}")


def _window_has_loaded_pdf(src_pdf: str) -> bool:
    try:
        from pywinauto import Desktop
    except Exception as exc:
        print(f"ACROBAT_OPEN_UI_CHECK_UNAVAILABLE {exc}")
        return False

    file_name = os.path.basename(src_pdf).lower()
    file_stem = os.path.splitext(file_name)[0]
    try:
        for window in Desktop(backend="uia").windows(title_re=".*(Adobe Acrobat|Acrobat).*"):
            title = (window.window_text() or "").lower()
            if file_name and file_name in title:
                print(f"ACROBAT_OPEN_WINDOW_TITLE_MATCH {window.window_text()}")
                return True
            if file_stem and file_stem in title and ".pdf" in title:
                print(f"ACROBAT_OPEN_WINDOW_TITLE_MATCH {window.window_text()}")
                return True

            try:
                texts = []
                for control in window.descendants():
                    text = control.window_text() or control.element_info.name or ""
                    if text:
                        texts.append(text)
                joined = "\n".join(texts)
                if ("/" in joined and any(token in joined for token in ("1 /", "1/", "Page", "页"))) or ".pdf" in title:
                    print(f"ACROBAT_OPEN_PAGE_CONTROL_DETECTED {window.window_text()}")
                    return True
            except Exception:
                continue
    except Exception as exc:
        print(f"ACROBAT_OPEN_UI_CHECK_FAILED {exc}")
    return False


def _wait_for_loaded_pdf_window(src_pdf: str, timeout: int = 30) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _window_has_loaded_pdf(src_pdf):
            return True
        time.sleep(1)
    return False


def _open_pdf_visible(src_pdf: str, timeout: int = 45):
    app = win32com.client.Dispatch("AcroExch.App")
    av_doc = None
    acrobat_exe = _find_acrobat_exe()

    _print_open_diagnostics(src_pdf, acrobat_exe)
    print("ACROBAT_OPEN_METHOD subprocess")
    subprocess.Popen([acrobat_exe, src_pdf], shell=False)

    if not _wait_for_loaded_pdf_window(src_pdf, 30):
        print("ACROBAT_OPEN_METHOD os.startfile")
        os.startfile(src_pdf)

    if not _wait_for_loaded_pdf_window(src_pdf, 30):
        print("ACROBAT_OPEN_METHOD AVDoc.Open")
        av_doc = win32com.client.Dispatch("AcroExch.AVDoc")
        app.Show()
        open_ok = av_doc.Open(src_pdf, "")
        print(f"AVDoc.Open result: {open_ok}")

    if not _wait_for_loaded_pdf_window(src_pdf, 30):
        _print_open_diagnostics(src_pdf, acrobat_exe)
        raise RuntimeError("Acrobat window title did not show the PDF file name; PDF open failed.")

    print("ACROBAT_OPEN_SUCCESS True")
    return app, av_doc, None


def _handle_save_dialogs() -> list[str]:
    handled: list[str] = []
    try:
        from pywinauto import Desktop
    except Exception as exc:
        print(f"ACROBAT_SAVE_DIALOG_CHECK_UNAVAILABLE {exc}")
        return handled

    main_window_titles = {"Adobe Acrobat", "Adobe Acrobat (64-bit)", "Adobe Acrobat Pro (64-bit)"}
    title_markers = (
        "\u4fdd\u5b58",
        "\u53e6\u5b58\u4e3a",
        "\u786e\u8ba4\u53e6\u5b58\u4e3a",
        "\u662f\u5426\u66ff\u6362",
        "\u9519\u8bef",
        "\u53d7\u4fdd\u62a4",
        "Save",
        "Replace",
        "Error",
        "Protected",
    )
    button_markers = (
        "\u786e\u5b9a",
        "\u662f",
        "\u66ff\u6362",
        "\u4fdd\u5b58",
        "OK",
        "Yes",
        "Replace",
        "Save",
    )

    try:
        for window in Desktop(backend="uia").windows():
            title = window.window_text() or ""
            if title.strip() in main_window_titles:
                continue
            if not any(marker.lower() in title.lower() for marker in title_markers):
                continue
            handled.append(title)
            print(f"ACROBAT_SAVE_DIALOG_DETECTED {title}")
            try:
                window.set_focus()
            except Exception:
                pass
            for button in window.descendants(control_type="Button"):
                text = button.window_text() or button.element_info.name or ""
                if any(marker.lower() in text.lower() for marker in button_markers):
                    print(f"ACROBAT_SAVE_DIALOG_CLICK {title} -> {text}")
                    try:
                        button.click_input()
                    except Exception:
                        button.invoke()
                    time.sleep(1)
                    break
    except Exception as exc:
        print(f"ACROBAT_SAVE_DIALOG_CHECK_FAILED {exc}")
    return handled


def _file_signature(path: str) -> tuple[int, float]:
    if not os.path.exists(path):
        return 0, 0.0
    stat = os.stat(path)
    return stat.st_size, stat.st_mtime


def _source_has_text_layer(path: str) -> bool:
    try:
        return _text_length(path) > 0
    except Exception as exc:
        print(f"ACROBAT_SAVE_SOURCE_TEXT_CHECK_FAILED {exc}")
        return False


def _wait_for_file_saved(path: str, before_signature: tuple[int, float], timeout: int = 180) -> bool:
    deadline = time.time() + timeout
    changed = False
    stable_rounds = 0
    last_signature = before_signature
    print(f"ACROBAT_SAVE_SOURCE_BEFORE size={before_signature[0]} mtime={before_signature[1]}")

    while time.time() < deadline:
        current = _file_signature(path)
        if _source_has_text_layer(path):
            print(f"ACROBAT_SAVE_SOURCE_TEXT_LAYER_READY size={current[0]} mtime={current[1]}")
            return True
        if current != before_signature:
            changed = True
        if changed and current == last_signature and current[0] > 1024:
            stable_rounds += 1
            if stable_rounds >= 4:
                print(f"ACROBAT_SAVE_SOURCE_AFTER size={current[0]} mtime={current[1]}")
                return True
        else:
            stable_rounds = 0
        last_signature = current
        time.sleep(2)

    current = _file_signature(path)
    print(f"ACROBAT_SAVE_SOURCE_WAIT_TIMEOUT size={current[0]} mtime={current[1]}")
    return changed


def _save_current_file_by_ui(src_pdf: str) -> bool:
    try:
        from pywinauto import Desktop, keyboard
    except Exception as exc:
        print(f"ACROBAT_UI_SAVE_UNAVAILABLE {exc}")
        return False

    before_signature = _file_signature(src_pdf)
    try:
        windows = Desktop(backend="uia").windows(
            title_re=".*(Adobe Acrobat|Acrobat).*",
            visible_only=True,
            enabled_only=False,
        )
        pdf_windows = [w for w in windows if ".pdf" in (w.window_text() or "").lower()]
        win = (pdf_windows or windows)[0]
        win.set_focus()
        print("ACROBAT_UI_SAVE_HOTKEY_CTRL_S")
        keyboard.send_keys("^s")
        _handle_save_dialogs()
        return _wait_for_file_saved(src_pdf, before_signature)
    except Exception as exc:
        print(f"ACROBAT_UI_SAVE_FAILED {exc}")
        return False


def _get_active_pddoc(app):
    try:
        active_avdoc = app.GetActiveDoc()
    except Exception as exc:
        print(f"ACROBAT_GET_ACTIVE_DOC_FAILED {exc}")
        print(f"ACROBAT_GET_ACTIVE_DOC_WINDOWS {_active_window_title()}")
        raise
    if active_avdoc is None:
        print(f"ACROBAT_GET_ACTIVE_DOC_NONE_WINDOWS {_active_window_title()}")
        raise RuntimeError("Acrobat has no active document after OCR.")
    active_pddoc = active_avdoc.GetPDDoc()
    if active_pddoc is None:
        raise RuntimeError("Acrobat active document has no PDDoc after OCR.")
    return active_avdoc, active_pddoc


def _print_save_diagnostics(src_pdf: str, out_pdf: str, dialogs: list[str]) -> None:
    output_dir = os.path.dirname(out_pdf)
    print(f"ACROBAT_SAVE_DIAG_INPUT {src_pdf}")
    print(f"ACROBAT_SAVE_DIAG_OUTPUT {out_pdf}")
    print(f"ACROBAT_SAVE_DIAG_OUTPUT_DIR_EXISTS {os.path.isdir(output_dir)}")
    print(f"ACROBAT_SAVE_DIAG_OUTPUT_EXISTS {os.path.exists(out_pdf)}")
    print(f"ACROBAT_SAVE_DIAG_ACTIVE_WINDOWS {_active_window_title()}")
    print(f"ACROBAT_SAVE_DIAG_DIALOGS {dialogs if dialogs else 'none'}")


def _save_pdf_after_ocr(app, src_pdf: str, out_pdf: str) -> None:
    print("ACROBAT_SAVE_WAIT_AFTER_OCR 8")
    time.sleep(8)
    ui_saved = _save_current_file_by_ui(src_pdf)
    print(f"ACROBAT_UI_SAVE_RESULT {ui_saved}")
    _handle_save_dialogs()

    Path(out_pdf).parent.mkdir(parents=True, exist_ok=True)
    if os.path.exists(out_pdf):
        print(f"ACROBAT_SAVE_REMOVE_EXISTING {out_pdf}")
        os.remove(out_pdf)

    try:
        active_avdoc, active_pddoc = _get_active_pddoc(app)
    except Exception:
        if os.path.exists(src_pdf) and os.path.getsize(src_pdf) > 1024 and _source_has_text_layer(src_pdf):
            print(f"ACROBAT_SAVE_ACTIVE_DOC_UNAVAILABLE_COPY_OPEN_FILE {src_pdf} -> {out_pdf}")
            shutil.copy2(src_pdf, out_pdf)
            return
        print("ACROBAT_SAVE_SOURCE_NOT_READY_FOR_COPY")
        _print_save_diagnostics(src_pdf, out_pdf, [])
        raise

    ok = False
    dialogs_before = _handle_save_dialogs()
    try:
        ok = bool(active_pddoc.Save(1, out_pdf))
    except Exception as exc:
        print(f"PDDoc.Save SaveFull exception: {exc}")
    print(f"PDDoc.Save SaveFull result: {ok}")
    dialogs_after_full = _handle_save_dialogs()

    if not ok:
        try:
            ok = bool(active_pddoc.Save(0, out_pdf))
        except Exception as exc:
            print(f"PDDoc.Save SaveCopy fallback exception: {exc}")
        print(f"PDDoc.Save SaveCopy fallback result: {ok}")
        dialogs_after_copy = _handle_save_dialogs()
    else:
        dialogs_after_copy = []

    if not ok and os.path.exists(src_pdf) and os.path.getsize(src_pdf) > 1024 and _source_has_text_layer(src_pdf):
        print(f"ACROBAT_SAVE_COPY_OPEN_FILE {src_pdf} -> {out_pdf}")
        shutil.copy2(src_pdf, out_pdf)
        ok = True

    if not ok:
        _print_save_diagnostics(src_pdf, out_pdf, dialogs_before + dialogs_after_full + dialogs_after_copy)
        raise RuntimeError(f"Acrobat failed to save PDF: {out_pdf}")

    try:
        active_avdoc.Close(False)
    except Exception:
        pass


def img_pdf_to_double_layer(
    src_img_pdf: str,
    out_double_pdf: str,
    ocr_lang: str = "CHS",
    timeout: int = 300,
) -> bool:
    app = None
    av_doc = None
    success = False
    src_img_pdf = os.path.abspath(src_img_pdf)
    out_double_pdf = os.path.abspath(out_double_pdf)
    _kill_acrobat()

    try:
        if not check_pdf_valid(src_img_pdf):
            print(f"[ERROR] Source PDF is missing or invalid: {src_img_pdf}")
            return False

        print("[4/5] Opening image-only PDF in Acrobat desktop UI...")
        app, av_doc, _ = _open_pdf_visible(src_img_pdf)
        print("OPEN_RESULT True")

        print("[4/5] Running Acrobat OCR through desktop UI automation...")
        ui_ocr_result = run_ocr_by_ui(timeout=timeout)
        print(f"UI_OCR_RESULT {ui_ocr_result}")

        print("[4/5] Saving Acrobat searchable-image PDF...")
        av_doc = None
        try:
            _save_pdf_after_ocr(app, src_img_pdf, out_double_pdf)
            save_result = True
        except Exception:
            print("SAVE_RESULT False")
            print("FAIL_STAGE=save")
            raise
        print(f"SAVE_RESULT {save_result}")

        verify_result = verify_dual_pdf(out_double_pdf, before_pdf=src_img_pdf)
        print(f"VERIFY_RESULT {verify_result}")
        if not verify_result:
            print("FAIL_STAGE=verify")
            raise RuntimeError(f"Saved PDF failed verification: {out_double_pdf}")
        print(f"[SUCCESS] Output complete: {out_double_pdf}")
        success = True
        return True
    except AcrobatUiOcrError:
        print("UI_OCR_RESULT False")
        print("FAIL_STAGE=ui_ocr")
        traceback.print_exc()
        return False
    except Exception:
        traceback.print_exc()
        return False
    finally:
        if success:
            if av_doc:
                try:
                    av_doc.Close(True)
                except Exception:
                    pass
            if app:
                try:
                    app.Exit()
                except Exception:
                    pass
            _kill_acrobat()
        else:
            print("ACROBAT_DEBUG_KEEP_OPEN True")


def batch_convert_double_layer(pdf_source_list: list, output_dir: str, ocr_lang: str = "CHS") -> bool:
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    all_ok = True

    for src_path in pdf_source_list:
        src_path = os.path.abspath(src_path)
        if not os.path.isfile(src_path):
            print(f"[SKIP] File does not exist: {src_path}")
            all_ok = False
            continue

        out_path = os.path.join(output_dir, os.path.basename(src_path))
        if not img_pdf_to_double_layer(src_path, out_path, ocr_lang=ocr_lang):
            all_ok = False
    return bool(all_ok)


def check_pdf_valid(pdf_path: str) -> bool:
    if not os.path.exists(pdf_path):
        return False
    if os.path.getsize(pdf_path) < 1024:
        return False
    return True


def main_workflow_demo() -> None:
    test_in = r"C:\temp\scan.pdf"
    test_out = r"C:\temp\dual.pdf"
    if not check_pdf_valid(test_in):
        raise RuntimeError("Source PDF is missing or invalid")
    img_pdf_to_double_layer(test_in, test_out, "CHS")


def main() -> int:
    parser = argparse.ArgumentParser(description="Acrobat UI-driven double-layer PDF conversion")
    parser.add_argument("--input", dest="input_pdf")
    parser.add_argument("--output", dest="output_pdf")
    parser.add_argument("--lang", default="CHS")
    parser.add_argument("--timeout", type=int, default=300)
    args = parser.parse_args()

    if not args.input_pdf and not args.output_pdf:
        main_workflow_demo()
        return 0
    if not args.input_pdf or not args.output_pdf:
        parser.error("--input and --output must be provided together")
    return 0 if img_pdf_to_double_layer(args.input_pdf, args.output_pdf, args.lang, args.timeout) else 1


if __name__ == "__main__":
    raise SystemExit(main())
