import argparse
import json
import time
from pathlib import Path
from typing import Optional


class AcrobatUiOcrError(RuntimeError):
    pass


ROOT = Path(__file__).resolve().parent
LOG_DIR = ROOT / "logs"
COORDS_PATH = ROOT / "ui_coords.json"


def _require_pywinauto():
    try:
        from pywinauto import Application, Desktop, keyboard, mouse
    except ImportError as exc:
        raise AcrobatUiOcrError(
            "pywinauto is not installed. Run install-python-packages.ps1 or pip install pywinauto."
        ) from exc
    return Application, Desktop, keyboard, mouse


def _load_coords() -> dict:
    with COORDS_PATH.open("r", encoding="utf-8") as handle:
        config = json.load(handle)
    if not config.get("manual_ui_mode"):
        raise AcrobatUiOcrError("ui_coords.json manual_ui_mode must be true for Acrobat UI OCR.")
    return config


def _connect_window(timeout: int = 30):
    Application, Desktop, _, _ = _require_pywinauto()
    deadline = time.time() + timeout
    last_error: Optional[Exception] = None

    while time.time() < deadline:
        try:
            candidates = Desktop(backend="uia").windows(
                title_re=".*(Adobe Acrobat|Acrobat).*",
                visible_only=True,
                enabled_only=False,
            )
            print(f"ACROBAT_UI_WINDOW_CANDIDATES {len(candidates)}")
            for candidate in candidates:
                try:
                    print(f"ACROBAT_UI_WINDOW_CANDIDATE {candidate.window_text() or ''}")
                except Exception:
                    pass

            pdf_windows = []
            other_windows = []
            for candidate in candidates:
                try:
                    title = candidate.window_text() or ""
                    if ".pdf" in title.lower():
                        pdf_windows.append(candidate)
                    elif title.strip() not in {"Adobe Acrobat Pro (64-bit)", "Adobe Acrobat (64-bit)", "Adobe Acrobat"}:
                        other_windows.append(candidate)
                except Exception as exc:
                    last_error = exc

            selectable = pdf_windows or other_windows or candidates
            if not selectable:
                raise AcrobatUiOcrError("No visible Acrobat window found.")
            win = selectable[0]
            if hasattr(win, "wait"):
                win.wait("exists visible ready", timeout=3)
            win.set_focus()
            app = Application(backend="uia").connect(handle=win.handle)
            return app, win
        except Exception as exc:
            last_error = exc
            time.sleep(1)

    raise AcrobatUiOcrError(f"Could not connect to Acrobat window: {last_error}")


def _screenshot(win, filename: str) -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    path = LOG_DIR / filename
    try:
        win.capture_as_image().save(path)
        print(f"ACROBAT_UI_SCREENSHOT {path.resolve()}")
    except Exception as exc:
        print(f"ACROBAT_UI_SCREENSHOT_FAILED {filename} {exc}")


def _maximize_window(win) -> None:
    win.set_focus()
    try:
        win.maximize()
    except Exception:
        try:
            win.type_keys("% x")
        except Exception:
            pass
    time.sleep(2)
    try:
        rect = win.rectangle()
        print(
            "ACROBAT_WINDOW_RECT "
            f"left={rect.left} top={rect.top} right={rect.right} bottom={rect.bottom} "
            f"width={rect.width()} height={rect.height()}"
        )
    except Exception as exc:
        print(f"ACROBAT_WINDOW_RECT_FAILED {exc}")
    _screenshot(win, "window_maximized.png")


def _check_resolution(config: dict) -> None:
    try:
        import ctypes

        width = ctypes.windll.user32.GetSystemMetrics(0)
        height = ctypes.windll.user32.GetSystemMetrics(1)
        expected = config.get("expected_resolution") or []
        print(f"ACROBAT_SCREEN_RESOLUTION {width}x{height}")
        if expected and [width, height] != expected:
            print(f"ACROBAT_SCREEN_RESOLUTION_WARNING expected={expected} actual={[width, height]}")
    except Exception as exc:
        print(f"ACROBAT_SCREEN_RESOLUTION_CHECK_FAILED {exc}")


def _click_absolute(point: list[int], label: str) -> None:
    _, _, _, mouse = _require_pywinauto()
    x, y = int(point[0]), int(point[1])
    print(f"ACROBAT_UI_CLICK_ABSOLUTE {label} {x},{y}")
    mouse.click(button="left", coords=(x, y))


def _click_confirm_button(point: list[int]) -> None:
    _, _, keyboard, mouse = _require_pywinauto()
    x, y = int(point[0]), int(point[1])
    print(f"ACROBAT_UI_CLICK_CONFIRM {x},{y}")
    mouse.click(button="left", coords=(x, y))
    time.sleep(0.5)
    mouse.click(button="left", coords=(x, y))
    time.sleep(0.5)
    keyboard.send_keys("{ENTER}")


def _visible_text(win) -> str:
    try:
        texts = []
        for control in win.descendants():
            text = ""
            try:
                text = control.window_text()
            except Exception:
                pass
            if not text:
                try:
                    text = control.element_info.name or ""
                except Exception:
                    text = ""
            if text:
                texts.append(text)
        return "\n".join(texts)
    except Exception:
        return ""


def _recognize_button_still_visible(win) -> bool:
    visible_text = _visible_text(win)
    return "Recognize Text" in visible_text or "\u8bc6\u522b\u6587\u672c" in visible_text


def _ocr_active_markers() -> tuple[str, ...]:
    return (
        "Recognizing",
        "Performing page recognition",
        "Scanning and Text Recognition",
        "Cancel",
        "\u6b63\u5728\u8bc6\u522b",
        "\u6b63\u5728\u6267\u884c\u9875\u9762\u8bc6\u522b",
        "\u6b63\u5728\u626b\u63cf",
        "\u5c06\u626b\u63cf\u7684\u9875\u9762\u8f6c\u6362\u4e3a\u53ef\u641c\u7d22\u7684\u56fe\u50cf",
        "\u8bf7\u7a0d\u5019",
        "\u53d6\u6d88",
    )


def _ocr_is_active(win) -> bool:
    visible_text = _visible_text(win)
    return any(marker in visible_text for marker in _ocr_active_markers())


def _wait_ocr_started(win, timeout: int = 20) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _ocr_is_active(win):
            print("ACROBAT_UI_OCR_STARTED True")
            return True
        visible_text = _visible_text(win)
        if "Recognize Text" not in visible_text and "\u8bc6\u522b\u6587\u672c" not in visible_text:
            print("ACROBAT_UI_OCR_STARTED True button_disappeared")
            return True
        time.sleep(1)
    print("ACROBAT_UI_OCR_STARTED False")
    return False


def _wait_ocr_complete(win, timeout: int = 300) -> bool:
    deadline = time.time() + timeout
    started_at = time.time()
    last_report = started_at
    stable_rounds = 0
    saw_active = False

    while time.time() < deadline:
        if _ocr_is_active(win):
            saw_active = True
            stable_rounds = 0
        else:
            stable_rounds += 1
            required_stable_rounds = 6 if saw_active else 15
            if stable_rounds >= required_stable_rounds:
                print(f"ACROBAT_UI_OCR_COMPLETE detected elapsed={int(time.time() - started_at)}s")
                return True
        if time.time() - last_report >= 60:
            print(
                "ACROBAT_UI_OCR_WAITING "
                f"elapsed={int(time.time() - started_at)}s timeout={timeout}s active_seen={saw_active}"
            )
            last_report = time.time()
        time.sleep(2)

    print(f"ACROBAT_UI_OCR_WAIT_TIMEOUT elapsed={int(time.time() - started_at)}s timeout={timeout}s")
    return False


def run_ocr_by_ui(timeout: int = 300, print_tree: bool = False) -> bool:
    config = _load_coords()
    _, win = _connect_window(timeout=30)
    title = win.window_text() or ""
    print(f"ACROBAT_UI_WINDOW_TITLE {title}")
    if title.strip() in {"Adobe Acrobat Pro (64-bit)", "Adobe Acrobat (64-bit)", "Adobe Acrobat"}:
        raise AcrobatUiOcrError("Acrobat window is open, but no PDF document is loaded.")
    if print_tree:
        win.print_control_identifiers()

    _check_resolution(config)
    if config.get("require_maximized"):
        _maximize_window(win)

    if _ocr_is_active(win):
        print("ACROBAT_UI_OCR_ALREADY_RUNNING wait_only")
        if not _wait_ocr_complete(win, timeout=timeout):
            raise AcrobatUiOcrError(f"Timed out waiting for Acrobat OCR to finish after {timeout} seconds")
        print("ACROBAT_UI_OCR_DONE")
        return True

    print("ACROBAT_UI_STEP scan_tool")
    _click_absolute(config["scan_tool"], "scan_tool")
    time.sleep(2)
    _screenshot(win, "step1_scan_tool.png")

    print("ACROBAT_UI_STEP recognize_text_menu")
    _click_absolute(config["recognize_text_menu"], "recognize_text_menu")
    time.sleep(1)
    _screenshot(win, "step2_recognize_menu.png")

    print("ACROBAT_UI_STEP in_this_file")
    _click_absolute(config["in_this_file"], "in_this_file")
    time.sleep(2)
    _screenshot(win, "step3_in_this_file.png")

    print("ACROBAT_UI_STEP recognize_text_confirm")
    _click_confirm_button(config["recognize_text_confirm"])
    time.sleep(3)
    _screenshot(win, "step4_recognize_confirm.png")

    if not _wait_ocr_started(win, timeout=20):
        print(f"ACROBAT_UI_CONFIRM_BUTTON_STILL_VISIBLE {_recognize_button_still_visible(win)}")
        print("FAIL_STAGE=recognize_confirm")
        raise AcrobatUiOcrError("Recognize Text confirm button was clicked, but Acrobat OCR did not start.")

    if not _wait_ocr_complete(win, timeout=timeout):
        raise AcrobatUiOcrError(f"Timed out waiting for Acrobat OCR to finish after {timeout} seconds")

    print("ACROBAT_UI_OCR_DONE")
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Run Acrobat OCR through desktop UI automation.")
    parser.add_argument("--timeout", type=int, default=300)
    parser.add_argument("--print-tree", action="store_true")
    args = parser.parse_args()
    return 0 if run_ocr_by_ui(timeout=args.timeout, print_tree=args.print_tree) else 1


if __name__ == "__main__":
    raise SystemExit(main())
