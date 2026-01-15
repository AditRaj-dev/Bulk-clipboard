import time
import threading
import state
from excel_manager import (
    append_text,
    read_row,
    get_cursor,
    set_cursor,
    clear_cursor,
)
from clipboard_manager import copy, paste
from logger import log


def start_capture(sheet):
    state.capture_active = True

    def loop():
        last = paste()
        while state.capture_active:
            text = paste()
            if text and text != last:
                row = append_text(sheet, text)
                log(f"Saved clipboard → {sheet} row {row}")
                last = text
            time.sleep(0.5)

    threading.Thread(target=loop, daemon=True).start()
    log("Capture started")


def stop_capture():
    state.capture_active = False
    log("Capture stopped")


def start_loading(sheet, start_row):
    state.loading_active = True
    clear_cursor(sheet)
    set_cursor(sheet, start_row)
    _load(sheet, start_row)
    log("Loading started")


def stop_loading(sheet):
    state.loading_active = False
    state.waiting_for_paste = False
    clear_cursor(sheet)
    log("Loading stopped")


def _load(sheet, start_row):
    row = get_cursor(sheet, start_row)
    value = read_row(sheet, row)
    if not value:
        stop_loading(sheet)
        return
    copy(value)
    state.waiting_for_paste = True
    log(f"Loaded row {row} → clipboard")


def handle_paste(sheet, start_row):
    if not state.loading_active or not state.waiting_for_paste:
        return
    row = get_cursor(sheet, start_row) + 1
    set_cursor(sheet, row)
    state.waiting_for_paste = False
    log(f"Paste detected → moved to row {row}")
    _load(sheet, start_row)


def calculate_question_count(start_q: int, end_q: int) -> int:
    if start_q > end_q:
        return 0
    return end_q - start_q + 1


def calculate_loop_count(start_q: int, end_q: int, from_q: int) -> int:
    if from_q < start_q:
        from_q = start_q
    if from_q > end_q:
        raise ValueError("Start-from question exceeds last question")
    return end_q - from_q + 1


def question_to_row(question_no: int, base_question: int) -> int:
    """Convert an absolute question number to an Excel row index.

    Raises ValueError if question is before base.
    """
    row = (question_no - base_question) + 1
    if row < 1:
        raise ValueError("Question number is before base question")
    return row


def run_submit_ahk_for_sheet(sheet: str, start_q: int, end_q: int, from_q: int, base_q: int):
    # validate and compute start row and loop count, then start loading
    try:
        start_row = question_to_row(from_q, base_q)
    except Exception as e:
        log(f"ERROR computing start row: {e}")
        return

    try:
        loop_count = calculate_loop_count(from_q, end_q, from_q)
    except Exception as e:
        log(f"ERROR calculating loop count: {e}")
        return

    # start loading on the given sheet at the computed start_row
    try:
        start_loading(sheet, start_row)
    except Exception as e:
        log(f"ERROR starting loading on sheet {sheet}: {e}")
        return
