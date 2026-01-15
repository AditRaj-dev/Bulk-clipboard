from datetime import datetime

_log_widget = None


def set_logger(widget):
    global _log_widget
    _log_widget = widget


def log(message):
    if not _log_widget:
        return
    ts = datetime.now().strftime("%H:%M:%S")
    _log_widget.insert("end", f"[{ts}] {message}\n")
    _log_widget.see("end")
