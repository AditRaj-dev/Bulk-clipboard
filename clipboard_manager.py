import pyperclip
import keyboard


def copy(text):
    pyperclip.copy(text)


def paste():
    return pyperclip.paste()


def on_paste(callback):
    keyboard.add_hotkey("ctrl+v", callback)


def trigger_ahk(hotkey: str):
    try:
        keyboard.press_and_release(hotkey)
    except Exception:
        # swallow exceptions; caller may log
        pass
