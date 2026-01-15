import tkinter as tk
from ui import build_ui

root = tk.Tk()
root.title("Excel Clipboard Automation")
root.geometry("1000x650")

build_ui(root)

root.mainloop()
