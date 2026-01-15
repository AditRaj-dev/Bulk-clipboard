import tkinter as tk
from config import *
import automation
import clipboard_manager
import logger
import ahk_manager


def build_ui(root):
    root.configure(bg=BG)

    # ================== VARIABLES ==================
    sheet_var = tk.StringVar(value=SHEETS[0])
    start_row_var = tk.StringVar(value="1")
    current_row_var = tk.StringVar(value="-")

    start_q_var = tk.StringVar()
    end_q_var = tk.StringVar()
    from_q_var = tk.StringVar()
    base_q_var = tk.StringVar()
    total_q_var = tk.StringVar(value="0")
    computed_row_var = tk.StringVar(value="-")
    ahk_loop_var = tk.StringVar(value="1")

    advanced_visible = tk.BooleanVar(value=False)
    ahk_visible = tk.BooleanVar(value=False)

    # ================== HELPERS ==================
    def log(msg):
        logger.log(msg)

    def set_start_row():
        current_row_var.set(start_row_var.get())
        log(f"Start row set to {start_row_var.get()}")

    def toggle_advanced():
        if advanced_visible.get():
            advanced_frame.pack_forget()
        else:
            advanced_frame.pack(fill=tk.X, padx=10, pady=5)
        advanced_visible.set(not advanced_visible.get())

    def toggle_ahk():
        if ahk_visible.get():
            ahk_frame.pack_forget()
        else:
            ahk_frame.pack(fill=tk.X, padx=10, pady=5)
        ahk_visible.set(not ahk_visible.get())

    def update_question_calcs():
        try:
            s = int(start_q_var.get())
            e = int(end_q_var.get())
            f = int(from_q_var.get())
            b = int(base_q_var.get())

            total_q_var.set(str(e - s + 1 if e >= s else 0))
            ahk_loop_var.set(str(automation.calculate_loop_count(s, e, f)))
            computed_row_var.set(str(automation.question_to_row(f, b)))
        except Exception:
            total_q_var.set("0")
            computed_row_var.set("-")

    def start_loading():
        try:
            if advanced_visible.get():
                base_q = int(base_q_var.get())
                from_q = int(from_q_var.get())
                start_row = automation.question_to_row(from_q, base_q)
                automation.start_loading(sheet_var.get(), start_row)
            else:
                automation.start_loading(sheet_var.get(), int(start_row_var.get()))
        except Exception as e:
            log(f"ERROR: {e}")

    # ================== CONTEXT ==================
    context = tk.LabelFrame(root, text="Context", bg=PANEL, fg=FG)
    context.pack(fill=tk.X, padx=10, pady=5)

    tk.Label(context, text="Sheet", bg=PANEL, fg=FG).grid(row=0, column=0, padx=5)
    tk.OptionMenu(context, sheet_var, *SHEETS).grid(row=0, column=1, padx=5)

    tk.Label(context, text="Current Row", bg=PANEL, fg=FG).grid(row=0, column=2, padx=10)
    tk.Entry(context, textvariable=current_row_var, width=6,
             state="readonly").grid(row=0, column=3)

    # ================== SIMPLE LOADING ==================
    simple = tk.LabelFrame(root, text="Simple Loading (Row-based)", bg=PANEL, fg=FG)
    simple.pack(fill=tk.X, padx=10, pady=5)

    tk.Label(simple, text="Start Row", bg=PANEL, fg=FG).grid(row=0, column=0, padx=5)
    tk.Entry(simple, textvariable=start_row_var, width=6).grid(row=0, column=1)
    tk.Button(simple, text="Set", bg=BTN, fg=FG,
              command=set_start_row).grid(row=0, column=2, padx=5)

    tk.Button(simple, text="Advanced Question Mode ▸", bg=BTN, fg=FG,
              command=toggle_advanced).grid(row=0, column=3, padx=20)

    # ================== ADVANCED LOADING ==================
    advanced_frame = tk.LabelFrame(root, text="Advanced Loading (Question-based)", bg=PANEL, fg=FG)

    tk.Label(advanced_frame, text="Start Q", bg=PANEL, fg=FG).grid(row=0, column=0)
    tk.Entry(advanced_frame, textvariable=start_q_var, width=6).grid(row=0, column=1)

    tk.Label(advanced_frame, text="End Q", bg=PANEL, fg=FG).grid(row=0, column=2)
    tk.Entry(advanced_frame, textvariable=end_q_var, width=6).grid(row=0, column=3)

    tk.Label(advanced_frame, text="From Q", bg=PANEL, fg=FG).grid(row=1, column=0)
    tk.Entry(advanced_frame, textvariable=from_q_var, width=6).grid(row=1, column=1)

    tk.Label(advanced_frame, text="Base Q", bg=PANEL, fg=FG).grid(row=1, column=2)
    tk.Entry(advanced_frame, textvariable=base_q_var, width=6).grid(row=1, column=3)

    tk.Label(advanced_frame, text="Total Q", bg=PANEL, fg=FG).grid(row=2, column=0)
    tk.Entry(advanced_frame, textvariable=total_q_var,
             width=6, state="readonly").grid(row=2, column=1)

    tk.Label(advanced_frame, text="Start Row (calc)", bg=PANEL, fg=FG).grid(row=2, column=2)
    tk.Entry(advanced_frame, textvariable=computed_row_var,
             width=6, state="readonly").grid(row=2, column=3)

    tk.Label(advanced_frame, text="AHK Loops", bg=PANEL, fg=FG).grid(row=3, column=0)
    tk.Entry(advanced_frame, textvariable=ahk_loop_var, width=6).grid(row=3, column=1)

    for var in (start_q_var, end_q_var, from_q_var, base_q_var):
        var.trace_add("write", lambda *_: update_question_calcs())

    # ================== ACTIONS ==================
    actions = tk.LabelFrame(root, text="Actions", bg=PANEL, fg=FG)
    actions.pack(fill=tk.X, padx=10, pady=5)

    tk.Button(actions, text="Start Capture", bg=BTN, fg=FG,
              command=lambda: automation.start_capture(sheet_var.get())).pack(side=tk.LEFT, padx=5)
    tk.Button(actions, text="Stop Capture", bg=BTN, fg=FG,
              command=automation.stop_capture).pack(side=tk.LEFT, padx=5)

    tk.Button(actions, text="Start Loading", bg=BTN, fg=FG,
              command=start_loading).pack(side=tk.LEFT, padx=20)
    tk.Button(actions, text="Stop Loading", bg=BTN, fg=FG,
              command=lambda: automation.stop_loading(sheet_var.get())).pack(side=tk.LEFT)

    tk.Button(actions, text="AHK Automation ▸", bg=BTN, fg=FG,
              command=toggle_ahk).pack(side=tk.RIGHT, padx=5)

    # ================== AHK ==================
    ahk_frame = tk.LabelFrame(root, text="AHK Automation", bg=PANEL, fg=FG)

    tk.Button(ahk_frame, text="Trigger Capture AHK", bg=BTN, fg=FG,
              command=lambda: clipboard_manager.trigger_ahk(AHK_CAPTURE_HOTKEY)).pack(side=tk.LEFT, padx=5)

    tk.Button(ahk_frame, text="Trigger Loading AHK", bg=BTN, fg=FG,
              command=lambda: clipboard_manager.trigger_ahk(AHK_LOADING_HOTKEY)).pack(side=tk.LEFT, padx=5)

    tk.Button(ahk_frame, text="Run Capture Script", bg=BTN, fg=FG,
              command=lambda: ahk_manager.run_ahk_file(CAPTURE_AHK_FILE)).pack(side=tk.LEFT, padx=10)

    tk.Button(ahk_frame, text="Run Loading Script", bg=BTN, fg=FG,
              command=lambda: ahk_manager.run_ahk_file(LOADING_AHK_FILE)).pack(side=tk.LEFT)

    # ================== LOGS ==================
    logs = tk.LabelFrame(root, text="Logs", bg=PANEL, fg=FG)
    logs.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    log_box = tk.Text(logs, bg=BG, fg=FG, height=10)
    log_box.pack(fill=tk.BOTH, expand=True)
    logger.set_logger(log_box)

    tk.Button(logs, text="Clear Logs", bg=BTN, fg=FG,
              command=lambda: (log_box.delete("1.0", tk.END),
                               logger.log("Logs cleared"))).pack(anchor="e", pady=3)

    # ================== PASTE HOOK ==================
    clipboard_manager.on_paste(
        lambda: automation.handle_paste(sheet_var.get(), int(start_row_var.get()))
    )
