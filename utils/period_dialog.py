"""程序运行前的会计期间输入弹窗：留空或 60 秒无操作按自然月运行。"""
import re

from log.loggingConfig import logger


DEFAULT_TIMEOUT_SECONDS = 60


def _valid_period(text):
    match = re.fullmatch(r"(\d{4})[-/年]?(\d{1,2})月?", text)
    if not match:
        return False
    return 1 <= int(match.group(2)) <= 12


def ask_accounting_period(timeout=DEFAULT_TIMEOUT_SECONDS):
    """弹出会计期间输入框，返回输入值；留空、超时或无法弹窗时返回空字符串。"""
    try:
        return _show_dialog(timeout)
    except Exception as exc:
        logger.warning(f"无法显示会计期间弹窗，按自然月运行：{exc}")
        return ""


def _show_dialog(timeout):
    import tkinter as tk

    root = tk.Tk()
    root.title("会计期间")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    result = {"value": "", "closed": False}
    remaining = {"value": max(1, int(timeout))}
    timer = {"id": None}

    frame = tk.Frame(root, padx=20, pady=16)
    frame.pack(fill="both", expand=True)
    tk.Label(frame, text="请输入本次运行的会计期间（如 2026-10 或 202610）：").pack(anchor="w")
    entry = tk.Entry(frame, width=30)
    entry.pack(fill="x", pady=(8, 4))
    entry.focus_set()
    tk.Label(
        frame,
        text="留空直接确定，或 60 秒无操作自动关闭，按自然月运行",
        fg="#666666",
    ).pack(anchor="w")
    countdown = tk.Label(frame, text="", fg="#cc0000")
    countdown.pack(anchor="e", pady=(6, 0))
    error_label = tk.Label(frame, text="", fg="#cc0000")
    error_label.pack(anchor="w")

    def close(value=""):
        if result["closed"]:
            return
        result["closed"] = True
        result["value"] = value
        if timer["id"] is not None:
            root.after_cancel(timer["id"])
            timer["id"] = None
        root.destroy()

    def tick():
        if result["closed"]:
            return
        remaining["value"] -= 1
        if remaining["value"] <= 0:
            close("")
            return
        countdown.config(text=f"{remaining['value']} 秒无操作将自动关闭")
        timer["id"] = root.after(1000, tick)

    def reset(event=None):
        if result["closed"]:
            return
        remaining["value"] = max(1, int(timeout))
        countdown.config(text=f"{remaining['value']} 秒无操作将自动关闭")
        if timer["id"] is not None:
            root.after_cancel(timer["id"])
        timer["id"] = root.after(1000, tick)

    def confirm(event=None):
        text = entry.get().strip()
        if text and not _valid_period(text):
            error_label.config(text="格式不正确，请输入 YYYY-MM 或 YYYYMM，例如 2026-10")
            reset()
            return
        close(text)

    buttons = tk.Frame(frame)
    buttons.pack(pady=(10, 0))
    tk.Button(buttons, text="确定", width=10, command=confirm).pack()

    root.bind("<KeyPress>", reset)
    root.bind("<ButtonPress>", reset)
    root.bind("<Return>", confirm)
    root.bind("<Escape>", lambda event: close(""))
    root.protocol("WM_DELETE_WINDOW", lambda: close(""))

    root.update_idletasks()
    width, height = root.winfo_width(), root.winfo_height()
    x = (root.winfo_screenwidth() - width) // 2
    y = (root.winfo_screenheight() - height) // 2
    root.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    reset()
    entry.focus_set()
    root.mainloop()
    return result["value"]
