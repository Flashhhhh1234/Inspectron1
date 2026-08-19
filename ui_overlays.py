import tkinter as tk


class PopupKeyboard:
    def __init__(self, root):
        self.root = root
        self.window = None
        self._target = None
        self._last_value = ""

    def attach(self):
        if getattr(self.root, "_popup_keyboard_attached", False):
            return
        self.root._popup_keyboard_attached = True
        self.root.bind_all("<FocusIn>", self._maybe_show, add="+")
        self.root.bind_all("<Button-1>", self._maybe_hide, add="+")

    def _maybe_show(self, event):
        widget = getattr(event, "widget", None)
        if not isinstance(widget, (tk.Entry, tk.Text, tk.Spinbox)):
            return
        self.show(widget)

    def _maybe_hide(self, event):
        widget = getattr(event, "widget", None)
        if widget is None or widget == self.window:
            return
        if isinstance(widget, (tk.Entry, tk.Text, tk.Spinbox)):
            return
        self.hide()

    def show(self, widget):
        self._target = widget
        if self.window and self.window.winfo_exists():
            self.window.lift()
            self.window.focus_force()
            return

        self.window = tk.Toplevel(self.root)
        self.window.title("Keyboard")
        self.window.configure(bg="#0f172a")
        self.window.attributes("-topmost", True)
        self.window.resizable(False, False)
        self.window.geometry("760x260")
        self.window.protocol("WM_DELETE_WINDOW", self.hide)

        header = tk.Frame(self.window, bg="#0f172a")
        header.pack(fill=tk.X, padx=10, pady=(10, 6))
        tk.Label(header, text="On-Screen Keyboard", bg="#0f172a", fg="white", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)
        tk.Button(header, text="×", command=self.hide, bg="#1f2937", fg="white", relief=tk.FLAT, padx=10, pady=2).pack(side=tk.RIGHT)

        body = tk.Frame(self.window, bg="#111827")
        body.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        rows = [
            ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "Backspace"],
            ["q", "w", "e", "r", "t", "y", "u", "i", "o", "p"],
            ["a", "s", "d", "f", "g", "h", "j", "k", "l", "Enter"],
            ["z", "x", "c", "v", "b", "n", "m", ",", ".", "?", "Space"],
        ]

        for row in rows:
            row_frame = tk.Frame(body, bg="#111827")
            row_frame.pack(fill=tk.X, pady=4)
            for key in row:
                width = 8 if len(key) == 1 else 10
                if key == "Space":
                    width = 24
                tk.Button(
                    row_frame,
                    text=key,
                    width=width,
                    command=lambda value=key: self._press(value),
                    bg="#1f2937",
                    fg="white",
                    relief=tk.FLAT,
                    activebackground="#334155",
                    activeforeground="white",
                    cursor="hand2",
                ).pack(side=tk.LEFT, padx=3)

    def _press(self, value):
        target = self._target
        if target is None:
            return
        if value == "Backspace":
            target.event_generate("<BackSpace>")
        elif value == "Enter":
            target.event_generate("<Return>")
        elif value == "Space":
            target.insert(tk.INSERT, " ")
        else:
            target.insert(tk.INSERT, value)

    def hide(self):
        if self.window and self.window.winfo_exists():
            self.window.destroy()
        self.window = None
        self._target = None


class ZoomAdjusterPopup:
    def __init__(self, root, current_zoom, on_change, on_close=None):
        self.root = root
        self.on_change = on_change
        self.on_close = on_close
        self.window = tk.Toplevel(root)
        self.window.title("Zoom Adjuster")
        self.window.configure(bg="#111827")
        self.window.attributes("-topmost", True)
        self.window.resizable(False, False)
        self.window.geometry("420x140")
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        top = tk.Frame(self.window, bg="#111827")
        top.pack(fill=tk.X, padx=12, pady=(10, 0))
        tk.Label(top, text="Zoom Level", bg="#111827", fg="white", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)
        tk.Button(top, text="×", command=self.close, bg="#1f2937", fg="white", relief=tk.FLAT, padx=10, pady=2).pack(side=tk.RIGHT)

        self.value_label = tk.Label(self.window, text="", bg="#111827", fg="#93c5fd", font=("Segoe UI", 11, "bold"))
        self.value_label.pack(pady=(10, 6))

        self.scale = tk.Scale(
            self.window,
            from_=100,
            to=200,
            orient=tk.HORIZONTAL,
            showvalue=False,
            resolution=1,
            length=360,
            bg="#111827",
            fg="white",
            troughcolor="#334155",
            highlightthickness=0,
            command=self._update,
        )
        self.scale.pack(padx=12, pady=(0, 12))
        self.scale.set(int(round(current_zoom * 100)))
        self._update(self.scale.get())

    def _update(self, value):
        percent = int(float(value))
        self.value_label.config(text=f"{percent}%")
        if self.on_change:
            self.on_change(percent / 100.0)

    def close(self):
        if self.on_close:
            self.on_close()
        if self.window and self.window.winfo_exists():
            self.window.destroy()
