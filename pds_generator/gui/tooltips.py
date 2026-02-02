import tkinter as tk


class Tooltip:
    def __init__(self, root, delay=400, wraplength=360):
        self.root = root
        self.delay = delay
        self.wraplength = wraplength
        self._after_id = None
        self._text_func = None
        self._last_xy = None
        self._tip = None
        self._label = None

    def bind(self, widget, text=None, text_func=None):
        if text is None and text_func is None:
            return

        def get_text():
            if text_func is not None:
                try:
                    return text_func()
                except Exception:
                    return ""
            return text or ""

        def on_enter(event):
            self.schedule(get_text, event.x_root, event.y_root)

        def on_leave(_event):
            self.hide()

        def on_motion(event):
            self._last_xy = (event.x_root, event.y_root)
            if self._tip is not None:
                self._move(event.x_root, event.y_root)

        widget.bind("<Enter>", on_enter, add="+")
        widget.bind("<Leave>", on_leave, add="+")
        widget.bind("<Motion>", on_motion, add="+")

    def bind_canvas(self, canvas, item, text_func):
        if text_func is None:
            return

        def on_enter(event):
            self.schedule(text_func, event.x_root, event.y_root)

        def on_leave(_event):
            self.hide()

        def on_motion(event):
            self._last_xy = (event.x_root, event.y_root)
            if self._tip is not None:
                self._move(event.x_root, event.y_root)

        canvas.tag_bind(item, "<Enter>", on_enter)
        canvas.tag_bind(item, "<Leave>", on_leave)
        canvas.tag_bind(item, "<Motion>", on_motion)

    def schedule(self, text_func, x_root, y_root):
        self.cancel()
        self._text_func = text_func
        self._last_xy = (x_root, y_root)
        self._after_id = self.root.after(self.delay, self._show_scheduled)

    def cancel(self):
        if self._after_id is not None:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show_scheduled(self):
        self._after_id = None
        if self._text_func is None:
            return
        text = self._text_func() or ""
        if not text:
            return
        x_root, y_root = self._last_xy or (0, 0)
        self._show(text, x_root, y_root)

    def _show(self, text, x_root, y_root):
        if self._tip is None:
            self._tip = tk.Toplevel(self.root)
            self._tip.wm_overrideredirect(True)
            try:
                self._tip.attributes("-topmost", True)
            except Exception:
                pass
            self._label = tk.Label(
                self._tip,
                text=text,
                justify="left",
                background="#ffffe0",
                relief="solid",
                borderwidth=1,
                wraplength=self.wraplength,
                padx=6,
                pady=4,
            )
            self._label.pack()
        else:
            self._label.configure(text=text)
            self._tip.deiconify()
        self._move(x_root, y_root)

    def _move(self, x_root, y_root):
        if self._tip is None:
            return
        x = x_root + 12
        y = y_root + 16
        self._tip.wm_geometry(f"+{x}+{y}")

    def hide(self):
        self.cancel()
        if self._tip is not None:
            try:
                self._tip.withdraw()
            except Exception:
                pass
