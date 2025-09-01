import os
import sys
import json
import time
import threading
import subprocess
import importlib
import math
import re
from io import BytesIO
from types import SimpleNamespace

# Automatically ensure required third-party modules are available
REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "PIL": "Pillow",
    "reportlab": "reportlab",
    "requests": "requests",
    "openpyxl": "openpyxl",  # backend for pandas.read_excel
}


def ensure_dependencies():
    """Install missing dependencies using pip."""
    for module, package in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(module)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])


ensure_dependencies()

# Imports that rely on third-party packages
import pandas as pd
from PIL import Image, ImageTk
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
import requests

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter import font as tkfont

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")


def to_reportlab_color(value):
    try:
        return colors.HexColor(value)
    except Exception:
        return colors.toColor(value)


def sanitize_filename(name: str) -> str:
    """Return a filesystem-safe version of *name*."""
    cleaned = re.sub(r"[^\w\s-]", "", str(name))
    cleaned = re.sub(r"\s+", "_", cleaned).strip("_")
    return cleaned

# ---------------------------------------------------------------------------
# Model classes
# ---------------------------------------------------------------------------


class DraggableElement:
    """Representation of a draggable/resizable item on the configuration canvas."""

    HANDLE_SIZE = 8

    def __init__(self, parent, canvas: tk.Canvas, name: str, text: str):
        self.parent = parent
        self.canvas = canvas
        self.name = name
        self.text = text
        self.x = canvas.winfo_width() // 2 - 50
        self.y = canvas.winfo_height() // 2 - 20
        self.width = 100
        self.height = 40
        self.font_size = 12
        self.bold = False
        self.font_family = "Arial"
        self.auto_font = True
        self.text_color = "black"
        self.bg_color = "white"
        self.bg_visible = True
        self.align = "left"
        self._create_items()

    # ------------------------------------------------------------------
    def _create_items(self):
        self.rect = self.canvas.create_rectangle(
            self.x,
            self.y,
            self.x + self.width,
            self.y + self.height,
            fill=self.bg_color,
            outline="black",
        )
        self.label = self.canvas.create_text(0, 0, text=self.text, fill=self.text_color)
        self.handle = self.canvas.create_rectangle(
            self.x + self.width - self.HANDLE_SIZE,
            self.y + self.height - self.HANDLE_SIZE,
            self.x + self.width,
            self.y + self.height,
            fill="black",
        )

        # Bind events for dragging and resizing
        self.canvas.tag_bind(self.rect, "<ButtonPress-1>", self.start_move)
        self.canvas.tag_bind(self.rect, "<B1-Motion>", self.moving)
        self.canvas.tag_bind(self.rect, "<ButtonRelease-1>", self.stop_move)
        self.canvas.tag_bind(self.label, "<ButtonPress-1>", self.start_move)
        self.canvas.tag_bind(self.label, "<B1-Motion>", self.moving)
        self.canvas.tag_bind(self.label, "<ButtonRelease-1>", self.stop_move)
        self.canvas.tag_bind(self.handle, "<ButtonPress-1>", self.start_resize)
        self.canvas.tag_bind(self.handle, "<B1-Motion>", self.resizing)
        self.canvas.tag_bind(self.handle, "<ButtonRelease-1>", self.stop_resize)
        # Context menu for layering
        self.menu = tk.Menu(self.canvas, tearoff=0)
        self.menu.add_command(label="Przenieś na wierzch", command=self.bring_to_front)
        self.menu.add_command(label="Przenieś na spód", command=self.send_to_back)
        self.canvas.tag_bind(self.rect, "<Button-3>", self.show_menu)
        self.canvas.tag_bind(self.label, "<Button-3>", self.show_menu)
        self.canvas.tag_bind(self.handle, "<Button-3>", self.show_menu)
        self.apply_font()
        self.fit_text()
        self._update_label_position()

    # ------------------------------------------------------------------
    def show_menu(self, event):
        self.menu.tk_popup(event.x_root, event.y_root)

    def bring_to_front(self):
        items = [self.rect, self.label, self.handle, getattr(self, "image_id", None)]
        for item in filter(None, items):
            self.canvas.tag_raise(item)

    def send_to_back(self):
        items = [self.rect, self.label, self.handle, getattr(self, "image_id", None)]
        for item in filter(None, items):
            self.canvas.tag_lower(item)

    # ------------------------------------------------------------------
    def start_move(self, event):
        additive = bool(event.state & 0x0001)
        if self in self.parent.selected_elements:
            additive = True
        self.parent.select_element(self, additive=additive)
        self.last_x = event.x
        self.last_y = event.y

    def moving(self, event):
        dx = event.x - self.last_x
        dy = event.y - self.last_y
        for el in self.parent.selected_elements:
            for item in (el.rect, el.label, el.handle, getattr(el, "image_id", None)):
                if item:
                    el.canvas.move(item, dx, dy)
            el.x += dx
            el.y += dy
        self.last_x = event.x
        self.last_y = event.y

    def stop_move(self, event):
        step = self.parent.snap_step
        for el in self.parent.selected_elements:
            # snap top-left corner to the grid with integer multiples to
            # avoid sub-pixel artefacts when adjacent blocks touch
            el.x = int(round(el.x / step)) * step
            el.y = int(round(el.y / step)) * step
            el.sync_canvas()
        self.parent.push_history()

    # ------------------------------------------------------------------
    def start_resize(self, event):
        self.parent.select_element(self)
        # remember starting mouse position and dimensions so the handle
        # follows the cursor smoothly without jumping to the handle corner
        self.start_w = self.width
        self.start_h = self.height
        self.start_x = event.x
        self.start_y = event.y

    def resizing(self, event):
        step = self.parent.snap_step
        dx = event.x - self.start_x
        dy = event.y - self.start_y
        self.width = max(step, self.start_w + dx)
        self.height = max(step, self.start_h + dy)
        self.sync_canvas()

    def stop_resize(self, event):
        step = self.parent.snap_step
        # normalise width/height so edges line up exactly on the grid
        self.width = max(step, int(round(self.width / step)) * step)
        self.height = max(step, int(round(self.height / step)) * step)
        self.sync_canvas()
        self.parent.push_history()

    # ------------------------------------------------------------------
    def to_dict(self):
        scale = self.parent.scale
        return {
            "name": self.name,
            "text": self.text,
            "x": int(round(self.x / scale)),
            "y": int(round(self.y / scale)),
            "width": int(round(self.width / scale)),
            "height": int(round(self.height / scale)),
            "font_size": int(round(self.font_size / scale)),
            "bold": self.bold,
            "auto_font": self.auto_font,
            "text_color": self.text_color,
            "bg_color": self.bg_color,
            "bg_visible": self.bg_visible,
            "align": self.align,
        }

    def sync_canvas(self):
        self.canvas.coords(
            self.rect,
            self.x,
            self.y,
            self.x + self.width,
            self.y + self.height,
        )
        if hasattr(self, "image_id") and hasattr(self, "raw_image"):
            resized = self.raw_image.resize((int(self.width), int(self.height)), Image.LANCZOS)
            self.image_obj = ImageTk.PhotoImage(resized)
            self.canvas.itemconfig(self.image_id, image=self.image_obj)
            self.canvas.coords(self.image_id, self.x, self.y)
        self._update_label_position()
        self.canvas.coords(
            self.handle,
            self.x + self.width - self.HANDLE_SIZE,
            self.y + self.height - self.HANDLE_SIZE,
            self.x + self.width,
            self.y + self.height,
        )
        self.apply_font()
        if self.auto_font:
            self.fit_text()
        self.update_colors()

    def update_value(self, value):
        """Update displayed value (text or image)."""
        # Remove previous image if any
        if hasattr(self, "image_id"):
            self.canvas.delete(self.image_id)
            del self.image_id
            if hasattr(self, "image_obj"):
                del self.image_obj
            if hasattr(self, "raw_image"):
                del self.raw_image
        try:
            if value is None or pd.isna(value):
                value = ""
        except Exception:
            if value is None:
                value = ""
        if isinstance(value, str) and value.lower().startswith("http"):
            try:
                resp = requests.get(value, timeout=5)
                self.raw_image = Image.open(BytesIO(resp.content))
                img = self.raw_image.resize((int(self.width), int(self.height)), Image.LANCZOS)
                self.image_obj = ImageTk.PhotoImage(img)
                self.image_id = self.canvas.create_image(
                    self.x,
                    self.y,
                    anchor="nw",
                    image=self.image_obj,
                )
                for tag in (self.image_id,):
                    self.canvas.tag_bind(tag, "<ButtonPress-1>", self.start_move)
                    self.canvas.tag_bind(tag, "<B1-Motion>", self.moving)
                    self.canvas.tag_bind(tag, "<ButtonRelease-1>", self.stop_move)
                    self.canvas.tag_bind(tag, "<Button-3>", self.show_menu)
                self.canvas.tag_raise(self.rect)
                self.canvas.tag_raise(self.handle)
                self.canvas.itemconfig(self.rect, fill="")
                self.canvas.itemconfig(self.label, text="", state="hidden")
                self.text = str(value)
                return
            except Exception:
                pass
        # default: text
        self.canvas.itemconfig(self.rect, fill=self.bg_color if self.bg_visible else "")
        self.canvas.itemconfig(
            self.label,
            text=str(value),
            fill=self.text_color,
            state="normal",
        )
        self.text = str(value)
        self.apply_font()
        if self.auto_font:
            self.fit_text()
        self._update_label_position()

    def apply_font(self):
        weight = "bold" if self.bold else "normal"
        self.canvas.itemconfig(self.label, font=(self.font_family, int(self.font_size), weight))

    def fit_text(self):
        if hasattr(self, "image_id") or not self.auto_font:
            return
        size = 1
        weight = "bold" if self.bold else "normal"
        test_font = tkfont.Font(family=self.font_family, size=size, weight=weight)
        while True:
            width = test_font.measure(self.text)
            height = test_font.metrics("linespace")
            if width > self.width - 4 or height > self.height - 4:
                break
            size += 1
            test_font.configure(size=size)
        self.font_size = max(1, size - 1)
        self.apply_font()

    def update_colors(self):
        if hasattr(self, "image_id"):
            self.canvas.itemconfig(self.rect, fill="")
        else:
            self.canvas.itemconfig(self.rect, fill=self.bg_color if self.bg_visible else "")
        self.canvas.itemconfig(self.label, fill=self.text_color)

    def _update_label_position(self):
        if self.align == "left":
            self.canvas.itemconfig(self.label, anchor="w")
            self.canvas.coords(self.label, self.x + 2, self.y + self.height / 2)
        elif self.align == "right":
            self.canvas.itemconfig(self.label, anchor="e")
            self.canvas.coords(self.label, self.x + self.width - 2, self.y + self.height / 2)
        else:
            self.canvas.itemconfig(self.label, anchor="center")
            self.canvas.coords(self.label, self.x + self.width / 2, self.y + self.height / 2)


# ---------------------------------------------------------------------------
# Group areas for automatic stacking
# ---------------------------------------------------------------------------


class GroupArea:
    """Semi-transparent rectangle grouping elements."""

    HANDLE_SIZE = 12

    def __init__(self, parent, canvas: tk.Canvas, name: str):
        self.parent = parent
        self.canvas = canvas
        self.name = name
        # Place the group roughly at the centre of the page rather than
        # relative to the widget size which could refer to a different canvas
        self.x = parent.page_width * parent.scale / 2 - 50
        self.y = parent.page_height * parent.scale / 2 - 50
        self.width = 100
        self.height = 100
        self.fields = []  # names of elements contained in this group
        self.field_pos = {}  # mapping name -> (x,y) inside the group
        self.field_conf = {}  # individual field styling inside the group
        self.conditions = []
        self.preview_items = []
        self.rect = canvas.create_rectangle(
            self.x,
            self.y,
            self.x + self.width,
            self.y + self.height,
            outline="blue",
            dash=(4, 2),
            # Tkinter doesn't support 8-digit hex colors; use stipple for translucency
            fill="#88aaff",
            stipple="gray50",
        )
        self.handle = canvas.create_rectangle(
            self.x + self.width - self.HANDLE_SIZE,
            self.y + self.height - self.HANDLE_SIZE,
            self.x + self.width,
            self.y + self.height,
            fill="black",
        )
        for tag in (self.rect,):
            canvas.tag_bind(tag, "<ButtonPress-1>", self.start_move)
            canvas.tag_bind(tag, "<B1-Motion>", self.moving)
            canvas.tag_bind(tag, "<ButtonRelease-1>", self.stop_move)
            canvas.tag_bind(tag, "<Double-1>", self.open_editor)
        canvas.tag_bind(self.handle, "<ButtonPress-1>", self.start_resize)
        canvas.tag_bind(self.handle, "<B1-Motion>", self.resizing)
        canvas.tag_bind(self.handle, "<ButtonRelease-1>", self.stop_resize)
        canvas.tag_bind(self.handle, "<Double-1>", self.open_editor)
        self.send_to_back()
        self.draw_preview()

    def send_to_back(self):
        self.canvas.tag_lower(self.rect)
        self.canvas.tag_lower(self.handle)
        # keep the group rectangle behind elements but ensure the handle is
        # always accessible above them
        self.canvas.tag_raise(self.rect, "grid")
        self.canvas.tag_raise(self.handle)

    def start_move(self, event):
        self.last_x = event.x
        self.last_y = event.y
        # capture elements currently inside so they move with the group
        self.children = [
            el
            for el in self.parent.elements.values()
            if self.parent.element_in_group(el, self)
        ]

    def moving(self, event):
        dx = event.x - self.last_x
        dy = event.y - self.last_y
        for item in (self.rect, self.handle):
            self.canvas.move(item, dx, dy)
        # move contained elements together with the group
        for el in self.children:
            for item in (el.rect, el.label, el.handle, getattr(el, "image_id", None)):
                if item:
                    self.canvas.move(item, dx, dy)
            el.x += dx
            el.y += dy
        self.x += dx
        self.y += dy
        self.last_x = event.x
        self.last_y = event.y

    def stop_move(self, event):
        step = self.parent.snap_step
        new_x = int(round(self.x / step)) * step
        new_y = int(round(self.y / step)) * step
        dx = new_x - self.x
        dy = new_y - self.y
        self.x = new_x
        self.y = new_y
        self.sync_canvas()
        # snap children by the same offset
        if dx or dy:
            for el in self.children:
                for item in (el.rect, el.label, el.handle, getattr(el, "image_id", None)):
                    if item:
                        self.canvas.move(item, dx, dy)
                el.x += dx
                el.y += dy

    def start_resize(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.start_w = self.width
        self.start_h = self.height

    def resizing(self, event):
        step = self.parent.snap_step
        dx = event.x - self.start_x
        dy = event.y - self.start_y
        self.width = max(step, self.start_w + dx)
        self.height = max(step, self.start_h + dy)
        self.sync_canvas()

    def stop_resize(self, event):
        step = self.parent.snap_step
        self.width = max(step, int(round(self.width / step)) * step)
        self.height = max(step, int(round(self.height / step)) * step)
        self.sync_canvas()

    def sync_canvas(self):
        self.canvas.coords(
            self.rect,
            self.x,
            self.y,
            self.x + self.width,
            self.y + self.height,
        )
        self.canvas.coords(
            self.handle,
            self.x + self.width - self.HANDLE_SIZE,
            self.y + self.height - self.HANDLE_SIZE,
            self.x + self.width,
            self.y + self.height,
        )
        self.send_to_back()
        self.draw_preview()

    def open_editor(self, event=None):
        if getattr(self, "editor", None) and self.editor.winfo_exists():
            self.editor.lift()
            self.editor.focus_force()
        else:
            self.editor = GroupEditor(self.parent, self)

    def to_dict(self):
        scale = self.parent.scale
        return {
            "name": self.name,
            "x": int(round(self.x / scale)),
            "y": int(round(self.y / scale)),
            "width": int(round(self.width / scale)),
            "height": int(round(self.height / scale)),
            "field_pos": {
                k: (int(round(v[0] / scale)), int(round(v[1] / scale)))
                for k, v in self.field_pos.items()
            },
            "field_conf": {
                k: {
                    "width": int(round(conf["width"] / scale)),
                    "height": int(round(conf["height"] / scale)),
                    "font_size": int(round(conf["font_size"] / scale)),
                    "bold": conf.get("bold", False),
                    "text_color": conf.get("text_color", "black"),
                    "bg_color": conf.get("bg_color", "white"),
                    "bg_visible": conf.get("bg_visible", True),
                    "align": conf.get("align", "left"),
                    "auto_font": conf.get("auto_font", True),
                }
                for k, conf in self.field_conf.items()
            },
            "conditions": list(self.conditions),
        }

    def draw_preview(self):
        for item in getattr(self, "preview_items", []):
            self.canvas.delete(item)
        self.preview_items = []
        if not self.fields:
            return
        # Build columns keyed by their x position
        cols = {}
        for name in self.fields:
            x, _y = self.field_pos.get(name, (0, 0))
            conf = self.field_conf.get(name, {})
            w = conf.get("width", 50)
            h = conf.get("height", 25)
            cols.setdefault(x, []).append((self.field_pos.get(name, (0, 0))[1], w, h, name))

        placed = []  # keep already positioned rectangles to detect collisions
        for x in sorted(cols):
            items = cols[x]
            items.sort(key=lambda t: t[0])
            cur_y = 0
            for _, w, h, name in items:
                y = cur_y
                # push down while colliding with any previously placed item
                while True:
                    overlap = False
                    for px, py, pw, ph in placed:
                        # Only treat as a collision when the candidate
                        # rectangle overlaps horizontally and vertically with
                        # a previously placed one. Ignoring blocks entirely
                        # below prevents reordering when a tall block exists
                        # underneath a smaller one.
                        if (
                            x < px + pw
                            and x + w > px
                            and y < py + ph
                            and y + h > py
                        ):
                            y = py + ph
                            overlap = True
                            break
                    if not overlap:
                        break
                if y + h > self.height:
                    continue
                x1 = self.x + x
                y1 = self.y + y
                r = self.canvas.create_rectangle(x1, y1, x1 + w, y1 + h, outline="blue")
                t = self.canvas.create_text(x1 + 2, y1 + h / 2, anchor="w", text=name)
                self.preview_items.extend([r, t])
                placed.append((x, y, w, h))
                cur_y = y + h
        self.send_to_back()



class GroupEditor(tk.Toplevel):
    """Editor window for configuring fields inside a group."""

    def __init__(self, parent, group: GroupArea):
        super().__init__(parent)
        self.parent = parent
        self.group = group
        self.title(group.name)
        self.scale = 1.0
        self.snap_step = parent.snap_step
        self.elements = {}
        self.selected_elements = []
        self.selected_element = None
        self.conditions = list(group.conditions)

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=5, pady=5)
        ttk.Button(toolbar, text="B", command=self.toggle_bold).pack(side="left")
        ttk.Button(toolbar, text="A+", command=self.increase_font).pack(side="left", padx=2)
        ttk.Button(toolbar, text="A-", command=self.decrease_font).pack(side="left")
        self.font_size_var = tk.StringVar()
        self.font_entry = ttk.Entry(toolbar, textvariable=self.font_size_var, width=4, state="disabled")
        self.font_entry.pack(side="left", padx=5)
        self.font_entry.bind("<Return>", lambda e: self.set_font_size())
        ttk.Button(toolbar, text="Kolor", command=self.choose_text_color).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Tło", command=self.choose_bg_color).pack(side="left", padx=2)
        self.transparent_var = tk.BooleanVar(value=False)
        self.bg_check = ttk.Checkbutton(toolbar, text="Przezroczyste", variable=self.transparent_var, command=self.toggle_bg_visible)
        self.bg_check.pack(side="left", padx=2)
        self.bg_check.state(["disabled"])
        ttk.Button(toolbar, text="L", command=lambda: self.set_alignment("left")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="C", command=lambda: self.set_alignment("center")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="R", command=lambda: self.set_alignment("right")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Warunki", command=self.open_conditions).pack(side="left", padx=5)

        main = ttk.Frame(self)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True)
        self.canvas = tk.Canvas(
            left,
            bg="white",
            width=group.width,
            height=group.height,
            scrollregion=(0, 0, group.width, group.height),
        )
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<ButtonPress-1>", self.canvas_button_press)
        self.canvas.bind("<B1-Motion>", self.canvas_drag_select)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_button_release)
        self.draw_grid()

        right = ttk.Frame(main)
        right.pack(side="right", fill="y", padx=5, pady=5)

        ttk.Label(right, text="Pola:").pack(anchor="w")
        avail_container = ttk.Frame(right)
        avail_container.pack(fill="both", expand=True)
        avail_canvas = tk.Canvas(avail_container, height=150)
        avail_canvas.pack(side="left", fill="both", expand=True)
        avail_scroll = ttk.Scrollbar(avail_container, orient="vertical", command=avail_canvas.yview)
        avail_scroll.pack(side="right", fill="y")
        avail_canvas.configure(yscrollcommand=avail_scroll.set)
        avail_canvas.bind("<MouseWheel>", lambda e: avail_canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.avail_frame = ttk.Frame(avail_canvas)
        avail_canvas.create_window((0, 0), window=self.avail_frame, anchor="nw")
        self.avail_frame.bind("<Configure>", lambda e: avail_canvas.configure(scrollregion=avail_canvas.bbox("all")))
        self.avail_frame.bind("<MouseWheel>", lambda e: avail_canvas.yview_scroll(int(-e.delta / 120), "units"))

        self.vars = {}
        available = list(parent.columns_vars.keys()) + list(parent.static_vars.keys())
        for name in available:
            var = tk.BooleanVar(value=name in group.field_pos)
            label = parent.display_name(name)
            cb = ttk.Checkbutton(
                self.avail_frame,
                text=label,
                variable=var,
                command=lambda n=name, v=var: self.toggle_field(n, v),
            )
            cb.pack(anchor="w")
            self.vars[name] = var

        for name, pos in group.field_pos.items():
            self.add_element(name, pos)

        self.protocol("WM_DELETE_WINDOW", self.close)

    def draw_grid(self):
        self.canvas.delete("grid")
        step = self.snap_step
        cols = int(self.group.width / step) + 1
        rows = int(self.group.height / step) + 1
        for i in range(cols):
            x = i * step
            self.canvas.create_line(x, 0, x, self.group.height, fill="#ddd", tags="grid")
        for i in range(rows):
            y = i * step
            self.canvas.create_line(0, y, self.group.width, y, fill="#ddd", tags="grid")

    def add_element(self, name, pos=None):
        el = DraggableElement(self, self.canvas, name, name)
        conf = self.group.field_conf.get(name)
        if conf:
            el.width = conf.get("width", el.width)
            el.height = conf.get("height", el.height)
            el.font_size = conf.get("font_size", el.font_size)
            el.bold = conf.get("bold", el.bold)
            el.text_color = conf.get("text_color", el.text_color)
            el.bg_color = conf.get("bg_color", el.bg_color)
            el.bg_visible = conf.get("bg_visible", el.bg_visible)
            el.align = conf.get("align", el.align)
            el.auto_font = conf.get("auto_font", el.auto_font)
        else:
            src = self.parent.elements.get(name)
            if src:
                el.width = src.width
                el.height = src.height
                el.font_size = src.font_size
                el.bold = src.bold
                el.text_color = src.text_color
                el.bg_color = src.bg_color
                el.bg_visible = src.bg_visible
                el.align = src.align
                el.auto_font = src.auto_font
        if pos is not None:
            el.x, el.y = pos
        el.sync_canvas()
        self.elements[name] = el
        if name not in self.group.fields:
            self.group.fields.append(name)

    def toggle_field(self, name, var):
        if var.get():
            if name not in self.elements:
                self.add_element(name, (10, 10))
        else:
            el = self.elements.pop(name, None)
            if el:
                for item in (el.rect, el.label, el.handle, getattr(el, "image_id", None)):
                    if item:
                        self.canvas.delete(item)
            if name in self.group.fields:
                self.group.fields.remove(name)

    def select_element(self, element, additive=False):
        if not additive:
            for el in self.selected_elements:
                self.canvas.itemconfig(el.rect, outline="black")
            self.selected_elements = []
        if element and element not in self.selected_elements:
            self.selected_elements.append(element)
        for el in self.selected_elements:
            self.canvas.itemconfig(el.rect, outline="red")
        self.selected_element = self.selected_elements[-1] if self.selected_elements else None
        if self.selected_element:
            self.font_entry.configure(state="normal")
            self.font_size_var.set(str(int(self.selected_element.font_size)))
            self.transparent_var.set(not self.selected_element.bg_visible)
            self.bg_check.state(["!disabled"])
        else:
            self.font_entry.configure(state="disabled")
            self.font_size_var.set("")
            self.transparent_var.set(False)
            self.bg_check.state(["disabled"])

    def canvas_button_press(self, event):
        if self.canvas.find_withtag("current"):
            return
        self.select_element(None)
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.sel_start = (x, y)
        self.sel_rect = self.canvas.create_rectangle(x, y, x, y, outline="blue", dash=(2, 2))
        self.canvas.tag_raise(self.sel_rect)

    def canvas_drag_select(self, event):
        if not getattr(self, "sel_start", None):
            return
        x0, y0 = self.sel_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.canvas.coords(self.sel_rect, x0, y0, x1, y1)

    def canvas_button_release(self, event):
        if not getattr(self, "sel_start", None):
            if not self.canvas.find_withtag("current"):
                self.select_element(None)
            return
        x0, y0 = self.sel_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.canvas.delete(self.sel_rect)
        self.sel_start = None
        self.sel_rect = None
        if x0 > x1:
            x0, x1 = x1, x0
        if y0 > y1:
            y0, y1 = y1, y0
        self.select_element(None)
        for el in self.elements.values():
            ex0, ey0, ex1, ey1 = self.canvas.coords(el.rect)
            if ex0 >= x0 and ex1 <= x1 and ey0 >= y0 and ey1 <= y1:
                self.select_element(el, additive=True)

    def toggle_bold(self):
        el = self.selected_element
        if not el:
            return
        el.bold = not el.bold
        el.apply_font()

    def increase_font(self):
        el = self.selected_element
        if not el:
            return
        el.font_size += 1
        el.auto_font = False
        el.apply_font()
        self.font_size_var.set(str(int(el.font_size)))

    def decrease_font(self):
        el = self.selected_element
        if not el or el.font_size <= 1:
            return
        el.font_size -= 1
        el.auto_font = False
        el.apply_font()
        self.font_size_var.set(str(int(el.font_size)))

    def set_font_size(self):
        el = self.selected_element
        if not el:
            return
        try:
            size = float(self.font_size_var.get())
        except ValueError:
            return
        if size <= 0:
            return
        el.font_size = size
        el.auto_font = False
        el.apply_font()

    def choose_text_color(self):
        el = self.selected_element
        if not el:
            return
        color = colorchooser.askcolor(color=el.text_color, parent=self)[1]
        if color:
            el.text_color = color
            el.update_colors()
        self.focus_force()

    def choose_bg_color(self):
        el = self.selected_element
        if not el:
            return
        color = colorchooser.askcolor(color=el.bg_color, parent=self)[1]
        if color:
            el.bg_color = color
            el.bg_visible = True
            self.transparent_var.set(False)
            el.update_colors()
        self.focus_force()

    def toggle_bg_visible(self):
        el = self.selected_element
        if not el:
            return
        el.bg_visible = not self.transparent_var.get()
        el.update_colors()

    def set_alignment(self, align):
        if not self.selected_elements:
            return
        for el in self.selected_elements:
            el.align = align
            el.sync_canvas()

    def open_conditions(self):
        win = tk.Toplevel(self)
        win.title("Warunki")
        src_var = tk.StringVar()
        tgt_var = tk.StringVar()
        options = list(self.elements.keys())
        ttk.Label(win, text="Jeśli puste:").grid(row=0, column=0)
        ttk.Combobox(win, values=options, textvariable=src_var, width=20).grid(row=0, column=1)
        ttk.Label(win, text="Ukryj:").grid(row=1, column=0)
        ttk.Combobox(win, values=options, textvariable=tgt_var, width=20).grid(row=1, column=1)
        box = tk.Listbox(win, height=6)
        box.grid(row=2, column=0, columnspan=2, sticky="nsew")
        for s, t in self.conditions:
            box.insert("end", f"{s} -> {t}")
        def add():
            s = src_var.get()
            t = tgt_var.get()
            if s and t:
                self.conditions.append((s, t))
                box.insert("end", f"{s} -> {t}")
        def remove():
            sel = box.curselection()
            if sel:
                idx = sel[0]
                self.conditions.pop(idx)
                box.delete(idx)
        ttk.Button(win, text="Dodaj", command=add).grid(row=3, column=0, sticky="ew")
        ttk.Button(win, text="Usuń", command=remove).grid(row=3, column=1, sticky="ew")

    def push_history(self):
        """Delegate history recording to the main window."""
        if hasattr(self.parent, "push_history"):
            self.parent.push_history()

    def close(self):
        self.group.field_pos = {name: (el.x, el.y) for name, el in self.elements.items()}
        self.group.fields = list(self.group.field_pos.keys())
        self.group.conditions = list(self.conditions)
        self.group.field_conf = {
            name: {
                "width": el.width,
                "height": el.height,
                "font_size": el.font_size,
                "bold": el.bold,
                "text_color": el.text_color,
                "bg_color": el.bg_color,
                "bg_visible": el.bg_visible,
                "align": el.align,
                "auto_font": el.auto_font,
            }
            for name, el in self.elements.items()
        }
        self.group.sync_canvas()
        self.group.draw_preview()
        self.parent.push_history()
        self.group.editor = None
        self.destroy()
# ---------------------------------------------------------------------------
# GUI Application
# ---------------------------------------------------------------------------


class PDSGeneratorGUI(tk.Tk):
    PAGE_SIZES = {
        "A4": (595, 842),  # 210 x 297 mm in points
        "B5": (516, 729),  # 176 x 250 mm
    }

    grid_size = 5

    DEFAULT_STATIC_FIELDS = ["Data", "Naglowek", "Stopka"]

    def __init__(self):
        super().__init__()
        self.title("PDS Generator")
        self.geometry("1200x800")
        self.excel_path = ""
        self.dataframes = {}
        self.elements = {}
        self.groups = {}
        self.conditions = []
        self.selected_elements = []
        self.selected_element = None
        self.sel_rect = None
        self.sel_start = None
        self.page_width, self.page_height = self.PAGE_SIZES["A4"]
        self.scale = 1.0
        self.max_scale = 4.0
        self.min_scale = 1.0
        self.margin = 100  # extra space around the page for panning
        self.snap_step = self.grid_size * self.scale
        self.history = []
        self.future = []
        self.setup_ui()
        self.bind_all("<Control-z>", self.undo)
        self.bind_all("<Control-x>", self.redo)
        self.update_idletasks()
        self.resize_canvas()
        self.load_config(startup=True)
        if not self.history:
            self.push_history()

    # ------------------------------------------------------------------
    def setup_ui(self):
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(top_frame, text="Plik Excel:").pack(side="left")
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(top_frame, textvariable=self.path_var, width=60)
        self.path_entry.pack(side="left", padx=5)
        ttk.Button(top_frame, text="Przeglądaj", command=self.browse_file).pack(side="left")

        ttk.Label(top_frame, text="Rozmiar strony:").pack(side="left", padx=(20, 0))
        self.size_var = tk.StringVar(value="A4")
        self.size_entry = ttk.Entry(top_frame, textvariable=self.size_var, width=10)
        self.size_entry.pack(side="left")
        ttk.Button(top_frame, text="Ustaw", command=self.update_canvas_size).pack(side="left", padx=2)
        self.size_entry.bind("<Return>", lambda e: self.update_canvas_size())

        format_frame = ttk.Frame(self)
        format_frame.pack(fill="x", padx=5)
        ttk.Button(format_frame, text="B", command=self.toggle_bold).pack(side="left")
        ttk.Button(format_frame, text="A+", command=self.increase_font).pack(side="left", padx=2)
        ttk.Button(format_frame, text="A-", command=self.decrease_font).pack(side="left")
        self.font_size_var = tk.StringVar()
        self.font_entry = ttk.Entry(format_frame, textvariable=self.font_size_var, width=4, state="disabled")
        self.font_entry.pack(side="left", padx=5)
        self.font_entry.bind("<Return>", lambda e: self.set_font_size())
        ttk.Button(format_frame, text="Kolor", command=self.choose_text_color).pack(side="left", padx=2)
        ttk.Button(format_frame, text="Tło", command=self.choose_bg_color).pack(side="left", padx=2)
        self.transparent_var = tk.BooleanVar(value=False)
        self.bg_check = ttk.Checkbutton(
            format_frame,
            text="Przezroczyste",
            variable=self.transparent_var,
            command=self.toggle_bg_visible,
        )
        self.bg_check.pack(side="left", padx=2)
        self.bg_check.state(["disabled"])
        ttk.Button(format_frame, text="L", command=lambda: self.set_alignment("left")).pack(side="left", padx=2)
        ttk.Button(format_frame, text="C", command=lambda: self.set_alignment("center")).pack(side="left", padx=2)
        ttk.Button(format_frame, text="R", command=lambda: self.set_alignment("right")).pack(side="left", padx=2)
        ttk.Button(format_frame, text="Środek H", command=self.center_selected_horizontal).pack(side="left", padx=2)
        ttk.Button(format_frame, text="Środek V", command=self.center_selected_vertical).pack(side="left", padx=2)
        self.canvas_container = tk.Frame(self, bg="#b0b0b0")
        self.canvas_container.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.canvas_container.pack_propagate(False)
        self.canvas_container.bind("<Configure>", self.resize_canvas)
        self.canvas = tk.Canvas(
            self.canvas_container,
            bg="#b0b0b0",
            width=self.page_width,
            height=self.page_height,
            highlightthickness=0,
        )
        self.canvas.pack(expand=True)
        self.canvas.bind("<ButtonPress-1>", self.canvas_button_press)
        self.canvas.bind("<B1-Motion>", self.canvas_drag_select)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_button_release)
        self.canvas.bind("<Control-MouseWheel>", self.ctrl_zoom)
        self.canvas.bind("<Control-Button-4>", lambda e: self.ctrl_zoom(e, 120))
        self.canvas.bind("<Control-Button-5>", lambda e: self.ctrl_zoom(e, -120))
        self.canvas.bind("<ButtonPress-2>", self.start_pan)
        self.canvas.bind("<B2-Motion>", self.pan_canvas)
        self.canvas.configure(scrollregion=(-self.margin, -self.margin, self.page_width + self.margin, self.page_height + self.margin))

        zoom_frame = ttk.Frame(self.canvas_container)
        zoom_frame.place(relx=1.0, rely=1.0, anchor="se", x=-5, y=-5)
        ttk.Button(zoom_frame, text="Dopasuj", command=self.fit_to_window).pack(side="right")
        self.zoom_var = tk.StringVar(value="100%")
        ttk.Label(zoom_frame, textvariable=self.zoom_var).pack(side="right", padx=5)

        right_container = ttk.Frame(self)
        right_container.pack(side="right", fill="y", padx=5, pady=5)
        self.right_canvas = tk.Canvas(right_container, width=300)
        right_scroll = ttk.Scrollbar(right_container, orient="vertical", command=self.right_canvas.yview)
        self.right_canvas.configure(yscrollcommand=right_scroll.set)
        right_scroll.pack(side="right", fill="y")
        self.right_canvas.pack(side="left", fill="both", expand=True)
        right_frame = ttk.Frame(self.right_canvas)
        self.right_canvas.create_window((0,0), window=right_frame, anchor="nw")
        right_frame.bind("<Configure>", lambda e: self.right_canvas.configure(scrollregion=self.right_canvas.bbox("all")))
        self.right_canvas.bind("<Enter>", lambda e: self.right_canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.right_canvas.bind("<Leave>", lambda e: self.right_canvas.unbind_all("<MouseWheel>"))
        self.right_canvas.bind("<Button-4>", lambda e: self.right_canvas.yview_scroll(-1, "units"))
        self.right_canvas.bind("<Button-5>", lambda e: self.right_canvas.yview_scroll(1, "units"))

        # Dynamic column checkboxes
        ttk.Label(right_frame, text="Kolumny z Excela:").pack(anchor="w")
        self.columns_frame = ttk.Frame(right_frame)
        self.columns_frame.pack(fill="y", expand=True)
        self.columns_vars = {}

        # Static field checkboxes
        ttk.Label(right_frame, text="Pola statyczne:").pack(anchor="w", pady=(10, 0))
        self.static_frame = ttk.Frame(right_frame)
        self.static_frame.pack(fill="x")
        self.static_vars = {}
        self.static_entries = {}
        self.static_rows = {}
        for field in self.DEFAULT_STATIC_FIELDS:
            self.create_static_row(field, "")
        self.add_static_btn = ttk.Button(self.static_frame, text="Dodaj pole", command=self.add_static_field)
        self.add_static_btn.pack(fill="x", pady=5)

        # Row preview controls
        preview_frame = ttk.Frame(right_frame)
        preview_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(preview_frame, text="Numer wiersza:").pack(side="left")
        self.row_var = tk.StringVar(value="1")
        ttk.Entry(preview_frame, textvariable=self.row_var, width=6).pack(side="left")
        ttk.Button(preview_frame, text="Podgląd", command=self.preview_row).pack(side="left", padx=5)

        # Group list
        ttk.Label(right_frame, text="Grupy:").pack(anchor="w", pady=(10, 0))
        grp_container = ttk.Frame(right_frame)
        grp_container.pack(fill="x")
        self.groups_list = tk.Listbox(grp_container, height=5)
        self.groups_list.pack(side="left", fill="both", expand=True)
        grp_scroll = ttk.Scrollbar(grp_container, orient="vertical", command=self.groups_list.yview)
        grp_scroll.pack(side="right", fill="y")
        self.groups_list.configure(yscrollcommand=grp_scroll.set)
        self.groups_list.bind("<Double-1>", lambda e: self.edit_selected_group())
        ttk.Button(right_frame, text="Usuń grupę", command=self.remove_group).pack(fill="x", pady=(5, 0))

        # Buttons
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        ttk.Button(button_frame, text="Zapisz konfigurację", command=self.save_config).pack(fill="x")
        ttk.Button(button_frame, text="Warunki", command=self.open_conditions).pack(fill="x", pady=5)
        ttk.Button(button_frame, text="Dodaj grupę", command=self.add_group).pack(fill="x", pady=5)
        ttk.Button(button_frame, text="Generuj PDS", command=self.generate_pds).pack(fill="x", pady=5)

        # Progress bar
        self.progress = ttk.Progressbar(right_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(20, 0))
        self.time_label = ttk.Label(right_frame, text="")
        self.time_label.pack()
        self.draw_grid()
        self.bind_all("<Delete>", self.delete_selected)

    # ------------------------------------------------------------------
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.path_var.set(path)
            self.excel_path = path
            self.load_excel(path)
            self.load_config(path=path)
            self.save_config()

    def load_excel(self, path):
        try:
            self.dataframes = pd.read_excel(path, sheet_name=None)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie można wczytać Excela: {e}")
            return
        # Clear previous
        for child in self.columns_frame.winfo_children():
            child.destroy()
        self.columns_vars.clear()

        for sheet, df in self.dataframes.items():
            lf = ttk.LabelFrame(self.columns_frame, text=sheet)
            lf.pack(fill="x", padx=2, pady=2)
            for col in df.columns:
                var = tk.BooleanVar()
                chk = ttk.Checkbutton(
                    lf,
                    text=col,
                    variable=var,
                    command=lambda s=sheet, c=col, v=var: self.toggle_column(f"{s}:{c}", v.get()),
                )
                chk.pack(anchor="w")
                self.columns_vars[f"{sheet}:{col}"] = var

    # ------------------------------------------------------------------
    def update_canvas_size(self):
        value = self.size_var.get().strip()
        if "x" in value.lower():
            try:
                w, h = value.lower().split("x", 1)
                size = (int(float(w)), int(float(h)))
            except Exception:
                messagebox.showerror("Błąd", "Nieprawidłowy format rozmiaru. Użyj np. 595x842")
                return
        else:
            size = self.PAGE_SIZES.get(value.upper(), self.PAGE_SIZES["A4"])
        factor_w = size[0] / self.page_width
        factor_h = size[1] / self.page_height
        self.page_width, self.page_height = size
        for el in self.elements.values():
            el.x *= factor_w
            el.y *= factor_h
            el.width *= factor_w
            el.height *= factor_h
            el.font_size *= factor_h
            step = self.snap_step
            el.x = round(el.x / step) * step
            el.y = round(el.y / step) * step
            el.width = max(step, round(el.width / step) * step)
            el.height = max(step, round(el.height / step) * step)
            el.sync_canvas()
            el.apply_font()
        for group in self.groups.values():
            group.x *= factor_w
            group.y *= factor_h
            group.width *= factor_w
            group.height *= factor_h
            step = self.snap_step
            group.x = round(group.x / step) * step
            group.y = round(group.y / step) * step
            group.width = max(step, round(group.width / step) * step)
            group.height = max(step, round(group.height / step) * step)
            group.sync_canvas()
        self.resize_canvas()

    # ------------------------------------------------------------------
    def toggle_column(self, name, state):
        if state:
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, name)
                self.elements[name] = element
        else:
            self.remove_element(name)
        self.push_history()

    def toggle_static(self, name, state):
        if state:
            value = self.static_entries[name].get()
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, value)
                self.elements[name] = element
            else:
                self.elements[name].update_value(value)
        else:
            self.remove_element(name)
        self.push_history()

    def update_static_value(self, name):
        if name in self.elements:
            self.elements[name].update_value(self.static_entries[name].get())
            self.push_history()

    def display_name(self, name):
        """Return field name including its current text value for lists."""
        if name in self.static_entries:
            text = self.static_entries[name].get()
            if text:
                return f"{name}: {text}"
        el = self.elements.get(name)
        if el and getattr(el, "text", ""):
            return f"{name}: {el.text}"
        return name

    def create_static_row(self, name, value=None):
        row = ttk.Frame(self.static_frame)
        if hasattr(self, "add_static_btn"):
            row.pack(fill="x", pady=2, before=self.add_static_btn)
        else:
            row.pack(fill="x", pady=2)
        var = tk.BooleanVar()
        chk = ttk.Checkbutton(row, variable=var, command=lambda n=name, v=var: self.toggle_static(n, v.get()))
        chk.pack(side="left")
        ttk.Label(row, text=name).pack(side="left")
        entry_var = tk.StringVar(value=value if value is not None else "")
        entry = ttk.Entry(row, textvariable=entry_var, width=15)
        entry.pack(side="left", padx=5)
        entry_var.trace_add("write", lambda *a, f=name: self.update_static_value(f))
        del_btn = ttk.Button(row, text="X", width=2, command=lambda n=name, r=row: self.remove_static_field(n, r))
        del_btn.pack(side="left")
        self.static_vars[name] = var
        self.static_entries[name] = entry_var
        self.static_rows[name] = row

    def add_static_field(self):
        idx = 1
        while f"Static{idx}" in self.static_vars:
            idx += 1
        name = f"Static{idx}"
        self.create_static_row(name, "")

    def remove_static_field(self, name, row):
        if name in self.elements:
            self.remove_element(name)
        row.destroy()
        self.static_vars.pop(name, None)
        self.static_entries.pop(name, None)
        self.static_rows.pop(name, None)
        self.push_history()

    def remove_element(self, name):
        element = self.elements.pop(name, None)
        if element:
            for item in (element.rect, element.label, element.handle):
                self.canvas.delete(item)
            if hasattr(element, "image_id"):
                self.canvas.delete(element.image_id)
            if element in self.selected_elements:
                self.selected_elements.remove(element)
            if self.selected_element is element:
                self.selected_element = None
                self.font_entry.configure(state="disabled")
                self.font_size_var.set("")
        
    def push_history(self):
        state = {
            "elements": [el.to_dict() for el in self.elements.values()],
            "groups": [g.to_dict() for g in self.groups.values()],
        }
        self.history.append(state)
        if len(self.history) > 50:
            self.history.pop(0)
        self.future.clear()

    def restore_state(self, state):
        target = {conf["name"]: conf for conf in state.get("elements", [])}
        # remove elements not in target
        for name in list(self.elements.keys()):
            if name not in target:
                self.remove_element(name)
        for name, conf in target.items():
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, conf.get("text", name))
                self.elements[name] = element
            el = self.elements[name]
            el.x = conf.get("x", 0) * self.scale
            el.y = conf.get("y", 0) * self.scale
            el.width = conf.get("width", 100) * self.scale
            el.height = conf.get("height", 40) * self.scale
            el.font_size = conf.get("font_size", 12) * self.scale
            el.bold = conf.get("bold", False)
            el.text_color = conf.get("text_color", "black")
            el.bg_color = conf.get("bg_color", "white")
            el.bg_visible = conf.get("bg_visible", True)
            el.align = conf.get("align", "left")
            el.auto_font = conf.get("auto_font", True)
            el.sync_canvas()

        # restore groups
        current_groups = list(self.groups.keys())
        for name in current_groups:
            grp = self.groups.pop(name)
            for item in (grp.rect, grp.handle) + tuple(grp.preview_items):
                self.canvas.delete(item)
        self.groups = {}
        for gconf in state.get("groups", []):
            group = GroupArea(self, self.canvas, gconf.get("name", "Group"))
            group.x = gconf.get("x", 0) * self.scale
            group.y = gconf.get("y", 0) * self.scale
            group.width = gconf.get("width", 100) * self.scale
            group.height = gconf.get("height", 100) * self.scale
            group.sync_canvas()
            group.field_pos = {
                k: (v[0] * self.scale, v[1] * self.scale)
                for k, v in gconf.get("field_pos", {}).items()
            }
            group.field_conf = {
                k: {
                    "width": fc.get("width", 100) * self.scale,
                    "height": fc.get("height", 40) * self.scale,
                    "font_size": fc.get("font_size", 12) * self.scale,
                    "bold": fc.get("bold", False),
                    "text_color": fc.get("text_color", "black"),
                    "bg_color": fc.get("bg_color", "white"),
                    "bg_visible": fc.get("bg_visible", True),
                    "align": fc.get("align", "left"),
                    "auto_font": fc.get("auto_font", True),
                }
                for k, fc in gconf.get("field_conf", {}).items()
            }
            group.fields = list(group.field_pos.keys())
            group.conditions = gconf.get("conditions", [])
            group.draw_preview()
            self.groups[group.name] = group
        if hasattr(self, "groups_list"):
            self.groups_list.delete(0, "end")
            for name in self.groups:
                self.groups_list.insert("end", name)

    def undo(self, event=None):
        if len(self.history) < 2:
            return
        state = self.history.pop()
        self.future.append(state)
        self.restore_state(self.history[-1])

    def redo(self, event=None):
        if not self.future:
            return
        state = self.future.pop()
        self.history.append(state)
        self.restore_state(state)

    def add_group(self):
        idx = 1
        while f"Group{idx}" in self.groups:
            idx += 1
        name = f"Group{idx}"
        group = GroupArea(self, self.canvas, name)
        self.groups[name] = group
        if hasattr(self, "groups_list"):
            self.groups_list.insert("end", name)
        self.push_history()

    def edit_selected_group(self):
        sel = self.groups_list.curselection()
        if sel:
            name = self.groups_list.get(sel[0])
            group = self.groups.get(name)
            if group:
                GroupEditor(self, group)

    def remove_group(self):
        sel = self.groups_list.curselection()
        if not sel:
            return
        name = self.groups_list.get(sel[0])
        group = self.groups.pop(name, None)
        if group:
            self.canvas.delete(group.rect)
            self.canvas.delete(group.handle)
            for item in getattr(group, "preview_items", []):
                self.canvas.delete(item)
        self.groups_list.delete(sel[0])
        self.push_history()

    def open_conditions(self):
        win = tk.Toplevel(self)
        win.title("Warunki")
        src_var = tk.StringVar()
        tgt_var = tk.StringVar()
        options = list(self.elements.keys())
        ttk.Label(win, text="Jeśli puste:").grid(row=0, column=0)
        ttk.Combobox(win, values=options, textvariable=src_var, width=20).grid(row=0, column=1)
        ttk.Label(win, text="ukryj:").grid(row=1, column=0)
        ttk.Combobox(win, values=options, textvariable=tgt_var, width=20).grid(row=1, column=1)
        listbox = tk.Listbox(win, width=40)
        listbox.grid(row=3, column=0, columnspan=2, pady=5)
        for s, t in self.conditions:
            listbox.insert("end", f"{t} jeśli {s} puste")
        def add():
            s = src_var.get()
            t = tgt_var.get()
            if s and t and (s, t) not in self.conditions:
                self.conditions.append((s, t))
                listbox.insert("end", f"{t} jeśli {s} puste")
        ttk.Button(win, text="Dodaj", command=add).grid(row=2, column=0, columnspan=2, pady=5)
        def remove():
            sel = listbox.curselection()
            if sel:
                idx = sel[0]
                listbox.delete(idx)
                self.conditions.pop(idx)
        ttk.Button(win, text="Usuń zaznaczone", command=remove).grid(row=4, column=0, columnspan=2, pady=5)

    def element_in_group(self, el, group):
        return (
            el.x >= group.x
            and el.y >= group.y
            and el.x + el.width <= group.x + group.width
            and el.y + el.height <= group.y + group.height
        )

    def draw_pdf_element(self, c, element, value, x, y):
        if isinstance(value, str) and value.lower().startswith("http"):
            try:
                resp = requests.get(value, timeout=5)
                img = Image.open(BytesIO(resp.content))
                c.drawImage(
                    ImageReader(img),
                    x,
                    y,
                    width=element.width / self.scale,
                    height=element.height / self.scale,
                )
                return
            except Exception:
                pass
        if element.bg_visible:
            c.setFillColor(to_reportlab_color(element.bg_color))
            c.rect(
                x,
                y,
                element.width / self.scale,
                element.height / self.scale,
                fill=1,
                stroke=0,
            )
        c.setFillColor(to_reportlab_color(element.text_color))
        c.setFont(
            "Helvetica-Bold" if element.bold else "Helvetica",
            element.font_size / self.scale,
        )
        if element.align == "center":
            c.drawCentredString(
                x + (element.width / self.scale) / 2,
                y + (element.height / self.scale) / 2,
                str(value),
            )
        elif element.align == "right":
            c.drawRightString(
                x + (element.width / self.scale),
                y + (element.height / self.scale) / 2,
                str(value),
            )
        else:
            c.drawString(x, y + (element.height / self.scale) / 2, str(value))

    # ------------------------------------------------------------------
    def preview_row(self):
        if not self.dataframes:
            return
        try:
            idx = int(self.row_var.get()) - 1
        except ValueError:
            messagebox.showerror("Błąd", "Nieprawidłowy numer wiersza")
            return
        for name, element in self.elements.items():
            if ":" in name:
                sheet, col = name.split(":", 1)
                df = self.dataframes.get(sheet)
                value = None
                if df is not None and 0 <= idx < len(df):
                    value = df.iloc[idx].get(col)
            else:
                value = self.static_entries[name].get() if name in getattr(self, "static_entries", {}) else name
            element.update_value(value)

    # ------------------------------------------------------------------
    def save_config(self):
        if not self.excel_path:
            messagebox.showerror("Błąd", "Najpierw wybierz plik Excel")
            return
        config = {
            "excel_path": self.excel_path,
            "page_width": self.page_width,
            "page_height": self.page_height,
            "elements": [el.to_dict() for el in self.elements.values()],
            "static_fields": {name: var.get() for name, var in self.static_entries.items()},
            "conditions": self.conditions,
            "groups": [g.to_dict() for g in self.groups.values()],
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("Zapisano", f"Zapisano konfigurację do {CONFIG_FILE}")

    def load_config(self, startup=False, path=None):
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception:
            return
        excel_cfg = config.get("excel_path")
        if startup and excel_cfg and os.path.exists(excel_cfg):
            self.excel_path = excel_cfg
            self.path_var.set(excel_cfg)
            self.load_excel(excel_cfg)
        if path and excel_cfg != path:
            return
        self.page_width = config.get("page_width", self.page_width)
        self.page_height = config.get("page_height", self.page_height)
        set_name = None
        for n, sz in self.PAGE_SIZES.items():
            if abs(sz[0] - self.page_width) < 1 and abs(sz[1] - self.page_height) < 1:
                set_name = n
                break
        if set_name:
            self.size_var.set(set_name)
        else:
            self.size_var.set(f"{int(self.page_width)}x{int(self.page_height)}")
        self.resize_canvas()
        for name, val in config.get("static_fields", {}).items():
            if name not in self.static_vars:
                self.create_static_row(name, val)
            else:
                self.static_entries[name].set(val)
        self.conditions = config.get("conditions", [])
        for elconf in config.get("elements", []):
            name = elconf["name"]
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, elconf.get("text", name))
                element.x = elconf.get("x", element.x) * self.scale
                element.y = elconf.get("y", element.y) * self.scale
                element.width = elconf.get("width", element.width) * self.scale
                element.height = elconf.get("height", element.height) * self.scale
                element.font_size = elconf.get("font_size", element.font_size) * self.scale
                element.bold = elconf.get("bold", element.bold)
                element.text_color = elconf.get("text_color", element.text_color)
                element.bg_color = elconf.get("bg_color", element.bg_color)
                element.bg_visible = elconf.get("bg_visible", element.bg_visible)
                element.align = elconf.get("align", element.align)
                element.auto_font = elconf.get("auto_font", element.auto_font)
                element.sync_canvas()
                self.elements[name] = element
                if name in self.columns_vars:
                    self.columns_vars[name].set(True)
                if name in self.static_vars:
                    self.static_vars[name].set(True)
                    self.static_entries[name].set(elconf.get("text", ""))
        for gconf in config.get("groups", []):
            group = GroupArea(self, self.canvas, gconf.get("name", f"Group{len(self.groups)+1}"))
            group.x = gconf.get("x", group.x) * self.scale
            group.y = gconf.get("y", group.y) * self.scale
            group.width = gconf.get("width", group.width) * self.scale
            group.height = gconf.get("height", group.height) * self.scale
            group.sync_canvas()
            group.field_pos = {
                k: (v[0] * self.scale, v[1] * self.scale)
                for k, v in gconf.get("field_pos", {}).items()
            }
            group.field_conf = {
                k: {
                    "width": fc.get("width", 100) * self.scale,
                    "height": fc.get("height", 40) * self.scale,
                    "font_size": fc.get("font_size", 12) * self.scale,
                    "bold": fc.get("bold", False),
                    "text_color": fc.get("text_color", "black"),
                    "bg_color": fc.get("bg_color", "white"),
                    "bg_visible": fc.get("bg_visible", True),
                    "align": fc.get("align", "left"),
                    "auto_font": fc.get("auto_font", True),
                }
                for k, fc in gconf.get("field_conf", {}).items()
            }
            group.fields = list(group.field_pos.keys())
            group.conditions = gconf.get("conditions", [])
            group.draw_preview()
            self.groups[group.name] = group
            if hasattr(self, "groups_list"):
                self.groups_list.insert("end", group.name)
        self.push_history()
    # ------------------------------------------------------------------
    def generate_pds(self):
        if not self.excel_path or not self.dataframes:
            messagebox.showerror("Błąd", "Brak danych do generowania")
            return
        # Determine number of rows: assume first dataframe
        first_df = next(iter(self.dataframes.values()))
        total_rows = len(first_df)
        if total_rows == 0:
            messagebox.showinfo("Info", "Brak wierszy w pliku Excel")
            return
        output_dir = os.path.join(os.path.dirname(self.excel_path), "PDS")
        os.makedirs(output_dir, exist_ok=True)

        page_width = self.page_width
        page_height = self.page_height

        def worker():
            start_time = time.time()
            used_names = set()
            for idx in range(total_rows):
                first_val = first_df.iloc[idx, 0] if first_df.shape[1] else ""
                base_filename = sanitize_filename(first_val) or f"pds_{idx+1}"
                filename = base_filename
                suffix = 2
                while (
                    filename in used_names
                    or os.path.exists(os.path.join(output_dir, f"{filename}.pdf"))
                ):
                    filename = f"{base_filename}_{suffix}"
                    suffix += 1
                used_names.add(filename)
                pdf_path = os.path.join(output_dir, f"{filename}.pdf")
                tmp_path = pdf_path + ".tmp"
                c = pdf_canvas.Canvas(tmp_path, pagesize=(page_width, page_height))
                needed = set(self.elements.keys())
                for g in self.groups.values():
                    needed.update(g.fields)
                needed.update(self.static_entries.keys())
                values = {}
                for name in needed:
                    if ":" in name:
                        sheet, col = name.split(":", 1)
                        df = self.dataframes.get(sheet)
                        value = df.iloc[idx].get(col, "") if df is not None else ""
                    else:
                        value = self.static_entries.get(name, tk.StringVar()).get()
                    if pd.isna(value):
                        value = ""
                    values[name] = value
                group_field_names = {fname for g in self.groups.values() for fname in g.fields}

                hidden = set()
                for src, tgt in self.conditions:
                    # Global conditions should not refer to group fields at all
                    if src in group_field_names or tgt in group_field_names:
                        continue
                    if pd.isna(values.get(src)) or values.get(src) == "":
                        hidden.add(tgt)
                for group in self.groups.values():
                    g_hidden = set()
                    for src, tgt in group.conditions:
                        if src not in group.fields or tgt not in group.fields:
                            continue
                        if pd.isna(values.get(src)) or values.get(src) == "":
                            g_hidden.add(tgt)
                    positions = group.field_pos
                    columns = {}
                    for fname in group.fields:
                        if fname in hidden or fname in g_hidden:
                            continue
                        val = values.get(fname, "")
                        if val == "":
                            continue
                        conf = group.field_conf.get(fname, {})
                        el = self.elements.get(fname)
                        if not conf and not el:
                            continue
                        width = conf.get("width", el.width if el else 0)
                        height = conf.get("height", el.height if el else 0)
                        x0, y0 = positions.get(fname, (0, 0))
                        columns.setdefault(x0, []).append((y0, fname, width, height, conf, el, val))

                    placed = []
                    for x0 in sorted(columns):
                        col_items = columns[x0]
                        col_items.sort(key=lambda t: t[0])
                        cur_y = 0
                        for _, fname, width, height, conf, el, val in col_items:
                            y = cur_y
                            while True:
                                overlap = False
                                for px, py, pw, ph in placed:
                                    if (
                                        x0 < px + pw
                                        and x0 + width > px
                                        and y < py + ph
                                        and y + height > py
                                    ):
                                        y = py + ph
                                        overlap = True
                                        break
                                if not overlap:
                                    break
                            if y + height > group.height:
                                continue
                            dummy = SimpleNamespace(
                                width=width,
                                height=height,
                                font_size=conf.get("font_size", el.font_size if el else 12),
                                bold=conf.get("bold", el.bold if el else False),
                                text_color=conf.get("text_color", el.text_color if el else "black"),
                                bg_color=conf.get("bg_color", el.bg_color if el else "white"),
                                bg_visible=conf.get("bg_visible", el.bg_visible if el else True),
                                align=conf.get("align", el.align if el else "left"),
                                auto_font=conf.get("auto_font", el.auto_font if el else True),
                            )
                            x_pdf = (group.x + x0) / self.scale
                            y_pdf = page_height - (group.y + y + height) / self.scale
                            self.draw_pdf_element(c, dummy, val, x_pdf, y_pdf)
                            placed.append((x0, y, width, height))
                            cur_y = y + height
                for name, element in self.elements.items():
                    if name in hidden:
                        continue
                    val = values.get(name, "")
                    x = element.x / self.scale
                    y = page_height - (element.y / self.scale) - (element.height / self.scale)
                    self.draw_pdf_element(c, element, val, x, y)
                c.showPage()
                c.save()
                try:
                    os.replace(tmp_path, pdf_path)
                except Exception:
                    alt_path = pdf_path.replace(
                        ".pdf", f"_{int(time.time())}.pdf"
                    )
                    try:
                        os.replace(tmp_path, alt_path)
                    except Exception:
                        pass
                # Update progress
                progress = (idx + 1) / total_rows * 100
                elapsed = time.time() - start_time
                remaining = (elapsed / (idx + 1)) * (total_rows - idx - 1)
                self.progress.after(0, lambda p=progress: self.progress.config(value=p))
                self.time_label.after(0, lambda r=remaining: self.time_label.config(text=f"Pozostały czas: {int(r)} s"))
            self.progress.after(0, lambda: self.progress.config(value=0))
            self.time_label.after(0, lambda: self.time_label.config(text="Zakończono"))
            messagebox.showinfo("Zakończono", f"Pliki zapisane w {output_dir}")

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    def resize_canvas(self, event=None):
        container_w = self.canvas_container.winfo_width()
        container_h = self.canvas_container.winfo_height()
        if container_w <= 0 or container_h <= 0:
            return
        self.min_scale = min(1.0, container_w / self.page_width, container_h / self.page_height)
        if self.scale < self.min_scale:
            self.fit_to_window()
        else:
            self.canvas.config(width=self.page_width * self.scale, height=self.page_height * self.scale)
            self.draw_grid()
            self.center_page()

    def draw_grid(self):
        self.canvas.delete("grid")
        self.canvas.delete("page")
        self.canvas.delete("ruler")
        step = self.grid_size * self.scale
        self.snap_step = step
        w = self.page_width * self.scale
        h = self.page_height * self.scale
        # keep only a small constant margin so the page can be panned
        # slightly without introducing large grey areas around it
        self.canvas_container.update_idletasks()
        self.margin = 20
        self.canvas.configure(
            scrollregion=(
                -self.margin - 20,
                -self.margin - 20,
                w + self.margin + 20,
                h + self.margin + 20,
            )
        )
        self.canvas.create_rectangle(0, 0, w, h, fill="white", outline="", tags="page")
        # draw rulers background
        self.canvas.create_rectangle(0, -20, w, 0, fill="#e0e0e0", outline="black", tags="ruler")
        self.canvas.create_rectangle(-20, 0, 0, h, fill="#e0e0e0", outline="black", tags="ruler")
        cols = int(w / step) + 1
        rows = int(h / step) + 1
        for i in range(cols):
            x = i * step
            self.canvas.create_line(x, 0, x, h, fill="#9b9b9b", tags="grid")
            self.canvas.create_line(x, -20, x, 0, fill="black", tags="ruler")
            if i % 5 == 0:
                self.canvas.create_text(x + 2, -18, text=str(int(x / self.scale)), anchor="nw", tags="ruler")
        for i in range(rows):
            y = i * step
            self.canvas.create_line(0, y, w, y, fill="#9b9b9b", tags="grid")
            self.canvas.create_line(-20, y, 0, y, fill="black", tags="ruler")
            if i % 5 == 0:
                self.canvas.create_text(-18, y + 2, text=str(int(y / self.scale)), anchor="nw", tags="ruler")
        self.canvas.create_rectangle(0, 0, w, h, outline="black", tags="grid")
        self.canvas.tag_lower("page")
        self.canvas.tag_lower("grid")
        self.canvas.tag_raise("grid", "page")
        self.canvas.tag_raise("ruler", "grid")
        self.zoom_var.set(f"{int(self.scale*100)}%")

    def center_page(self):
        self.canvas.update_idletasks()
        w = self.page_width * self.scale
        h = self.page_height * self.scale
        container_w = self.canvas_container.winfo_width()
        container_h = self.canvas_container.winfo_height()
        if container_w <= 0 or container_h <= 0:
            return
        total_w = w + 2 * (self.margin + 20)
        total_h = h + 2 * (self.margin + 20)
        left = self.margin + 20 + w / 2 - container_w / 2
        top = self.margin + 20 + h / 2 - container_h / 2
        left = max(0, min(left, total_w - container_w))
        top = max(0, min(top, total_h - container_h))
        self.canvas.xview_moveto(left / total_w)
        self.canvas.yview_moveto(top / total_h)
    def ctrl_zoom(self, event, delta=None):
        if delta is None:
            delta = event.delta
        factor = 1.1 if delta > 0 else 0.9
        new_scale = self.scale * factor
        new_scale = max(self.min_scale, min(self.max_scale, new_scale))
        factor = new_scale / self.scale
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        for el in self.elements.values():
            rel_x = el.x / self.scale
            rel_y = el.y / self.scale
            rel_w = el.width / self.scale
            rel_h = el.height / self.scale
            rel_f = el.font_size / self.scale
            el.x = rel_x * new_scale
            el.y = rel_y * new_scale
            el.width = rel_w * new_scale
            el.height = rel_h * new_scale
            el.font_size = rel_f * new_scale
            el.sync_canvas()
            el.apply_font()
        for group in self.groups.values():
            rel_x = group.x / self.scale
            rel_y = group.y / self.scale
            rel_w = group.width / self.scale
            rel_h = group.height / self.scale
            group.x = rel_x * new_scale
            group.y = rel_y * new_scale
            group.width = rel_w * new_scale
            group.height = rel_h * new_scale
            group.sync_canvas()
        self.scale = new_scale
        self.canvas.config(width=self.page_width * self.scale, height=self.page_height * self.scale)
        self.draw_grid()
        w = self.page_width * self.scale
        h = self.page_height * self.scale
        total_w = w + 2 * (self.margin + 20)
        total_h = h + 2 * (self.margin + 20)
        self.canvas.xview_moveto((x * factor - event.x + self.margin + 20) / total_w)
        self.canvas.yview_moveto((y * factor - event.y + self.margin + 20) / total_h)

    def fit_to_window(self):
        container_w = self.canvas_container.winfo_width()
        container_h = self.canvas_container.winfo_height()
        if container_w <= 0 or container_h <= 0:
            return
        new_scale = min(container_w / self.page_width, container_h / self.page_height)
        new_scale = max(self.min_scale, min(self.max_scale, new_scale))
        for el in self.elements.values():
            rel_x = el.x / self.scale
            rel_y = el.y / self.scale
            rel_w = el.width / self.scale
            rel_h = el.height / self.scale
            rel_f = el.font_size / self.scale
            el.x = rel_x * new_scale
            el.y = rel_y * new_scale
            el.width = rel_w * new_scale
            el.height = rel_h * new_scale
            el.font_size = rel_f * new_scale
            el.sync_canvas()
            el.apply_font()
        for group in self.groups.values():
            rel_x = group.x / self.scale
            rel_y = group.y / self.scale
            rel_w = group.width / self.scale
            rel_h = group.height / self.scale
            group.x = rel_x * new_scale
            group.y = rel_y * new_scale
            group.width = rel_w * new_scale
            group.height = rel_h * new_scale
            group.sync_canvas()
        self.scale = new_scale
        self.canvas.config(width=self.page_width * self.scale, height=self.page_height * self.scale)
        self.draw_grid()
        self.after_idle(self.center_page)
        if self.selected_element:
            self.font_size_var.set(str(int(self.selected_element.font_size / self.scale)))

    def start_pan(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def pan_canvas(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def select_element(self, element, additive=False):
        if not additive:
            for el in self.selected_elements:
                self.canvas.itemconfig(el.rect, outline="black")
            self.selected_elements = []
        if element and element not in self.selected_elements:
            self.selected_elements.append(element)
        for el in self.selected_elements:
            self.canvas.itemconfig(el.rect, outline="red")
        self.selected_element = self.selected_elements[-1] if self.selected_elements else None
        if self.selected_element:
            self.font_entry.configure(state="normal")
            self.font_size_var.set(str(int(self.selected_element.font_size / self.scale)))
            self.bg_check.state(["!disabled"])
            self.transparent_var.set(not self.selected_element.bg_visible)
        else:
            self.font_entry.configure(state="disabled")
            self.font_size_var.set("")
            self.transparent_var.set(False)
            self.bg_check.state(["disabled"])

    def canvas_button_press(self, event):
        current = self.canvas.find_withtag("current")
        if current:
            item = current[0]
            for el in self.elements.values():
                if item in (el.rect, el.label, el.handle, getattr(el, "image_id", None)):
                    return
            for group in self.groups.values():
                if item in (group.rect, group.handle):
                    return
        self.select_element(None)
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.sel_start = (x, y)
        self.sel_rect = self.canvas.create_rectangle(
            x,
            y,
            x,
            y,
            outline="blue",
            dash=(2, 2),
        )
        self.canvas.tag_raise(self.sel_rect)

    def canvas_drag_select(self, event):
        if not self.sel_start:
            return
        x0, y0 = self.sel_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.canvas.coords(self.sel_rect, x0, y0, x1, y1)

    def canvas_button_release(self, event):
        if not self.sel_start:
            if not self.canvas.find_withtag("current"):
                self.select_element(None)
            return
        x0, y0 = self.sel_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.canvas.delete(self.sel_rect)
        self.sel_start = None
        self.sel_rect = None
        if x0 > x1:
            x0, x1 = x1, x0
        if y0 > y1:
            y0, y1 = y1, y0
        self.select_element(None)
        for el in self.elements.values():
            ex0, ey0, ex1, ey1 = self.canvas.coords(el.rect)
            if ex0 >= x0 and ex1 <= x1 and ey0 >= y0 and ey1 <= y1:
                self.select_element(el, additive=True)

    def toggle_bold(self):
        el = self.selected_element
        if not el:
            return
        el.bold = not el.bold
        el.apply_font()

    def increase_font(self):
        el = self.selected_element
        if not el:
            return
        el.font_size += self.scale
        el.auto_font = False
        el.apply_font()
        self.font_size_var.set(str(int(el.font_size / self.scale)))
        self.push_history()

    def decrease_font(self):
        el = self.selected_element
        if not el:
            return
        if el.font_size > self.scale:
            el.font_size -= self.scale
            el.auto_font = False
            el.apply_font()
            self.font_size_var.set(str(int(el.font_size / self.scale)))
            self.push_history()

    def set_font_size(self):
        el = self.selected_element
        if not el:
            return
        try:
            size = float(self.font_size_var.get()) * self.scale
        except ValueError:
            return
        if size <= 0:
            return
        el.font_size = size
        el.auto_font = False
        el.apply_font()
        self.push_history()

    def choose_text_color(self):
        el = self.selected_element
        if not el:
            return
        color = colorchooser.askcolor(color=el.text_color, parent=self)[1]
        if color:
            el.text_color = color
            el.update_colors()
            self.push_history()
        self.focus_force()

    def choose_bg_color(self):
        el = self.selected_element
        if not el:
            return
        color = colorchooser.askcolor(color=el.bg_color, parent=self)[1]
        if color:
            el.bg_color = color
            el.bg_visible = True
            self.transparent_var.set(False)
            el.update_colors()
            self.push_history()
        self.focus_force()

    def toggle_bg_visible(self):
        el = self.selected_element
        if not el:
            return
        el.bg_visible = not self.transparent_var.get()
        el.update_colors()
        self.push_history()

    def set_alignment(self, align):
        if not self.selected_elements:
            return
        for el in self.selected_elements:
            el.align = align
            el.sync_canvas()
        self.push_history()

    def center_selected_horizontal(self):
        if not self.selected_elements:
            return
        for el in self.selected_elements:
            el.x = (self.page_width * self.scale - el.width) / 2
            el.sync_canvas()
        self.push_history()

    def center_selected_vertical(self):
        if not self.selected_elements:
            return
        for el in self.selected_elements:
            el.y = (self.page_height * self.scale - el.height) / 2
            el.sync_canvas()
        self.push_history()

    def delete_selected(self, event=None):
        if not self.selected_elements:
            return
        for el in list(self.selected_elements):
            name = el.name
            self.remove_element(name)
            if name in self.columns_vars:
                self.columns_vars[name].set(False)
            if name in self.static_vars:
                self.static_vars[name].set(False)
        self.selected_elements = []
        self.selected_element = None
        self.font_entry.configure(state="disabled")
        self.font_size_var.set("")
        self.push_history()

    def _on_mousewheel(self, event):
        self.right_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


if __name__ == "__main__":
    app = PDSGeneratorGUI()
    app.mainloop()
