import os
import sys
import json
import time
import threading
import subprocess
import importlib
import math
from io import BytesIO

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
            el.x = round(el.x / step) * step
            el.y = round(el.y / step) * step
            el.sync_canvas()

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
        self.width = max(step, round(self.width / step) * step)
        self.height = max(step, round(self.height / step) * step)
        self.sync_canvas()

    # ------------------------------------------------------------------
    def to_dict(self):
        scale = self.parent.scale
        return {
            "name": self.name,
            "text": self.text,
            "x": self.x / scale,
            "y": self.y / scale,
            "width": self.width / scale,
            "height": self.height / scale,
            "font_size": self.font_size / scale,
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

    HANDLE_SIZE = 8

    def __init__(self, parent, canvas: tk.Canvas, name: str):
        self.parent = parent
        self.canvas = canvas
        self.name = name
        self.x = canvas.winfo_width() // 2 - 50
        self.y = canvas.winfo_height() // 2 - 50
        self.width = 100
        self.height = 100
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
            fill="blue",
        )
        for tag in (self.rect,):
            canvas.tag_bind(tag, "<ButtonPress-1>", self.start_move)
            canvas.tag_bind(tag, "<B1-Motion>", self.moving)
            canvas.tag_bind(tag, "<ButtonRelease-1>", self.stop_move)
        canvas.tag_bind(self.handle, "<ButtonPress-1>", self.start_resize)
        canvas.tag_bind(self.handle, "<B1-Motion>", self.resizing)
        canvas.tag_bind(self.handle, "<ButtonRelease-1>", self.stop_resize)
        self.send_to_back()

    def send_to_back(self):
        self.canvas.tag_lower(self.rect)
        self.canvas.tag_lower(self.handle)

    def start_move(self, event):
        self.last_x = event.x
        self.last_y = event.y

    def moving(self, event):
        dx = event.x - self.last_x
        dy = event.y - self.last_y
        for item in (self.rect, self.handle):
            self.canvas.move(item, dx, dy)
        self.x += dx
        self.y += dy
        self.last_x = event.x
        self.last_y = event.y

    def stop_move(self, event):
        step = self.parent.snap_step
        self.x = round(self.x / step) * step
        self.y = round(self.y / step) * step
        self.sync_canvas()

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
        self.width = max(step, round(self.width / step) * step)
        self.height = max(step, round(self.height / step) * step)
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

    def to_dict(self):
        scale = self.parent.scale
        return {
            "name": self.name,
            "x": self.x / scale,
            "y": self.y / scale,
            "width": self.width / scale,
            "height": self.height / scale,
        }


# ---------------------------------------------------------------------------
# GUI Application
# ---------------------------------------------------------------------------


class PDSGeneratorGUI(tk.Tk):
    PAGE_SIZES = {
        "A4": (595, 842),  # 210 x 297 mm in points
        "B5": (516, 729),  # 176 x 250 mm
    }

    grid_size = 10

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
        self.setup_ui()
        self.update_idletasks()
        self.resize_canvas()
        self.load_config(startup=True)

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
        right_container.pack(side="left", fill="y", padx=5, pady=5)
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
        self.resize_canvas()

    # ------------------------------------------------------------------
    def toggle_column(self, name, state):
        if state:
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, name)
                self.elements[name] = element
        else:
            self.remove_element(name)

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

    def update_static_value(self, name):
        if name in self.elements:
            self.elements[name].update_value(self.static_entries[name].get())

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

    def add_group(self):
        idx = 1
        while f"Group{idx}" in self.groups:
            idx += 1
        name = f"Group{idx}"
        group = GroupArea(self, self.canvas, name)
        self.groups[name] = group

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
            self.groups[group.name] = group

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
            for idx in range(total_rows):
                pdf_path = os.path.join(output_dir, f"pds_{idx+1}.pdf")
                tmp_path = pdf_path + ".tmp"
                c = pdf_canvas.Canvas(tmp_path, pagesize=(page_width, page_height))
                values = {}
                for name, element in self.elements.items():
                    if ":" in name:
                        sheet, col = name.split(":", 1)
                        df = self.dataframes.get(sheet)
                        value = df.iloc[idx].get(col, "") if df is not None else ""
                    else:
                        value = self.static_entries.get(name, tk.StringVar()).get()
                    if pd.isna(value):
                        value = ""
                    values[name] = value
                hidden = set()
                for src, tgt in self.conditions:
                    if pd.isna(values.get(src)) or values.get(src) == "":
                        hidden.add(tgt)
                processed = set()
                for group in self.groups.values():
                    inside = [el for el in self.elements.values() if self.element_in_group(el, group)]
                    inside.sort(key=lambda e: e.y)
                    current_y = group.y
                    for el in inside:
                        if el.name in hidden:
                            continue
                        val = values.get(el.name, "")
                        if val == "":
                            continue
                        x = el.x / self.scale
                        y = page_height - (current_y / self.scale) - (el.height / self.scale)
                        self.draw_pdf_element(c, el, val, x, y)
                        current_y += el.height
                        processed.add(el.name)
                for name, element in self.elements.items():
                    if name in hidden or name in processed:
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
        while step < 10:
            step *= 2
        while step > 40:
            step /= 2
        self.snap_step = step
        w = self.page_width * self.scale
        h = self.page_height * self.scale
        # keep a constant margin based on the window size so zooming
        # does not shrink the available panning area
        base = max(
            self.canvas_container.winfo_width(),
            self.canvas_container.winfo_height(),
        )
        self.margin = base * 2
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
            x = int(round(i * step))
            self.canvas.create_line(x, 0, x, int(h), fill="#9b9b9b", tags="grid")
            self.canvas.create_line(x, -20, x, 0, fill="black", tags="ruler")
            if i % 5 == 0:
                self.canvas.create_text(x + 2, -18, text=str(int(x / self.scale)), anchor="nw", tags="ruler")
        for i in range(rows):
            y = int(round(i * step))
            self.canvas.create_line(0, y, int(w), y, fill="#9b9b9b", tags="grid")
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
        left = (w - container_w) / 2 + self.margin + 20
        top = (h - container_h) / 2 + self.margin + 20
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
        new_scale = min(1.0, container_w / self.page_width, container_h / self.page_height)
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
        if self.canvas.find_withtag("current"):
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

    def decrease_font(self):
        el = self.selected_element
        if not el:
            return
        if el.font_size > self.scale:
            el.font_size -= self.scale
            el.auto_font = False
            el.apply_font()
            self.font_size_var.set(str(int(el.font_size / self.scale)))

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

    def choose_text_color(self):
        el = self.selected_element
        if not el:
            return
        color = colorchooser.askcolor(color=el.text_color)[1]
        if color:
            el.text_color = color
            el.update_colors()

    def choose_bg_color(self):
        el = self.selected_element
        if not el:
            return
        color = colorchooser.askcolor(color=el.bg_color)[1]
        if color:
            el.bg_color = color
            el.bg_visible = True
            self.transparent_var.set(False)
            el.update_colors()

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

    def center_selected_horizontal(self):
        if not self.selected_elements:
            return
        for el in self.selected_elements:
            el.x = (self.page_width * self.scale - el.width) / 2
            el.sync_canvas()

    def center_selected_vertical(self):
        if not self.selected_elements:
            return
        for el in self.selected_elements:
            el.y = (self.page_height * self.scale - el.height) / 2
            el.sync_canvas()

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

    def _on_mousewheel(self, event):
        self.right_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


if __name__ == "__main__":
    app = PDSGeneratorGUI()
    app.mainloop()
