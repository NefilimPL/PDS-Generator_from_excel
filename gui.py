import logging
import os
import json
import time
import threading
import math
import re
from io import BytesIO
from types import SimpleNamespace

import pandas as pd
from PIL import Image, ImageTk
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter import font as tkfont

from elements import DraggableElement
from groups import GroupArea, GroupEditor

logger = logging.getLogger(__name__)

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")


def to_reportlab_color(value):
    try:
        return colors.HexColor(value)
    except (ValueError, TypeError):
        return colors.toColor(value)


def sanitize_filename(name: str) -> str:
    """Return a filesystem-safe version of *name*."""
    cleaned = re.sub(r"[^\w\s-]", "", str(name))
    cleaned = re.sub(r"\s+", "_", cleaned).strip("_")
    return cleaned

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
        self.image_cache = {}
        self.selected_elements = []
        self.selected_element = None
        self.sel_rect = None
        self.sel_start = None
        self.align_line_h = None
        self.align_line_v = None
        self.page_width, self.page_height = self.PAGE_SIZES["A4"]
        self.scale = 1.0
        self.max_scale = 4.0
        self.min_scale = 1.0
        # Minimal margins around the page; updated dynamically
        self.margin_x = 20
        self.margin_y = 20
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
        ttk.Label(format_frame, text="Warstwa:").pack(side="left", padx=(5, 0))
        self.layer_var = tk.StringVar()
        self.layer_entry = ttk.Entry(format_frame, textvariable=self.layer_var, width=4, state="disabled")
        self.layer_entry.pack(side="left", padx=2)
        self.layer_entry.bind("<Return>", lambda e: self.set_layer())
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
        self.canvas.configure(
            scrollregion=(
                -self.margin_x,
                -self.margin_y,
                self.page_width + self.margin_x,
                self.page_height + self.margin_y,
            )
        )

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
            self.image_cache = {}
            self.load_excel(path)
            self.load_config(path=path)
            self.save_config()

    def load_excel(self, path):
        try:
            self.dataframes = pd.read_excel(path, sheet_name=None)
        except (OSError, ValueError) as e:
            logger.exception("Failed to read Excel file %s", path)
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
    def find_local_image(self, filename):
        """Search for an image file relative to the Excel file directory."""
        if not filename or not getattr(self, "excel_path", ""):
            return None
        key = filename.lower()
        if key in self.image_cache:
            return self.image_cache[key]
        base_dir = os.path.dirname(self.excel_path)
        if os.path.isabs(filename):
            path = filename if os.path.exists(filename) else None
        else:
            candidate = os.path.join(base_dir, filename)
            if os.path.exists(candidate):
                path = candidate
            else:
                path = None
                for root, _, files in os.walk(base_dir):
                    for f in files:
                        if f.lower() == key:
                            path = os.path.join(root, f)
                            break
                    if path:
                        break
        self.image_cache[key] = path
        return path

    # ------------------------------------------------------------------
    def update_canvas_size(self):
        value = self.size_var.get().strip()
        if "x" in value.lower():
            try:
                w, h = value.lower().split("x", 1)
                size = (int(float(w)), int(float(h)))
            except ValueError:
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
                self.restack_elements()
        else:
            self.remove_element(name)
        self.push_history()

    def toggle_static(self, name, state):
        if state:
            value = self.static_entries[name].get()
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, value)
                self.elements[name] = element
                self.restack_elements()
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
                self.layer_entry.configure(state="disabled")
                self.layer_var.set("")
        self.restack_elements()

    def restack_elements(self):
        if not self.elements:
            return
        min_layer = min(el.layer for el in self.elements.values())
        if min_layer < 1:
            shift = 1 - min_layer
            for el in self.elements.values():
                el.layer += shift
        for el in sorted(self.elements.values(), key=lambda e: e.layer):
            for item in filter(None, [
                el.rect,
                el.label,
                getattr(el, "image_id", None),
                el.handle,
            ]):
                self.canvas.tag_raise(item)
        self.canvas.tag_lower("page")
        self.canvas.tag_lower("grid")
        self.canvas.tag_raise("grid", "page")
        if self.selected_element:
            self.layer_var.set(str(int(self.selected_element.layer)))
        
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
            el.layer = conf.get("layer", el.layer)
            el.sync_canvas()

        self.restack_elements()

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
                k: (v[0], v[1]) for k, v in gconf.get("field_pos", {}).items()
            }
            group.field_conf = {
                k: {
                    "width": fc.get("width", 100),
                    "height": fc.get("height", 40),
                    "font_size": fc.get("font_size", 12),
                    "bold": fc.get("bold", False),
                    "text_color": fc.get("text_color", "black"),
                    "bg_color": fc.get("bg_color", "white"),
                    "bg_visible": fc.get("bg_visible", True),
                    "align": fc.get("align", "left"),
                    "auto_font": fc.get("auto_font", True),
                    "layer": fc.get("layer", 1),
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
            except (requests.RequestException, OSError) as exc:
                logger.exception("Failed to load remote image %s", value)
        if isinstance(value, str):
            local_path = self.find_local_image(value)
            if local_path:
                try:
                    img = Image.open(local_path)
                    c.drawImage(
                        ImageReader(img),
                        x,
                        y,
                        width=element.width / self.scale,
                        height=element.height / self.scale,
                    )
                    return
                except OSError as exc:
                    logger.exception("Failed to load local image %s", local_path)
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
        for name, element in sorted(self.elements.items(), key=lambda kv: kv[1].layer):
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
        except (OSError, json.JSONDecodeError) as exc:
            logger.exception("Failed to load config from %s", CONFIG_FILE)
            return
        excel_cfg = config.get("excel_path")
        if startup and excel_cfg and os.path.exists(excel_cfg):
            self.excel_path = excel_cfg
            self.image_cache = {}
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
                element.layer = elconf.get("layer", element.layer)
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
                k: (v[0], v[1]) for k, v in gconf.get("field_pos", {}).items()
            }
            group.field_conf = {
                k: {
                    "width": fc.get("width", 100),
                    "height": fc.get("height", 40),
                    "font_size": fc.get("font_size", 12),
                    "bold": fc.get("bold", False),
                    "text_color": fc.get("text_color", "black"),
                    "bg_color": fc.get("bg_color", "white"),
                    "bg_visible": fc.get("bg_visible", True),
                    "align": fc.get("align", "left"),
                    "auto_font": fc.get("auto_font", True),
                    "layer": fc.get("layer", 1),
                }
                for k, fc in gconf.get("field_conf", {}).items()
            }
            group.fields = list(group.field_pos.keys())
            group.conditions = gconf.get("conditions", [])
            group.draw_preview()
            self.groups[group.name] = group
            if hasattr(self, "groups_list"):
                self.groups_list.insert("end", group.name)
        self.restack_elements()
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
            for idx in range(total_rows):
                first_val = first_df.iloc[idx, 0] if first_df.shape[1] else ""
                filename = sanitize_filename(first_val) or f"pds_{idx+1}"
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
                for name, element in sorted(self.elements.items(), key=lambda kv: kv[1].layer):
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
                except OSError as exc:
                    logger.exception("Failed to replace %s, trying alternative name", pdf_path)
                    alt_path = pdf_path.replace(
                        ".pdf", f"_{int(time.time())}.pdf"
                    )
                    try:
                        os.replace(tmp_path, alt_path)
                    except OSError:
                        logger.exception("Failed to move temp PDF to %s", alt_path)
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
            self.update_scrollregion()
            self.center_page()

    def update_scrollregion(self):
        """Update canvas scrollregion to allow panning around the page."""
        w = self.page_width * self.scale
        h = self.page_height * self.scale
        container_w = self.canvas_container.winfo_width()
        container_h = self.canvas_container.winfo_height()
        # Ensure the scrollable area is always larger than the visible
        # container so the page can be panned freely without disappearing
        # under the gray background. When the window is larger than the
        # page, use the container dimensions as extra margins so there is
        # always space to drag the page in any direction.
        margin_x = max(20, container_w)
        margin_y = max(20, container_h)
        self.margin_x = margin_x
        self.margin_y = margin_y
        self.canvas.configure(scrollregion=(-margin_x, -margin_y, w + margin_x, h + margin_y))

    def draw_grid(self):
        self.canvas.delete("grid")
        self.canvas.delete("page")
        self.canvas.delete("ruler")
        step = self.grid_size * self.scale
        self.snap_step = step
        w = self.page_width * self.scale
        h = self.page_height * self.scale
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
        self.update_scrollregion()

    def clear_alignment_guides(self):
        for line in (self.align_line_h, self.align_line_v):
            if line:
                self.canvas.delete(line)
        self.align_line_h = self.align_line_v = None

    def update_alignment_guides(self, element, resize=False):
        self.clear_alignment_guides()
        others = [
            el
            for el in list(self.elements.values()) + list(self.groups.values())
            if el is not element
        ]
        x1, y1 = element.x, element.y
        x2, y2 = element.x + element.width, element.y + element.height
        tol = 5
        snap_dx = snap_dy = 0
        for other in others:
            ox1, oy1 = other.x, other.y
            ox2, oy2 = other.x + other.width, other.y + other.height
            if not self.align_line_v:
                edges = [x2] if resize else [x1, x2]
                for x in edges:
                    for ox in (ox1, ox2):
                        if abs(x - ox) <= tol:
                            snap_dx = ox - x
                            self.align_line_v = self.canvas.create_line(
                                ox, min(y1, oy1), ox, max(y2, oy2), fill="red"
                            )
                            break
                    if self.align_line_v:
                        break
            if not self.align_line_h:
                edges = [y2] if resize else [y1, y2]
                for y in edges:
                    for oy in (oy1, oy2):
                        if abs(y - oy) <= tol:
                            snap_dy = oy - y
                            self.align_line_h = self.canvas.create_line(
                                min(x1, ox1), oy, max(x2, ox2), oy, fill="red"
                            )
                            break
                    if self.align_line_h:
                        break
            if self.align_line_h and self.align_line_v:
                break
        self.zoom_var.set(f"{int(self.scale*100)}%")
        return snap_dx, snap_dy

    def center_page(self):
        self.canvas.update_idletasks()
        w = self.page_width * self.scale
        h = self.page_height * self.scale
        container_w = self.canvas_container.winfo_width()
        container_h = self.canvas_container.winfo_height()
        if container_w <= 0 or container_h <= 0:
            return
        total_w = w + 2 * self.margin_x
        total_h = h + 2 * self.margin_y
        left = self.margin_x + w / 2 - container_w / 2
        top = self.margin_y + h / 2 - container_h / 2
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
        total_w = w + 2 * self.margin_x
        total_h = h + 2 * self.margin_y
        self.canvas.xview_moveto((x * factor - event.x + self.margin_x) / total_w)
        self.canvas.yview_moveto((y * factor - event.y + self.margin_y) / total_h)

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
        self.update_scrollregion()
        self.after_idle(self.center_page)
        if self.selected_element:
            self.font_size_var.set(str(int(self.selected_element.font_size / self.scale)))

    def start_pan(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def pan_canvas(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def select_element(self, element, additive=False):
        self.clear_alignment_guides()
        if not additive:
            for el in self.selected_elements:
                self.canvas.itemconfig(el.rect, outline="black", width=1)
            self.selected_elements = []
        if element and element not in self.selected_elements:
            self.selected_elements.append(element)
        for el in self.selected_elements:
            self.canvas.itemconfig(el.rect, outline="red", width=2)
        self.selected_element = self.selected_elements[-1] if self.selected_elements else None
        if self.selected_element:
            self.font_entry.configure(state="normal")
            self.font_size_var.set(str(int(self.selected_element.font_size / self.scale)))
            self.bg_check.state(["!disabled"])
            self.transparent_var.set(not self.selected_element.bg_visible)
            self.layer_entry.configure(state="normal")
            self.layer_var.set(str(int(self.selected_element.layer)))
        else:
            self.font_entry.configure(state="disabled")
            self.font_size_var.set("")
            self.transparent_var.set(False)
            self.bg_check.state(["disabled"])
            self.layer_entry.configure(state="disabled")
            self.layer_var.set("")

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
            width=2,
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

    def set_layer(self):
        el = self.selected_element
        if not el:
            return
        try:
            layer = int(float(self.layer_var.get()))
        except ValueError:
            return
        if layer < 1:
            layer = 1
        el.layer = layer
        self.restack_elements()
        self.push_history()
        self.layer_var.set(str(int(el.layer)))

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


