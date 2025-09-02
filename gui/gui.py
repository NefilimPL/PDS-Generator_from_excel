import logging
import os
import sys
import webbrowser

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter import font as tkfont
from PIL import Image, ImageTk

from elements import DraggableElement
from groups import GroupArea, GroupEditor

from .ui_layout import setup_ui as build_ui
from .pdf_export import (
    generate_pds as export_pds,
    draw_pdf_element as render_pdf_element,
)
from .config_io import (
    save_config as save_config_func,
    load_config as load_config_func,
)

from github_utils import (
    get_repo_info,
    get_remote_hash,
    get_remote_version,
    pull_updates,
    get_last_update_date,
    get_version,
)

logger = logging.getLogger(__name__)

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
        self.repo_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        self.version = get_version(self.repo_dir)
        self.last_update = get_last_update_date(self.repo_dir) or ""
        icon_path = os.path.join(os.path.dirname(__file__), "github_icon.png")
        self.github_image = None
        if os.path.exists(icon_path):
            try:
                img = Image.open(icon_path)
                img = img.resize((24, 24), Image.LANCZOS)
                self.github_image = ImageTk.PhotoImage(img)
            except Exception:  # pragma: no cover - depends on Pillow/Tk installation
                logger.debug("Failed to load github icon from %s", icon_path)
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
        self.margin = 100  # extra space around the page for panning
        self.snap_step = self.grid_size * self.scale
        self.history = []
        self.future = []
        self.ignore_updates = False
        self.update_test = False
        self.update_available = False
        self.repo_owner = None
        self.repo_name = None
        self.blink_state = False
        self.setup_ui()
        self.bind_all("<Control-z>", self.undo)
        self.bind_all("<Control-x>", self.redo)
        self.update_idletasks()
        self.resize_canvas()
        self.load_config(startup=True)
        self.check_for_updates()
        if not self.history:
            self.push_history()

    # ------------------------------------------------------------------
    def setup_ui(self):
        build_ui(self)

    # ------------------------------------------------------------------
    def check_for_updates(self):
        local_hash, owner, repo = get_repo_info(self.repo_dir)
        self.repo_owner, self.repo_name = owner, repo

        remote_version = get_remote_version(owner, repo)
        if remote_version and remote_version != self.version:
            self.update_available = True
        else:
            remote_hash = get_remote_hash(owner, repo)
            self.update_available = bool(local_hash and remote_hash and remote_hash != local_hash)

        if self.update_available:
            self.update_button.pack(side="left", padx=5)
            self.blink_update_button()
        else:
            self.update_button.pack_forget()
        should_prompt = False
        if self.update_test:
            should_prompt = True
        elif self.update_available and not self.ignore_updates:
            should_prompt = True
        if should_prompt:
            win = tk.Toplevel(self)
            win.title("Aktualizacja")
            ttk.Label(
                win, text="Dostępna jest nowa wersja aplikacji."
            ).pack(padx=10, pady=10)
            if owner and repo:
                link = ttk.Label(
                    win, text="Repozytorium", foreground="blue", cursor="hand2"
                )
                link.pack()
                link.bind(
                    "<Button-1>",
                    lambda e: webbrowser.open(
                        f"https://github.com/{owner}/{repo}"
                    ),
                )
            btns = ttk.Frame(win)
            btns.pack(pady=10)

            def do_update():
                win.destroy()
                if self.update_test:
                    messagebox.showinfo(
                        "Aktualizacja", "Symulacja pobierania aktualizacji."
                    )
                    return
                if not pull_updates(self.repo_dir):
                    messagebox.showerror(
                        "Błąd", "Aktualizacja nie powiodła się"
                    )
                    return
                python = sys.executable
                os.execl(python, python, *sys.argv)

            ttk.Button(btns, text="Aktualizuj", command=do_update).pack(
                side="left", padx=5
            )
            ttk.Button(btns, text="Pomiń", command=win.destroy).pack(
                side="left", padx=5
            )

    def manual_update(self):
        if self.update_test:
            messagebox.showinfo(
                "Aktualizacja", "Symulacja pobierania aktualizacji."
            )
            return
        if not pull_updates(self.repo_dir):
            messagebox.showerror("Błąd", "Aktualizacja nie powiodła się")
            return
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def open_github(self):
        if self.repo_owner and self.repo_name:
            webbrowser.open(
                f"https://github.com/{self.repo_owner}/{self.repo_name}"
            )

    def blink_update_button(self):
        if not self.update_available:
            self.update_button.configure(background=self.update_button_bg)
            return
        color = "red" if not self.blink_state else self.update_button_bg
        self.update_button.configure(background=color)
        self.blink_state = not self.blink_state
        self.after(500, self.blink_update_button)

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
        render_pdf_element(self, c, element, value, x, y)

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
        save_config_func(self)

    def load_config(self, startup=False, path=None):
        load_config_func(self, startup=startup, path=path)
    # ------------------------------------------------------------------
    def generate_pds(self):
        export_pds(self)

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
            self.canvas.config(width=container_w, height=container_h)
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
        container_w = self.canvas_container.winfo_width()
        container_h = self.canvas_container.winfo_height()
        self.canvas.config(width=container_w, height=container_h)
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
        container_w = self.canvas_container.winfo_width()
        container_h = self.canvas_container.winfo_height()
        self.canvas.config(width=container_w, height=container_h)
        self.draw_grid()
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


