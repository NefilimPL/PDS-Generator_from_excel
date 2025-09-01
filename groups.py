import logging

import tkinter as tk
from tkinter import ttk, colorchooser

from elements import DraggableElement

logger = logging.getLogger(__name__)

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
        snap_dx, snap_dy = self.parent.update_alignment_guides(self)
        if snap_dx or snap_dy:
            for item in (self.rect, self.handle):
                self.canvas.move(item, snap_dx, snap_dy)
            for el in self.children:
                for item in (el.rect, el.label, el.handle, getattr(el, "image_id", None)):
                    if item:
                        self.canvas.move(item, snap_dx, snap_dy)
                el.x += snap_dx
                el.y += snap_dy
            self.x += snap_dx
            self.y += snap_dy
            self.last_x += snap_dx
            self.last_y += snap_dy
            self.parent.update_alignment_guides(self)

    def stop_move(self, event):
        step = self.parent.snap_step
        new_x = int(round(self.x / step)) * step
        new_y = int(round(self.y / step)) * step
        dx = new_x - self.x
        dy = new_y - self.y
        self.x = new_x
        self.y = new_y
        # ensure the group's dimensions also align with the grid
        self.width = max(step, int(round(self.width / step)) * step)
        self.height = max(step, int(round(self.height / step)) * step)
        self.sync_canvas()
        # snap children by the same offset
        if dx or dy:
            for el in self.children:
                for item in (el.rect, el.label, el.handle, getattr(el, "image_id", None)):
                    if item:
                        self.canvas.move(item, dx, dy)
                el.x += dx
                el.y += dy
        self.parent.clear_alignment_guides()

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
        snap_w, snap_h = self.parent.update_alignment_guides(self, resize=True)
        if snap_w or snap_h:
            self.width += snap_w
            self.height += snap_h
            self.sync_canvas()
            self.start_w += snap_w
            self.start_h += snap_h
            self.parent.update_alignment_guides(self, resize=True)

    def stop_resize(self, event):
        step = self.parent.snap_step
        self.width = max(step, int(round(self.width / step)) * step)
        self.height = max(step, int(round(self.height / step)) * step)
        self.sync_canvas()
        self.parent.clear_alignment_guides()

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
                k: (int(round(v[0])), int(round(v[1])))
                for k, v in self.field_pos.items()
            },
            "field_conf": {
                k: {
                    "width": int(round(conf["width"])),
                    "height": int(round(conf["height"])),
                    "font_size": int(round(conf["font_size"])),
                    "bold": conf.get("bold", False),
                    "text_color": conf.get("text_color", "black"),
                    "bg_color": conf.get("bg_color", "white"),
                    "bg_visible": conf.get("bg_visible", True),
                    "align": conf.get("align", "left"),
                    "auto_font": conf.get("auto_font", True),
                    "layer": conf.get("layer", 1),
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
        # Build columns keyed by their x position (unscaled values)
        cols = {}
        scale = self.parent.scale
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
                # scale positions and sizes for canvas display
                sx = x * scale
                sy = y * scale
                sw = w * scale
                sh = h * scale
                x1 = self.x + sx
                y1 = self.y + sy
                r = self.canvas.create_rectangle(
                    x1, y1, x1 + sw, y1 + sh, outline="blue", fill="white"
                )
                t = self.canvas.create_text(x1 + 2, y1 + sh / 2, anchor="w", text=name)
                for item in (r, t):
                    self.canvas.tag_bind(item, "<ButtonPress-1>", self.start_move)
                    self.canvas.tag_bind(item, "<B1-Motion>", self.moving)
                    self.canvas.tag_bind(item, "<ButtonRelease-1>", self.stop_move)
                    self.canvas.tag_bind(item, "<Double-1>", self.open_editor)
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
        # adopt main editor zoom so items match in size
        self.scale = parent.scale
        self.grid_size = parent.grid_size
        self.snap_step = self.grid_size * self.scale
        self.elements = {}
        self.selected_elements = []
        self.selected_element = None
        self.conditions = list(group.conditions)
        self.align_line_h = None
        self.align_line_v = None

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=5, pady=5)
        ttk.Button(toolbar, text="B", command=self.toggle_bold).pack(side="left")
        ttk.Button(toolbar, text="A+", command=self.increase_font).pack(side="left", padx=2)
        ttk.Button(toolbar, text="A-", command=self.decrease_font).pack(side="left")
        self.font_size_var = tk.StringVar()
        self.font_entry = ttk.Entry(toolbar, textvariable=self.font_size_var, width=4, state="disabled")
        self.font_entry.pack(side="left", padx=5)
        self.font_entry.bind("<Return>", lambda e: self.set_font_size())
        ttk.Button(toolbar, text="Z+", command=lambda: self.ctrl_zoom(factor=1.1)).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Z-", command=lambda: self.ctrl_zoom(factor=0.9)).pack(side="left")
        ttk.Button(toolbar, text="Dopasuj", command=self.fit_to_window).pack(side="left", padx=5)
        ttk.Label(toolbar, text="Warstwa:").pack(side="left", padx=(5, 0))
        self.layer_var = tk.StringVar()
        self.layer_entry = ttk.Entry(toolbar, textvariable=self.layer_var, width=4, state="disabled")
        self.layer_entry.pack(side="left", padx=2)
        self.layer_entry.bind("<Return>", lambda e: self.set_layer())
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
        # store unscaled size for zoom calculations
        self.base_width = group.width / self.scale
        self.base_height = group.height / self.scale
        self.width = int(round(self.base_width * self.scale))
        self.height = int(round(self.base_height * self.scale))
        self.canvas = tk.Canvas(
            left,
            bg="white",
            width=self.width,
            height=self.height,
            scrollregion=(0, 0, self.width, self.height),
        )
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<ButtonPress-1>", self.canvas_button_press)
        self.canvas.bind("<B1-Motion>", self.canvas_drag_select)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_button_release)
        self.canvas.bind("<Control-MouseWheel>", self.ctrl_zoom)
        self.canvas_container = left
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
        cols = int(self.width / step) + 1
        rows = int(self.height / step) + 1
        for i in range(cols):
            x = i * step
            self.canvas.create_line(x, 0, x, self.height, fill="#ddd", tags="grid")
        for i in range(rows):
            y = i * step
            self.canvas.create_line(0, y, self.width, y, fill="#ddd", tags="grid")

    def clear_alignment_guides(self):
        for line in (self.align_line_h, self.align_line_v):
            if line:
                self.canvas.delete(line)
        self.align_line_h = self.align_line_v = None

    def update_alignment_guides(self, element, resize=False):
        self.clear_alignment_guides()
        others = [el for el in self.elements.values() if el is not element]
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
        return snap_dx, snap_dy

    def add_element(self, name, pos=None):
        el = DraggableElement(self, self.canvas, name, name)
        conf = self.group.field_conf.get(name)
        if conf:
            el.width = conf.get("width", el.width) * self.scale
            el.height = conf.get("height", el.height) * self.scale
            el.font_size = conf.get("font_size", el.font_size) * self.scale
            el.bold = conf.get("bold", el.bold)
            el.text_color = conf.get("text_color", el.text_color)
            el.bg_color = conf.get("bg_color", el.bg_color)
            el.bg_visible = conf.get("bg_visible", el.bg_visible)
            el.align = conf.get("align", el.align)
            el.auto_font = conf.get("auto_font", el.auto_font)
            el.layer = conf.get("layer", el.layer)
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
                el.layer = src.layer
        if pos is not None:
            el.x, el.y = pos[0] * self.scale, pos[1] * self.scale
        el.sync_canvas()
        self.elements[name] = el
        self.restack_elements()
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
            self.transparent_var.set(not self.selected_element.bg_visible)
            self.bg_check.state(["!disabled"])
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
        if self.canvas.find_withtag("current"):
            return
        self.select_element(None)
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.sel_start = (x, y)
        self.sel_rect = self.canvas.create_rectangle(
            x, y, x, y, outline="blue", dash=(2, 2), width=2
        )
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
        el.font_size += self.scale
        el.auto_font = False
        el.apply_font()
        self.font_size_var.set(str(int(el.font_size / self.scale)))

    def decrease_font(self):
        el = self.selected_element
        if not el or el.font_size <= self.scale:
            return
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
        self.layer_var.set(str(int(el.layer)))

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
        if self.selected_element:
            self.layer_var.set(str(int(self.selected_element.layer)))

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
        self.group.field_pos = {
            name: (int(round(el.x / self.scale)), int(round(el.y / self.scale)))
            for name, el in self.elements.items()
        }
        self.group.fields = list(self.group.field_pos.keys())
        self.group.conditions = list(self.conditions)
        self.group.field_conf = {
            name: {
                "width": el.width / self.scale,
                "height": el.height / self.scale,
                "font_size": el.font_size / self.scale,
                "bold": el.bold,
                "text_color": el.text_color,
                "bg_color": el.bg_color,
                "bg_visible": el.bg_visible,
                "align": el.align,
                "auto_font": el.auto_font,
                "layer": el.layer,
            }
            for name, el in self.elements.items()
        }
        self.group.sync_canvas()
        self.group.draw_preview()
        self.parent.push_history()
        self.group.editor = None
        self.destroy()

    def ctrl_zoom(self, event=None, factor=None):
        if factor is None:
            factor = 1.1 if event.delta > 0 else 0.9
        new_scale = self.scale * factor
        if new_scale <= 0:
            return
        factor = new_scale / self.scale
        for el in self.elements.values():
            el.x *= factor
            el.y *= factor
            el.width *= factor
            el.height *= factor
            el.font_size *= factor
            el.sync_canvas()
            el.apply_font()
        self.scale = new_scale
        self.snap_step = self.grid_size * self.scale
        self.width = int(round(self.base_width * self.scale))
        self.height = int(round(self.base_height * self.scale))
        self.canvas.config(
            width=self.width,
            height=self.height,
            scrollregion=(0, 0, self.width, self.height),
        )
        self.draw_grid()
        if self.selected_element:
            self.font_size_var.set(str(int(self.selected_element.font_size / self.scale)))

    def fit_to_window(self):
        container_w = self.canvas_container.winfo_width()
        container_h = self.canvas_container.winfo_height()
        if container_w <= 0 or container_h <= 0:
            return
        new_scale = min(container_w / self.base_width, container_h / self.base_height)
        self.ctrl_zoom(factor=new_scale / self.scale)
# ---------------------------------------------------------------------------
# GUI Application
# ---------------------------------------------------------------------------


