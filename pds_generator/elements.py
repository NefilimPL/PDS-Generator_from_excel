import logging
from io import BytesIO

import pandas as pd
import requests
from PIL import Image, ImageTk, UnidentifiedImageError
import tkinter as tk
from tkinter import font as tkfont

logger = logging.getLogger(__name__)

class DraggableElement:
    """Representation of a draggable/resizable item on the configuration canvas."""

    HANDLE_SIZE = 8

    def __init__(self, parent, canvas: tk.Canvas, name: str, text: str):
        self.parent = parent
        self.canvas = canvas
        self.name = name
        self.text = text
        self.is_image = name in getattr(parent, "image_fields", set())
        self.x = canvas.winfo_width() // 2 - 50
        self.y = canvas.winfo_height() // 2 - 20
        self.width = 100
        self.height = 40
        self.font_size = 12
        self.max_font_size = self.font_size
        self.bold = False
        self.font_family = "Arial"
        self.auto_font = True
        self.text_color = "black"
        self.bg_color = "white"
        self.bg_visible = True
        self.align = "left"
        # layering (1-based, 0 reserved for page background)
        self.layer = max((el.layer for el in parent.elements.values()), default=0) + 1
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
        self.menu.add_command(label="Przenieś warstwę +1", command=self.raise_layer)
        self.menu.add_command(label="Przenieś warstwę -1", command=self.lower_layer)
        self.canvas.tag_bind(self.rect, "<Button-3>", self.show_menu)
        self.canvas.tag_bind(self.label, "<Button-3>", self.show_menu)
        self.canvas.tag_bind(self.handle, "<Button-3>", self.show_menu)
        if hasattr(self.parent, "tooltip") and hasattr(self.parent, "get_formula_tooltip"):
            for item in (self.rect, self.label, self.handle):
                self.parent.tooltip.bind_canvas(
                    self.canvas,
                    item,
                    text_func=lambda n=self.name: self.parent.get_formula_tooltip(n),
                )
        self.apply_font()
        self.fit_text()
        self._update_label_position()

    # ------------------------------------------------------------------
    def show_menu(self, event):
        self.menu.tk_popup(event.x_root, event.y_root)

    def raise_layer(self):
        self.layer += 1
        self.parent.restack_elements()
        if getattr(self.parent, "selected_element", None) is self and hasattr(self.parent, "layer_var"):
            self.parent.layer_var.set(str(int(self.layer)))
        if hasattr(self.parent, "push_history"):
            self.parent.push_history()

    def lower_layer(self):
        if self.layer > 1:
            self.layer -= 1
            self.parent.restack_elements()
            if getattr(self.parent, "selected_element", None) is self and hasattr(self.parent, "layer_var"):
                self.parent.layer_var.set(str(int(self.layer)))
            if hasattr(self.parent, "push_history"):
                self.parent.push_history()

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
        snap_dx, snap_dy = self.parent.update_alignment_guides(self)
        if snap_dx or snap_dy:
            for el in self.parent.selected_elements:
                for item in (el.rect, el.label, el.handle, getattr(el, "image_id", None)):
                    if item:
                        el.canvas.move(item, snap_dx, snap_dy)
                el.x += snap_dx
                el.y += snap_dy
            self.last_x += snap_dx
            self.last_y += snap_dy
            self.parent.update_alignment_guides(self)

    def stop_move(self, event):
        step = self.parent.snap_step
        for el in self.parent.selected_elements:
            # snap top-left corner to the grid with integer multiples to
            # avoid sub-pixel artefacts when adjacent blocks touch
            el.x = int(round(el.x / step)) * step
            el.y = int(round(el.y / step)) * step
            # also normalise width/height so the entire block aligns to the grid
            el.width = max(step, int(round(el.width / step)) * step)
            el.height = max(step, int(round(el.height / step)) * step)
            el.sync_canvas()
        self.parent.clear_alignment_guides()
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
        if event.state & 0x0004:  # Ctrl pressed
            delta = dx if abs(dx) > abs(dy) else dy
            self.width = max(step, self.start_w + delta)
            self.height = max(step, self.start_h + delta)
        else:
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
        # normalise width/height so edges line up exactly on the grid
        self.width = max(step, int(round(self.width / step)) * step)
        self.height = max(step, int(round(self.height / step)) * step)
        self.sync_canvas()
        self.parent.clear_alignment_guides()
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
            "max_font_size": int(round(getattr(self, "max_font_size", self.font_size) / scale)),
            "bold": self.bold,
            "auto_font": self.auto_font,
            "text_color": self.text_color,
            "bg_color": self.bg_color,
            "bg_visible": self.bg_visible,
            "align": self.align,
            "layer": self.layer,
            "is_image": self.is_image,
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
        if not self.auto_font and not hasattr(self, "image_id"):
            self.canvas.itemconfig(self.label, text=self.text)
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
        except TypeError:
            if value is None:
                value = ""
        value_str = value if isinstance(value, str) else str(value)
        if self.is_image and value_str:
            if value_str.lower().startswith("http"):
                try:
                    resp = requests.get(value_str, timeout=5)
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
                    self.text = str(value_str)
                    if hasattr(self.parent, "restack_elements"):
                        self.parent.restack_elements()
                    return
                except (requests.RequestException, OSError, UnidentifiedImageError) as exc:
                    logger.exception("Failed to load remote image %s", value_str)
            local_path = self.parent.find_local_image(value_str)
            if local_path:
                try:
                    self.raw_image = Image.open(local_path)
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
                    self.text = str(value_str)
                    if hasattr(self.parent, "restack_elements"):
                        self.parent.restack_elements()
                    return
                except (OSError, UnidentifiedImageError) as exc:
                    logger.exception("Failed to load local image %s", local_path)
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
        if hasattr(self.parent, "restack_elements"):
            self.parent.restack_elements()

    def apply_font(self):
        weight = "bold" if self.bold else "normal"
        self.canvas.itemconfig(self.label, font=(self.font_family, int(self.font_size), weight))

    def _wrap_text(self, text, font, max_width):
        if text is None:
            return [""]
        text = str(text)
        if not text:
            return [""]
        lines = []
        for para in text.splitlines():
            if not para:
                lines.append("")
                continue
            words = para.split()
            if not words:
                lines.append("")
                continue
            current = words[0]
            for word in words[1:]:
                candidate = f"{current} {word}"
                if font.measure(candidate) <= max_width:
                    current = candidate
                    continue
                lines.append(current)
                if font.measure(word) <= max_width:
                    current = word
                else:
                    part = ""
                    for ch in word:
                        candidate_part = f"{part}{ch}"
                        if part and font.measure(candidate_part) > max_width:
                            lines.append(part)
                            part = ch
                        else:
                            part = candidate_part
                    current = part
            lines.append(current)
        return lines

    def fit_text(self):
        if hasattr(self, "image_id") or not self.auto_font:
            return
        max_size = int(round(getattr(self, "max_font_size", self.font_size)))
        max_size = max(1, max_size)
        weight = "bold" if self.bold else "normal"
        max_width = max(1, int(self.width - 4))
        max_height = max(1, int(self.height - 4))
        test_font = tkfont.Font(family=self.font_family, size=max_size, weight=weight)
        first_shrink = 3
        second_shrink = 3
        min_size_stage1 = max(1, max_size - first_shrink)
        min_size_stage2 = max(1, max_size - first_shrink - second_shrink)

        def split_lines(text):
            if text is None:
                return [""]
            text = str(text)
            if not text:
                return [""]
            lines = text.splitlines()
            return lines if lines else [""]

        def lines_fit_no_wrap(lines):
            line_height = test_font.metrics("linespace")
            if line_height * len(lines) > max_height:
                return False
            for line in lines:
                if test_font.measure(line) > max_width:
                    return False
            return True

        raw_lines = split_lines(self.text)
        for size in range(max_size, min_size_stage1 - 1, -1):
            test_font.configure(size=size)
            if lines_fit_no_wrap(raw_lines):
                self.font_size = size
                self.apply_font()
                self.canvas.itemconfig(self.label, text="\n".join(raw_lines))
                return

        fallback_size = min_size_stage1
        fallback_lines = self._wrap_text(self.text, test_font, max_width)
        for size in range(max_size, min_size_stage1 - 1, -1):
            test_font.configure(size=size)
            lines = self._wrap_text(self.text, test_font, max_width)
            if lines_fit_no_wrap(lines):
                self.font_size = size
                self.apply_font()
                self.canvas.itemconfig(self.label, text="\n".join(lines))
                return
            fallback_size = size
            fallback_lines = lines

        for size in range(min_size_stage1 - 1, min_size_stage2 - 1, -1):
            test_font.configure(size=size)
            lines = self._wrap_text(self.text, test_font, max_width)
            if lines_fit_no_wrap(lines):
                self.font_size = size
                self.apply_font()
                self.canvas.itemconfig(self.label, text="\n".join(lines))
                return
            fallback_size = size
            fallback_lines = lines

        self.font_size = max(1, fallback_size)
        self.apply_font()
        self.canvas.itemconfig(self.label, text="\n".join(fallback_lines))

    def update_colors(self):
        if hasattr(self, "image_id"):
            self.canvas.itemconfig(self.rect, fill="")
        else:
            self.canvas.itemconfig(self.rect, fill=self.bg_color if self.bg_visible else "")
        self.canvas.itemconfig(self.label, fill=self.text_color)

    def _update_label_position(self):
        justify = "left"
        if self.align == "left":
            justify = "left"
            self.canvas.itemconfig(self.label, anchor="w")
            self.canvas.coords(self.label, self.x + 2, self.y + self.height / 2)
        elif self.align == "right":
            justify = "right"
            self.canvas.itemconfig(self.label, anchor="e")
            self.canvas.coords(self.label, self.x + self.width - 2, self.y + self.height / 2)
        else:
            justify = "center"
            self.canvas.itemconfig(self.label, anchor="center")
            self.canvas.coords(self.label, self.x + self.width / 2, self.y + self.height / 2)
        self.canvas.itemconfig(self.label, justify=justify)


# ---------------------------------------------------------------------------
# Group areas for automatic stacking
# ---------------------------------------------------------------------------


