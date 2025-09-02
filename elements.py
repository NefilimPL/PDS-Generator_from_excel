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
            "bold": self.bold,
            "auto_font": self.auto_font,
            "text_color": self.text_color,
            "bg_color": self.bg_color,
            "bg_visible": self.bg_visible,
            "align": self.align,
            "layer": self.layer,
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
        except TypeError:
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
                if hasattr(self.parent, "restack_elements"):
                    self.parent.restack_elements()
                return
            except (requests.RequestException, OSError, UnidentifiedImageError) as exc:
                logger.exception("Failed to load remote image %s", value)
        if isinstance(value, str):
            local_path = self.parent.find_local_image(value)
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
                    self.text = str(value)
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


