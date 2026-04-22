import logging
from io import BytesIO
import threading

import pandas as pd
import requests
from PIL import Image, ImageTk, UnidentifiedImageError
import tkinter as tk

from .image_auto_zoom import render_image_to_box
from .text_layout import DEFAULT_FONT_FAMILY, fit_text_lines, pdf_font_name
from .value_sources import (
    FILE_DATE_KIND_MODIFIED,
    VALUE_SOURCE_DEFAULT,
    normalize_file_date_kind,
    normalize_value_source,
)

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
        self.font_family = getattr(parent, "ui_font_family", DEFAULT_FONT_FAMILY)
        self.auto_font = True
        self.text_color = "black"
        self.bg_color = "white"
        self.bg_visible = True
        self.align = "left"
        self.image_auto_zoom = False
        self.value_source = VALUE_SOURCE_DEFAULT
        self.file_date_kind = FILE_DATE_KIND_MODIFIED
        # layering (1-based, 0 reserved for page background)
        self.layer = max((el.layer for el in parent.elements.values()), default=0) + 1
        self._image_request_id = 0
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
            "image_auto_zoom": getattr(self, "image_auto_zoom", False),
            "value_source": normalize_value_source(
                getattr(self, "value_source", VALUE_SOURCE_DEFAULT)
            ),
            "file_date_kind": normalize_file_date_kind(
                getattr(self, "file_date_kind", FILE_DATE_KIND_MODIFIED)
            ),
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
            resized = self._render_preview_image(self.raw_image)
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

    def _clear_image(self):
        if hasattr(self, "image_id"):
            self.canvas.delete(self.image_id)
            del self.image_id
        if hasattr(self, "image_obj"):
            del self.image_obj
        if hasattr(self, "raw_image"):
            del self.raw_image

    def _show_text_value(self, value):
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

    def _set_image_preview_loading(self):
        self.canvas.itemconfig(self.rect, fill=self.bg_color if self.bg_visible else "")
        self.canvas.itemconfig(
            self.label,
            text="Ładowanie...",
            fill=self.text_color,
            state="normal",
        )
        self.text = "Ładowanie..."
        self.apply_font()
        if self.auto_font:
            self.fit_text()
        self._update_label_position()
        if hasattr(self.parent, "restack_elements"):
            self.parent.restack_elements()

    def _render_preview_image(self, image):
        return render_image_to_box(
            image,
            self.width,
            self.height,
            auto_zoom=getattr(self, "image_auto_zoom", False),
        )

    def _apply_loaded_image(self, request_id, value_str, loaded_image):
        if request_id != self._image_request_id:
            return
        self._clear_image()
        self.raw_image = loaded_image
        resized = self._render_preview_image(self.raw_image)
        self.image_obj = ImageTk.PhotoImage(resized)
        self.image_id = self.canvas.create_image(
            self.x,
            self.y,
            anchor="nw",
            image=self.image_obj,
        )
        self.canvas.tag_bind(self.image_id, "<ButtonPress-1>", self.start_move)
        self.canvas.tag_bind(self.image_id, "<B1-Motion>", self.moving)
        self.canvas.tag_bind(self.image_id, "<ButtonRelease-1>", self.stop_move)
        self.canvas.tag_bind(self.image_id, "<Button-3>", self.show_menu)
        self.canvas.tag_raise(self.rect)
        self.canvas.tag_raise(self.handle)
        self.canvas.itemconfig(self.rect, fill="")
        self.canvas.itemconfig(self.label, text="", state="hidden")
        self.text = str(value_str)
        if hasattr(self.parent, "restack_elements"):
            self.parent.restack_elements()

    def _load_image_async(self, request_id, value_str):
        loaded_image = None
        if value_str.lower().startswith("http"):
            response = None
            try:
                response = requests.get(value_str, timeout=5)
                response.raise_for_status()
                with BytesIO(response.content) as buffer:
                    with Image.open(buffer) as image:
                        loaded_image = image.copy()
            except (requests.RequestException, OSError, UnidentifiedImageError):
                logger.exception("Failed to load remote image %s", value_str)
            finally:
                if response is not None:
                    try:
                        response.close()
                    except Exception:
                        pass
        else:
            local_path = self.parent.find_local_image(value_str)
            if local_path:
                try:
                    with Image.open(local_path) as image:
                        loaded_image = image.copy()
                except (OSError, UnidentifiedImageError):
                    logger.exception("Failed to load local image %s", local_path)

        def apply_result():
            if request_id != self._image_request_id:
                return
            if loaded_image is not None:
                self._apply_loaded_image(request_id, value_str, loaded_image)
            else:
                self._show_text_value(value_str)

        self._dispatch_ui(apply_result)

    def _dispatch_ui(self, func, *args, **kwargs):
        target = self.parent
        seen = set()
        while target is not None and id(target) not in seen:
            seen.add(id(target))
            ui_call = getattr(target, "ui_call", None)
            if callable(ui_call):
                ui_call(func, *args, **kwargs)
                return True
            target = getattr(target, "parent", None)
        if threading.get_ident() == threading.main_thread().ident:
            func(*args, **kwargs)
            return True
        logger.warning(
            "Skipping UI update for %s because no thread-safe ui_call dispatcher was found",
            self.name,
        )
        return False

    def update_value(self, value):
        """Update displayed value (text or image) without blocking UI."""
        self._image_request_id += 1
        request_id = self._image_request_id
        self._clear_image()
        try:
            if value is None or pd.isna(value):
                value = ""
        except TypeError:
            if value is None:
                value = ""
        value_str = value if isinstance(value, str) else str(value)
        if self.is_image and value_str:
            self._set_image_preview_loading()
            threading.Thread(
                target=self._load_image_async,
                args=(request_id, value_str),
                daemon=True,
            ).start()
            return
        self._show_text_value(value)

    def apply_font(self):
        weight = "bold" if self.bold else "normal"
        self.canvas.itemconfig(
            self.label,
            font=(self.font_family, int(round(self.font_size)), weight),
        )

    def fit_text(self):
        if hasattr(self, "image_id") or not self.auto_font:
            return
        scale = getattr(self.parent, "scale", 1.0) or 1.0
        font_name = pdf_font_name(self.bold)
        max_size = max(
            1,
            getattr(self, "max_font_size", self.font_size) / scale,
        )
        box_width = self.width / scale
        box_height = self.height / scale
        size, lines = fit_text_lines(
            self.text,
            font_name,
            max_size,
            box_width,
            box_height,
            pad=2,
        )
        self.font_size = max(1, size * scale)
        self.apply_font()
        self.canvas.itemconfig(self.label, text="\n".join(lines))

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


