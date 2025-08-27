import os
import sys
import json
import time
import threading
import subprocess
import importlib
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
import requests

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

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
        self._create_items()

    # ------------------------------------------------------------------
    def _create_items(self):
        self.rect = self.canvas.create_rectangle(
            self.x,
            self.y,
            self.x + self.width,
            self.y + self.height,
            fill="white",
            outline="black",
        )
        self.label = self.canvas.create_text(
            self.x + self.width / 2,
            self.y + self.height / 2,
            text=self.text,
        )
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
        self.canvas.tag_bind(self.label, "<ButtonPress-1>", self.start_move)
        self.canvas.tag_bind(self.label, "<B1-Motion>", self.moving)
        self.canvas.tag_bind(self.handle, "<ButtonPress-1>", self.start_resize)
        self.canvas.tag_bind(self.handle, "<B1-Motion>", self.resizing)
        # Context menu for layering
        self.menu = tk.Menu(self.canvas, tearoff=0)
        self.menu.add_command(label="Przenieś na wierzch", command=self.bring_to_front)
        self.menu.add_command(label="Przenieś na spód", command=self.send_to_back)
        self.canvas.tag_bind(self.rect, "<Button-3>", self.show_menu)
        self.canvas.tag_bind(self.label, "<Button-3>", self.show_menu)
        self.canvas.tag_bind(self.handle, "<Button-3>", self.show_menu)

        self.apply_font()
        self.fit_text()

    # ------------------------------------------------------------------
    def show_menu(self, event):
        self.menu.tk_popup(event.x_root, event.y_root)

    def bring_to_front(self):
        for item in (self.rect, self.label, self.handle):
            self.canvas.tag_raise(item)

    def send_to_back(self):
        for item in (self.rect, self.label, self.handle):
            self.canvas.tag_lower(item)

    # ------------------------------------------------------------------
    def start_move(self, event):
        self.parent.select_element(self)
        self.last_x = event.x
        self.last_y = event.y

    def moving(self, event):
        dx = event.x - self.last_x
        dy = event.y - self.last_y
        new_x = self.x + dx
        new_y = self.y + dy
        step = self.parent.grid_size * self.parent.scale
        new_x = round(new_x / step) * step
        new_y = round(new_y / step) * step
        dx = new_x - self.x
        dy = new_y - self.y
        for item in (self.rect, self.label, self.handle):
            self.canvas.move(item, dx, dy)
        self.x = new_x
        self.y = new_y
        self.last_x = event.x
        self.last_y = event.y

    # ------------------------------------------------------------------
    def start_resize(self, event):
        self.parent.select_element(self)
        self.last_x = event.x
        self.last_y = event.y

    def resizing(self, event):
        dx = event.x - self.last_x
        dy = event.y - self.last_y
        step = self.parent.grid_size * self.parent.scale
        new_width = max(step, self.width + dx)
        new_height = max(step, self.height + dy)
        new_width = round(new_width / step) * step
        new_height = round(new_height / step) * step
        self.width = new_width
        self.height = new_height
        self.last_x = event.x
        self.last_y = event.y
        self.canvas.coords(
            self.rect,
            self.x,
            self.y,
            self.x + self.width,
            self.y + self.height,
        )
        self.canvas.coords(
            self.label,
            self.x + self.width / 2,
            self.y + self.height / 2,
        )
        self.canvas.coords(
            self.handle,
            self.x + self.width - self.HANDLE_SIZE,
            self.y + self.height - self.HANDLE_SIZE,
            self.x + self.width,
            self.y + self.height,
        )
        self.fit_text()

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
        }

    def sync_canvas(self):
        self.canvas.coords(
            self.rect,
            self.x,
            self.y,
            self.x + self.width,
            self.y + self.height,
        )
        self.canvas.coords(
            self.label,
            self.x + self.width / 2,
            self.y + self.height / 2,
        )
        self.canvas.coords(
            self.handle,
            self.x + self.width - self.HANDLE_SIZE,
            self.y + self.height - self.HANDLE_SIZE,
            self.x + self.width,
            self.y + self.height,
        )
        self.apply_font()
        self.fit_text()

    def update_value(self, value):
        """Update displayed value (text or image)."""
        # Remove previous image if any
        if hasattr(self, "image_id"):
            self.canvas.delete(self.image_id)
            del self.image_id
            if hasattr(self, "image_obj"):
                del self.image_obj
        if value is None:
            value = ""
        if isinstance(value, str) and value.lower().startswith("http"):
            try:
                resp = requests.get(value, timeout=5)
                img = Image.open(BytesIO(resp.content))
                img = img.resize((int(self.width), int(self.height)))
                self.image_obj = ImageTk.PhotoImage(img)
                self.image_id = self.canvas.create_image(
                    self.x,
                    self.y,
                    anchor="nw",
                    image=self.image_obj,
                )
                self.canvas.tag_raise(self.rect)
                self.canvas.tag_raise(self.handle)
                self.canvas.delete(self.label)
                return
            except Exception:
                pass
        # default: text
        self.canvas.delete(self.label)
        self.label = self.canvas.create_text(
            self.x + self.width / 2,
            self.y + self.height / 2,
            text=str(value),
        )
        self.canvas.tag_bind(self.label, "<ButtonPress-1>", self.start_move)
        self.canvas.tag_bind(self.label, "<B1-Motion>", self.moving)
        self.canvas.tag_bind(self.label, "<Button-3>", self.show_menu)
        self.text = str(value)
        self.apply_font()
        self.fit_text()

    def apply_font(self):
        weight = "bold" if self.bold else "normal"
        self.canvas.itemconfig(self.label, font=(self.font_family, int(self.font_size), weight))

    def fit_text(self):
        if hasattr(self, "image_id"):
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


# ---------------------------------------------------------------------------
# GUI Application
# ---------------------------------------------------------------------------


class PDSGeneratorGUI(tk.Tk):
    PAGE_SIZES = {
        "A4": (595, 842),  # 210 x 297 mm in points
        "B5": (516, 729),  # 176 x 250 mm
    }

    grid_size = 20

    STATIC_FIELDS = ["Data", "Naglowek", "Stopka"]

    def __init__(self):
        super().__init__()
        self.title("PDS Generator")
        self.geometry("1200x800")
        self.excel_path = ""
        self.dataframes = {}
        self.elements = {}
        self.selected_element = None
        self.page_width, self.page_height = self.PAGE_SIZES["A4"]
        self.scale = 1.0
        self.setup_ui()
        self.update_idletasks()
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
        sizes = list(self.PAGE_SIZES.keys())
        self.size_combo = ttk.Combobox(top_frame, textvariable=self.size_var, values=sizes)
        self.size_combo.pack(side="left")
        self.size_combo.bind("<<ComboboxSelected>>", lambda e: self.update_canvas_size())

        format_frame = ttk.Frame(self)
        format_frame.pack(fill="x", padx=5)
        ttk.Button(format_frame, text="B", command=self.toggle_bold).pack(side="left")
        ttk.Button(format_frame, text="A+", command=self.increase_font).pack(side="left", padx=2)
        ttk.Button(format_frame, text="A-", command=self.decrease_font).pack(side="left")
        self.font_size_var = tk.StringVar()
        self.font_entry = ttk.Entry(format_frame, textvariable=self.font_size_var, width=4, state="disabled")
        self.font_entry.pack(side="left", padx=5)
        self.font_entry.bind("<Return>", lambda e: self.set_font_size())

        self.canvas_container = ttk.Frame(self)
        self.canvas_container.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.canvas_container.bind("<Configure>", self.resize_canvas)
        self.canvas = tk.Canvas(self.canvas_container, bg="lightgrey", width=self.page_width, height=self.page_height)
        self.canvas.pack(expand=True)

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
        for field in self.STATIC_FIELDS:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(
                self.static_frame,
                text=field,
                variable=var,
                command=lambda f=field, v=var: self.toggle_static(f, v.get()),
            )
            chk.pack(anchor="w")
            self.static_vars[field] = var

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
        ttk.Button(button_frame, text="Generuj PDS", command=self.generate_pds).pack(fill="x", pady=5)

        # Progress bar
        self.progress = ttk.Progressbar(right_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(20, 0))
        self.time_label = ttk.Label(right_frame, text="")
        self.time_label.pack()
        self.draw_grid()

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
        value = self.size_var.get()
        if "x" in value:
            try:
                w, h = value.split("x")
                size = (int(float(w)), int(float(h)))
            except Exception:
                messagebox.showerror("Błąd", "Nieprawidłowy format rozmiaru. Użyj np. 595x842")
                return
        else:
            size = self.PAGE_SIZES.get(value, self.PAGE_SIZES["A4"])
        factor_w = size[0] / self.page_width
        factor_h = size[1] / self.page_height
        self.page_width, self.page_height = size
        for el in self.elements.values():
            el.x *= factor_w
            el.y *= factor_h
            el.width *= factor_w
            el.height *= factor_h
            el.font_size *= factor_h
            el.sync_canvas()
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
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, name)
                self.elements[name] = element
        else:
            self.remove_element(name)

    def remove_element(self, name):
        element = self.elements.pop(name, None)
        if element:
            for item in (element.rect, element.label, element.handle):
                self.canvas.delete(item)
            if hasattr(element, "image_id"):
                self.canvas.delete(element.image_id)
            if self.selected_element is element:
                self.selected_element = None
                self.font_entry.configure(state="disabled")
                self.font_size_var.set("")

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
                value = name
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
        self.size_var.set(f"{int(self.page_width)}x{int(self.page_height)}")
        self.resize_canvas()
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
                element.sync_canvas()
                self.elements[name] = element
                if name in self.columns_vars:
                    self.columns_vars[name].set(True)
                if name in self.static_vars:
                    self.static_vars[name].set(True)

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
                c = pdf_canvas.Canvas(pdf_path, pagesize=(page_width, page_height))
                for name, element in self.elements.items():
                    if ":" in name:
                        sheet, col = name.split(":", 1)
                        df = self.dataframes.get(sheet)
                        value = df.iloc[idx].get(col, "") if df is not None else ""
                    else:
                        value = name
                    x = element.x / self.scale
                    y = page_height - (element.y / self.scale) - (element.height / self.scale)
                    if isinstance(value, str) and value.lower().startswith("http"):
                        try:
                            resp = requests.get(value, timeout=5)
                            img = Image.open(BytesIO(resp.content))
                            img = img.resize((int(element.width / self.scale), int(element.height / self.scale)))
                            c.drawImage(ImageReader(img), x, y, width=element.width / self.scale, height=element.height / self.scale)
                        except Exception:
                            c.setFont("Helvetica-Bold" if element.bold else "Helvetica", element.font_size / self.scale)
                            c.drawString(x, y + (element.height / self.scale) / 2, str(value))
                    else:
                        c.setFont("Helvetica-Bold" if element.bold else "Helvetica", element.font_size / self.scale)
                        c.drawString(x, y + (element.height / self.scale) / 2, str(value))
                c.showPage()
                c.save()
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
        new_scale = min(container_w / self.page_width, container_h / self.page_height)
        if new_scale <= 0:
            return
        factor = new_scale / self.scale
        self.canvas.config(width=self.page_width * new_scale, height=self.page_height * new_scale)
        self.canvas.scale("all", 0, 0, factor, factor)
        for el in self.elements.values():
            el.x *= factor
            el.y *= factor
            el.width *= factor
            el.height *= factor
            el.font_size *= factor
            el.apply_font()
        self.scale = new_scale
        if self.selected_element:
            self.font_size_var.set(str(int(self.selected_element.font_size / self.scale)))
        self.draw_grid()

    def draw_grid(self):
        self.canvas.delete("grid")
        step = int(self.grid_size * self.scale)
        if step <= 0:
            return
        w = int(self.page_width * self.scale)
        h = int(self.page_height * self.scale)
        for i in range(0, w, step):
            self.canvas.create_line(i, 0, i, h, fill="#d0d0d0", tags="grid")
        for i in range(0, h, step):
            self.canvas.create_line(0, i, w, i, fill="#d0d0d0", tags="grid")
        self.canvas.tag_lower("grid")

    def select_element(self, element):
        if self.selected_element and self.selected_element is not element:
            self.canvas.itemconfig(self.selected_element.rect, outline="black")
        self.selected_element = element
        self.canvas.itemconfig(element.rect, outline="red")
        if element:
            self.font_entry.configure(state="normal")
            self.font_size_var.set(str(int(element.font_size / self.scale)))
        else:
            self.font_entry.configure(state="disabled")
            self.font_size_var.set("")

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
        el.apply_font()
        self.font_size_var.set(str(int(el.font_size / self.scale)))

    def decrease_font(self):
        el = self.selected_element
        if not el:
            return
        if el.font_size > self.scale:
            el.font_size -= self.scale
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
        el.apply_font()


if __name__ == "__main__":
    app = PDSGeneratorGUI()
    app.mainloop()
