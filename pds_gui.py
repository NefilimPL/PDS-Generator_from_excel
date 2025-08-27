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
        for item in (self.rect, self.label, self.handle):
            self.canvas.move(item, dx, dy)
        self.x += dx
        self.y += dy
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
        self.width += dx
        self.height += dy
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
        return {
            "name": self.name,
            "text": self.text,
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "font_size": self.font_size,
            "bold": self.bold,
        }

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

    STATIC_FIELDS = ["Data", "Naglowek", "Stopka"]

    def __init__(self):
        super().__init__()
        self.title("PDS Generator")
        self.geometry("1200x800")
        self.excel_path = ""
        self.dataframes = {}
        self.elements = {}
        self.selected_element = None
        self.setup_ui()

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

        self.canvas = tk.Canvas(self, bg="lightgrey", width=595, height=842)
        self.canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.canvas.bind("<Configure>", lambda e: self.draw_grid())

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
            self.load_config()

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
        self.canvas.config(width=size[0], height=size[1])
        self.draw_grid()

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
            "page_size": self.canvas.winfo_width(),
            "page_height": self.canvas.winfo_height(),
            "elements": [el.to_dict() for el in self.elements.values()],
        }
        config_path = os.path.join(os.path.dirname(self.excel_path), "config.json")
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("Zapisano", f"Zapisano konfigurację do {config_path}")

    def load_config(self):
        config_path = os.path.join(os.path.dirname(self.excel_path), "config.json")
        if not os.path.exists(config_path):
            return
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception:
            return
        self.canvas.config(width=config.get("page_size", 595), height=config.get("page_height", 842))
        for elconf in config.get("elements", []):
            name = elconf["name"]
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, elconf.get("text", name))
                element.x = elconf.get("x", element.x)
                element.y = elconf.get("y", element.y)
                element.width = elconf.get("width", element.width)
                element.height = elconf.get("height", element.height)
                element.font_size = elconf.get("font_size", element.font_size)
                element.bold = elconf.get("bold", element.bold)
                self.canvas.coords(
                    element.rect,
                    element.x,
                    element.y,
                    element.x + element.width,
                    element.y + element.height,
                )
                self.canvas.coords(
                    element.label,
                    element.x + element.width / 2,
                    element.y + element.height / 2,
                )
                self.canvas.coords(
                    element.handle,
                    element.x + element.width - element.HANDLE_SIZE,
                    element.y + element.height - element.HANDLE_SIZE,
                    element.x + element.width,
                    element.y + element.height,
                )
                element.apply_font()
                element.fit_text()
                self.elements[name] = element
                # tick checkbox
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

        page_width = self.canvas.winfo_width()
        page_height = self.canvas.winfo_height()

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
                    x = element.x
                    y = page_height - element.y - element.height
                    if isinstance(value, str) and value.lower().startswith("http"):
                        try:
                            resp = requests.get(value, timeout=5)
                            img = Image.open(BytesIO(resp.content))
                            img = img.resize((int(element.width), int(element.height)))
                            c.drawImage(ImageReader(img), x, y, width=element.width, height=element.height)
                        except Exception:
                            c.drawString(x, y + element.height / 2, str(value))
                    else:
                        c.drawString(x, y + element.height / 2, str(value))
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
    def draw_grid(self):
        self.canvas.delete("grid")
        step = 20
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
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
        el.apply_font()

    def decrease_font(self):
        el = self.selected_element
        if not el:
            return
        if el.font_size > 1:
            el.font_size -= 1
            el.apply_font()


if __name__ == "__main__":
    app = PDSGeneratorGUI()
    app.mainloop()
