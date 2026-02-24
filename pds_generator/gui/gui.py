import logging
import os
import sys
import webbrowser
import threading
from copy import deepcopy

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter import font as tkfont
from PIL import Image, ImageTk

from ..elements import DraggableElement
from ..groups import GroupArea, GroupEditor

from .ui_layout import setup_ui as build_ui
from .tooltips import Tooltip
from .pdf_export import (
    generate_pds as export_pds,
    draw_pdf_element as render_pdf_element,
)
from .config_io import (
    save_config as save_config_func,
    load_config as load_config_func,
)
from . import mailer
from . import windows_auth
from .excel_tracking import is_tracking_column
from ..excel_io import read_excel_data, detect_formula_columns, collect_formula_samples
from . import locks

from ..number_format import round_numeric_value

from ..github_utils import (
    get_repo_info,
    get_remote_commit_info,
    get_remote_version,
    pull_updates,
    get_version,
)
from .. import image_index as image_index_utils

logger = logging.getLogger(__name__)

IMAGE_EXTENSIONS = (
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
    ".webp",
    ".ico",
)

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
        self.repo_dir = os.path.abspath(
            os.path.join(os.path.dirname(__file__), "..", "..")
        )
        self.version = get_version(self.repo_dir)
        self.remote_version = None
        self.remote_date = None
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
        self.conditions_win = None
        self.image_cache = {}
        self.image_fields = set()
        self.image_vars = {}
        self.image_dirs = []
        self.image_index_data = None
        self.image_index_roots = ()
        self.image_index_lock = threading.Lock()
        self.excel_lock_path = None
        self.config_lock_path = None
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
        self.generation_in_progress = False
        self.cancel_event = None
        self.tracking_excluded = set()
        self.mail_config = mailer.normalize_mail_config({})
        self.require_admin_for_mail_settings = (
            str(os.getenv("PDS_REQUIRE_ADMIN_MAIL_SETTINGS", "1")).strip().lower()
            not in {"0", "false", "no"}
        )
        self.preview_in_progress = False
        self.preview_animation_after = None
        self.preview_animation_step = 0
        self.mail_settings_win = None
        self.last_generation_report = None
        self.formula_map = {}
        self.column_rows = {}
        self.column_indicators = {}
        self.static_indicators = {}
        self._highlighted_fields = set()
        self.panel_bg = self.cget("background")
        self.highlight_color = "#ffd46a"
        self.tooltip = Tooltip(self)
        self.setup_ui()
        self.bind_all("<Control-z>", self.undo)
        self.bind_all("<Control-x>", self.redo)
        self.update_idletasks()
        self.resize_canvas()
        self.load_config(startup=True)
        self.check_for_updates()
        if not self.history:
            self.push_history()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------------
    def setup_ui(self):
        build_ui(self)

    # ------------------------------------------------------------------
    def check_for_updates(self):
        local_hash, owner, repo = get_repo_info(self.repo_dir)
        self.repo_owner, self.repo_name = owner, repo

        remote_hash = None
        remote_date = None
        remote_version = None
        if local_hash:
            remote_hash, remote_date = get_remote_commit_info(owner, repo)
            self.update_available = bool(remote_hash and remote_hash != local_hash)
            if self.update_available:
                remote_version = get_remote_version(owner, repo)
        else:
            remote_version = get_remote_version(owner, repo)
            self.update_available = bool(remote_version and remote_version != self.version)
            if self.update_available:
                _, remote_date = get_remote_commit_info(owner, repo)

        self.remote_version = remote_version
        self.remote_date = remote_date

        if self.update_available:
            if remote_version:
                info = f"Aktualna wersja: {self.version} | Dostępna: {remote_version}"
                if remote_date:
                    info += f" ({remote_date})"
            else:
                info = f"Aktualna wersja: {self.version}"
            self.update_info_var.set(info)
            self.update_button.pack(side="left", padx=5)
            self.blink_update_button()
        else:
            self.update_info_var.set(f"Aktualna wersja: {self.version}")
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
                self._run_update()

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
        self._run_update()

    def _run_update(self):
        previous_excel_lock = self.excel_lock_path
        previous_config_lock = self.config_lock_path

        self.release_lock("excel_lock_path")
        self.release_lock("config_lock_path")

        try:
            updated = pull_updates(self.repo_dir)
        except Exception:  # pragma: no cover - defensive logging
            logger.exception("Unexpected error while pulling updates")
            updated = False

        if not updated:
            messagebox.showerror("Błąd", "Aktualizacja nie powiodła się")
            self._restore_locks_after_failed_update(
                previous_excel_lock, previous_config_lock
            )
            return

        python = sys.executable
        os.execl(python, python, *sys.argv)

    def _restore_locks_after_failed_update(
        self, excel_lock_path, config_lock_path
    ):
        if excel_lock_path and self.excel_path:
            if not self.acquire_excel_lock(self.excel_path):
                logger.warning(
                    "Failed to reacquire Excel lock for %s", self.excel_path
                )

        if config_lock_path:
            resource_path = (
                config_lock_path[:-5]
                if str(config_lock_path).endswith(".lock")
                else None
            )
            if resource_path:
                lock = locks.acquire_lock(
                    resource_path, os.path.basename(resource_path)
                )
                if lock:
                    self.config_lock_path = lock
                else:
                    logger.warning(
                        "Failed to reacquire config lock for %s", resource_path
                    )

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
            if not self.acquire_excel_lock(path):
                return
            self.path_var.set(path)
            self.excel_path = path
            self.image_cache = {}
            self._invalidate_image_index()
            self.load_excel(path)
            self.load_config(path=path)

    def load_excel(self, path):
        self._invalidate_image_index()
        try:
            self.dataframes = read_excel_data(path)
        except (OSError, ValueError) as e:
            logger.exception("Failed to read Excel file %s", path)
            messagebox.showerror("Błąd", f"Nie można wczytać Excela: {e}")
            return

        self.formula_map = {}
        formula_samples = collect_formula_samples(path)
        for sheet, cols in formula_samples.items():
            for col, formulas in cols.items():
                self.formula_map[f"{sheet}:{col}"] = formulas

        formula_cols = detect_formula_columns(path)
        missing_formula_cols = {}
        for sheet, cols in formula_cols.items():
            df = self.dataframes.get(sheet)
            if df is None:
                continue
            for col in cols:
                if col not in df.columns:
                    continue
                series = df[col].head(200)
                has_value = False
                for val in series:
                    try:
                        if pd.isna(val):
                            continue
                    except Exception:
                        pass
                    if val != "":
                        has_value = True
                        break
                if not has_value:
                    missing_formula_cols.setdefault(sheet, []).append(col)
        if missing_formula_cols:
            preview = []
            for sheet, cols in missing_formula_cols.items():
                for col in cols:
                    preview.append(f"{sheet}:{col}")
                    if len(preview) >= 6:
                        break
                if len(preview) >= 6:
                    break
            logger.warning(
                "Formula columns without cached values: %s",
                ", ".join(preview),
            )
            messagebox.showwarning(
                "Uwaga",
                "Wykryto kolumny z formułami bez zapisanych wyników. "
                "Otwórz plik w Excelu, odśwież obliczenia i zapisz, "
                "a następnie wczytaj ponownie. "
                f"Przykłady: {', '.join(preview)}",
            )
        # Clear previous
        for child in self.columns_frame.winfo_children():
            child.destroy()
        old_column_keys = list(self.columns_vars.keys())
        self.columns_vars.clear()
        self.column_rows.clear()
        self.column_indicators.clear()
        for key in old_column_keys:
            self.image_vars.pop(key, None)

        for sheet, df in self.dataframes.items():
            lf = ttk.LabelFrame(self.columns_frame, text=sheet)
            lf.pack(fill="x", padx=2, pady=2)
            for col in df.columns:
                if is_tracking_column(col):
                    continue
                name = f"{sheet}:{col}"
                row = ttk.Frame(lf)
                row.pack(fill="x", padx=2, pady=1)
                indicator = tk.Frame(row, width=6, height=18, bg=self.panel_bg)
                indicator.pack(side="left", fill="y", padx=(0, 4))
                indicator.pack_propagate(False)
                var = tk.BooleanVar()
                chk = ttk.Checkbutton(
                    row,
                    text=col,
                    variable=var,
                    command=lambda s=sheet, c=col, v=var: self.toggle_column(f"{s}:{c}", v.get()),
                )
                chk.pack(side="left", anchor="w")
                img_var = tk.BooleanVar(value=name in self.image_fields)
                img_chk = ttk.Checkbutton(
                    row,
                    text="IMG",
                    variable=img_var,
                    command=lambda n=name, v=img_var: self.set_field_image(n, v.get()),
                )
                img_chk.pack(side="right")
                self.columns_vars[name] = var
                self.image_vars[name] = img_var
                self.column_rows[name] = row
                self.column_indicators[name] = indicator
                self.bind_formula_tooltip(row, name)
                self.bind_formula_tooltip(chk, name)
                if hasattr(self, "tooltip"):
                    self.tooltip.bind(img_chk, text="Traktuj wartość jako obraz")

        self.update_field_highlights()

    # ------------------------------------------------------------------
    def _image_search_roots(self):
        base_dir = os.path.dirname(self.excel_path) if self.excel_path else ""
        roots = [base_dir] + list(getattr(self, "image_dirs", []) or [])
        return image_index_utils.normalize_roots(roots)

    def _invalidate_image_index(self):
        with self.image_index_lock:
            self.image_index_data = None
            self.image_index_roots = ()

    def _ensure_image_index(self, force_rebuild=False, max_age_hours=24):
        roots = self._image_search_roots()
        roots_key = tuple(roots)
        with self.image_index_lock:
            cached = self.image_index_data
            cached_roots = self.image_index_roots
        if (
            cached
            and not force_rebuild
            and cached_roots == roots_key
        ):
            return cached
        index_data = image_index_utils.ensure_index(
            roots,
            force_rebuild=force_rebuild,
            max_age_hours=max_age_hours,
        )
        with self.image_index_lock:
            self.image_index_data = index_data
            self.image_index_roots = roots_key
        return index_data

    def rebuild_image_index(self):
        roots = self._image_search_roots()
        if not roots:
            messagebox.showinfo(
                "Indeks obrazów",
                "Brak katalogów obrazów do zindeksowania.",
            )
            return

        if hasattr(self, "image_index_btn") and self.image_index_btn:
            self.image_index_btn.state(["disabled"])
        if hasattr(self, "image_index_status_var") and self.image_index_status_var is not None:
            self.image_index_status_var.set("Indeksowanie...")
        self.set_status("Budowanie indeksu obrazów...")

        def worker():
            try:
                started = time.time()
                index_data = self._ensure_image_index(
                    force_rebuild=True,
                    max_age_hours=0,
                )
                elapsed = max(0.0, time.time() - started)
                count = int(index_data.get("file_count", 0))

                def on_success():
                    if hasattr(self, "image_index_btn") and self.image_index_btn:
                        self.image_index_btn.state(["!disabled"])
                    if hasattr(self, "image_index_status_var") and self.image_index_status_var is not None:
                        self.image_index_status_var.set(
                            f"Indeks: {count} plików ({elapsed:.1f}s)"
                        )
                    self.set_status("Indeks obrazów gotowy")

                self.ui_call(on_success)
            except Exception as exc:
                logger.exception("Failed to rebuild image index")

                def on_error():
                    if hasattr(self, "image_index_btn") and self.image_index_btn:
                        self.image_index_btn.state(["!disabled"])
                    if hasattr(self, "image_index_status_var") and self.image_index_status_var is not None:
                        self.image_index_status_var.set("Błąd indeksu")
                    self.set_status("Błąd indeksu obrazów")
                    messagebox.showerror(
                        "Indeks obrazów",
                        f"Nie udało się zbudować indeksu: {exc}",
                    )

                self.ui_call(on_error)

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    def find_local_image(self, filename):
        """Search for an image file relative to the Excel file directory."""
        if not filename:
            return None
        if not getattr(self, "excel_path", "") and not getattr(self, "image_dirs", None):
            return None
        name = str(filename).strip()
        if not name:
            return None
        key = name.lower()
        if key in self.image_cache:
            return self.image_cache[key]

        search_roots = self._image_search_roots()

        stem, _ext = os.path.splitext(os.path.basename(name))
        stem_lower = stem.lower()

        path = None
        if os.path.isabs(name):
            if os.path.isfile(name):
                path = name
            elif stem:
                abs_dir = os.path.dirname(name)
                for ext in IMAGE_EXTENSIONS:
                    candidate = os.path.join(abs_dir, stem + ext)
                    if os.path.isfile(candidate):
                        path = candidate
                        break
        else:
            for root in search_roots:
                candidate = os.path.join(root, name)
                if os.path.isfile(candidate):
                    path = candidate
                    break
                if stem:
                    rel_dir = os.path.dirname(name)
                    if rel_dir:
                        for ext in IMAGE_EXTENSIONS:
                            candidate = os.path.join(root, rel_dir, stem + ext)
                            if os.path.isfile(candidate):
                                path = candidate
                                break
                        if path:
                            break
                    for ext in IMAGE_EXTENSIONS:
                        candidate = os.path.join(root, stem + ext)
                        if os.path.isfile(candidate):
                            path = candidate
                            break
                    if path:
                        break

        if path is None:
            try:
                index_data = self._ensure_image_index(force_rebuild=False, max_age_hours=24)
            except Exception:
                logger.exception("Failed to load/build image index")
                index_data = None
            indexed_path = image_index_utils.find_in_index(index_data, name)
            if indexed_path:
                path = indexed_path

        if path is None and stem and not self.image_index_data:
            target_name = os.path.basename(name).lower()
            for root in search_roots:
                for current_root, _dirs, files in os.walk(root):
                    for f in files:
                        f_lower = f.lower()
                        if f_lower == target_name:
                            path = os.path.join(current_root, f)
                            break
                        file_stem, file_ext = os.path.splitext(f_lower)
                        if file_stem == stem_lower and file_ext in IMAGE_EXTENSIONS:
                            path = os.path.join(current_root, f)
                            break
                    if path:
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
            if hasattr(el, "max_font_size"):
                el.max_font_size *= factor_h
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
                element.is_image = name in self.image_fields
                self.elements[name] = element
                self.restack_elements()
        else:
            self.remove_element(name)
        self.push_history()

    def set_field_image(self, name, state):
        if state:
            self.image_fields.add(name)
        else:
            self.image_fields.discard(name)
        var = self.image_vars.get(name)
        if var is not None and var.get() != state:
            var.set(state)
        element = self.elements.get(name)
        if element:
            element.is_image = state
            value = self._resolve_preview_value(name)
            element.update_value(value)
        self.image_cache = {}
        self.push_history()

    def _resolve_preview_value(self, name):
        if ":" in name:
            if not self.dataframes:
                return ""
            try:
                idx = int(self.row_var.get()) - 1
            except (ValueError, AttributeError):
                idx = 0
            sheet, col = name.split(":", 1)
            df = self.dataframes.get(sheet)
            value = None
            if df is not None and 0 <= idx < len(df):
                value = df.iloc[idx].get(col)
                value = round_numeric_value(value)
            return value
        if name in getattr(self, "static_entries", {}):
            return self.static_entries[name].get()
        return name

    def apply_image_field_state(self):
        for name, var in self.image_vars.items():
            var.set(name in self.image_fields)
        for name, element in self.elements.items():
            element.is_image = name in self.image_fields

    def add_image_dir(self):
        path = filedialog.askdirectory()
        if not path:
            return
        path = os.path.abspath(path)
        if path in self.image_dirs:
            return
        self.image_dirs.append(path)
        self.refresh_image_dir_list()
        self.image_cache = {}
        self._invalidate_image_index()
        self.push_history()

    def remove_image_dir(self):
        if not hasattr(self, "image_dirs_list"):
            return
        selection = list(self.image_dirs_list.curselection())
        if not selection:
            return
        for idx in sorted(selection, reverse=True):
            if 0 <= idx < len(self.image_dirs):
                self.image_dirs.pop(idx)
        self.refresh_image_dir_list()
        self.image_cache = {}
        self._invalidate_image_index()
        self.push_history()

    def refresh_image_dir_list(self):
        if not hasattr(self, "image_dirs_list"):
            return
        self.image_dirs_list.delete(0, "end")
        for path in self.image_dirs:
            self.image_dirs_list.insert("end", path)

    def _collect_dynamic_fields(self):
        dynamic_fields = set()

        def collect(name):
            if isinstance(name, str) and ":" in name:
                dynamic_fields.add(name)

        for name in self.elements.keys():
            collect(name)
        for group in self.groups.values():
            for fname in group.fields:
                collect(fname)
            for src, tgt in group.conditions:
                collect(src)
                collect(tgt)
        for src, tgt in self.conditions:
            collect(src)
            collect(tgt)
        return dynamic_fields

    def open_tracking_settings(self):
        dynamic_fields = self._collect_dynamic_fields()
        if not dynamic_fields:
            messagebox.showinfo(
                "Śledzenie zmian",
                "Brak pól z Excela w szablonie. Najpierw dodaj kolumny.",
            )
            return

        win = tk.Toplevel(self)
        win.title("Śledzenie zmian")
        ttk.Label(
            win,
            text="Zaznaczone pola będą brane pod uwagę przy wykrywaniu zmian.",
        ).pack(padx=10, pady=5)

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=5)
        listbox = tk.Listbox(frame, selectmode="multiple", height=12)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scroll.set)
        listbox.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        fields = sorted(dynamic_fields)
        for idx, field in enumerate(fields):
            listbox.insert("end", field)
            if field not in self.tracking_excluded:
                listbox.selection_set(idx)

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=10, pady=5)

        def select_all():
            listbox.selection_set(0, "end")

        def clear_all():
            listbox.selection_clear(0, "end")

        def save():
            selected = {listbox.get(i) for i in listbox.curselection()}
            self.tracking_excluded = set(fields) - selected
            win.destroy()

        ttk.Button(btns, text="Zaznacz wszystkie", command=select_all).pack(side="left")
        ttk.Button(btns, text="Wyczyść", command=clear_all).pack(side="left", padx=5)
        ttk.Button(btns, text="Zapisz", command=save).pack(side="right")

    def _mail_settings_from_inputs(
        self,
        transport,
        enabled,
        smtp_host,
        smtp_port,
        smtp_security,
        username,
        password,
        sender,
        entra_token,
        entra_tenant_id,
        entra_client_id,
        entra_client_secret,
        entra_sender,
        entra_endpoint,
        secret_key_id,
        recipients_text,
        subject_prefix,
        timeout_seconds,
        strict=False,
        require_recipients=False,
    ):
        transport_value = str(transport or "").strip().lower()
        if transport_value == mailer.TRANSPORT_SMTP:
            try:
                port = int(str(smtp_port).strip())
            except (TypeError, ValueError):
                messagebox.showerror("E-mail", "Port SMTP musi być liczbą całkowitą.")
                return None
        else:
            try:
                port = int(str(smtp_port).strip())
            except (TypeError, ValueError):
                port = mailer.DEFAULT_MAIL_CONFIG["smtp_port"]
        try:
            timeout = int(str(timeout_seconds).strip())
        except (TypeError, ValueError):
            messagebox.showerror("E-mail", "Timeout musi być liczbą całkowitą.")
            return None

        cfg = mailer.normalize_mail_config(
            {
                "transport": transport,
                "enabled": bool(enabled),
                "smtp_host": smtp_host,
                "smtp_port": port,
                "smtp_security": smtp_security,
                "username": username,
                "password": password,
                "sender": sender,
                "entra_token": entra_token,
                "entra_tenant_id": entra_tenant_id,
                "entra_client_id": entra_client_id,
                "entra_client_secret": entra_client_secret,
                "entra_sender": entra_sender,
                "entra_endpoint": entra_endpoint,
                "secret_key_id": secret_key_id,
                "recipients": recipients_text,
                "subject_prefix": subject_prefix,
                "timeout_seconds": timeout,
            }
        )
        if strict:
            try:
                cfg = mailer.validate_mail_config(
                    cfg, require_recipients=require_recipients
                )
            except ValueError as exc:
                messagebox.showerror("E-mail", str(exc))
                return None
        return cfg

    def _run_mail_action_async(
        self,
        config,
        action,
        success_message,
        start_status,
        success_status,
        error_prefix,
    ):
        self.set_status(start_status)

        def worker():
            try:
                action(config)
            except Exception as exc:
                logger.exception("%s failed", error_prefix)
                self.ui_call(
                    messagebox.showerror,
                    "E-mail",
                    f"{error_prefix}: {exc}",
                )
                self.ui_call(self.set_status, "Błąd e-mail")
                return
            self.ui_call(messagebox.showinfo, "E-mail", success_message)
            self.ui_call(self.set_status, success_status)

        threading.Thread(target=worker, daemon=True).start()

    def open_mail_settings(self):
        if self.mail_settings_win and self.mail_settings_win.winfo_exists():
            self.mail_settings_win.lift()
            self.mail_settings_win.focus_force()
            return

        if self.require_admin_for_mail_settings and os.name == "nt":
            ok, err = windows_auth.prompt_admin_credentials(
                parent_hwnd=self.winfo_id(),
                caption="Autoryzacja administratora",
                message=(
                    "Aby otworzyć konfigurację e-mail, zaloguj się kontem administratora."
                ),
            )
            if not ok:
                if err and err != "cancelled":
                    messagebox.showerror("Uprawnienia", err)
                self.set_status("Odmowa dostępu do konfiguracji e-mail")
                return

        cfg = mailer.normalize_mail_config(getattr(self, "mail_config", {}))

        win = tk.Toplevel(self)
        self.mail_settings_win = win
        win.title("Konfiguracja e-mail")
        win.columnconfigure(1, weight=1)

        enabled_var = tk.BooleanVar(value=cfg["enabled"])
        transport_var = tk.StringVar(value=cfg["transport"])
        host_var = tk.StringVar(value=cfg["smtp_host"])
        port_var = tk.StringVar(value=str(cfg["smtp_port"]))
        security_var = tk.StringVar(value=cfg["smtp_security"])
        username_var = tk.StringVar(value=cfg["username"])
        password_var = tk.StringVar(value=cfg["password"])
        sender_var = tk.StringVar(value=cfg["sender"])
        entra_token_var = tk.StringVar(value=cfg["entra_token"])
        entra_tenant_var = tk.StringVar(value=cfg["entra_tenant_id"])
        entra_client_var = tk.StringVar(value=cfg["entra_client_id"])
        entra_client_secret_var = tk.StringVar(value=cfg["entra_client_secret"])
        entra_sender_var = tk.StringVar(value=cfg["entra_sender"])
        entra_endpoint_var = tk.StringVar(value=cfg["entra_endpoint"])
        secret_key_id_var = tk.StringVar(value=cfg.get("secret_key_id", ""))
        subject_var = tk.StringVar(value=cfg["subject_prefix"])
        timeout_var = tk.StringVar(value=str(cfg["timeout_seconds"]))
        show_sensitive_var = tk.BooleanVar(value=False)

        ttk.Checkbutton(
            win,
            text="Włącz wysyłkę raportów po generowaniu PDF",
            variable=enabled_var,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 8))

        ttk.Label(win, text="Metoda wysyłki:").grid(
            row=1, column=0, sticky="w", padx=10, pady=2
        )
        transport_frame = ttk.Frame(win)
        transport_frame.grid(row=1, column=1, sticky="w", padx=10, pady=2)
        ttk.Radiobutton(
            transport_frame,
            text="SMTP",
            value=mailer.TRANSPORT_SMTP,
            variable=transport_var,
        ).pack(side="left")
        ttk.Radiobutton(
            transport_frame,
            text="Microsoft Entra API (token)",
            value=mailer.TRANSPORT_ENTRA_API,
            variable=transport_var,
        ).pack(side="left", padx=(8, 0))

        smtp_frame = ttk.LabelFrame(win, text="Ustawienia SMTP")
        smtp_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(6, 2))
        smtp_frame.columnconfigure(1, weight=1)

        ttk.Label(smtp_frame, text="Serwer SMTP:").grid(
            row=0, column=0, sticky="w", padx=8, pady=2
        )
        host_entry = ttk.Entry(smtp_frame, textvariable=host_var)
        host_entry.grid(
            row=0, column=1, sticky="ew", padx=8, pady=2
        )

        ttk.Label(smtp_frame, text="Port SMTP:").grid(
            row=1, column=0, sticky="w", padx=8, pady=2
        )
        port_entry = ttk.Entry(smtp_frame, textvariable=port_var, width=10)
        port_entry.grid(
            row=1, column=1, sticky="w", padx=8, pady=2
        )

        ttk.Label(smtp_frame, text="Zabezpieczenie:").grid(
            row=2, column=0, sticky="w", padx=8, pady=2
        )
        security_box = ttk.Combobox(
            smtp_frame,
            textvariable=security_var,
            values=(mailer.SECURITY_STARTTLS, mailer.SECURITY_SSL, mailer.SECURITY_NONE),
            state="readonly",
        )
        security_box.grid(row=2, column=1, sticky="w", padx=8, pady=2)

        ttk.Label(smtp_frame, text="Login SMTP:").grid(
            row=3, column=0, sticky="w", padx=8, pady=2
        )
        username_entry = ttk.Entry(smtp_frame, textvariable=username_var)
        username_entry.grid(
            row=3, column=1, sticky="ew", padx=8, pady=2
        )

        ttk.Label(smtp_frame, text="Hasło SMTP:").grid(
            row=4, column=0, sticky="w", padx=8, pady=2
        )
        password_entry = ttk.Entry(smtp_frame, textvariable=password_var, show="*")
        password_entry.grid(
            row=4, column=1, sticky="ew", padx=8, pady=2
        )

        ttk.Label(smtp_frame, text="Adres nadawcy:").grid(
            row=5, column=0, sticky="w", padx=8, pady=2
        )
        sender_entry = ttk.Entry(smtp_frame, textvariable=sender_var)
        sender_entry.grid(
            row=5, column=1, sticky="ew", padx=8, pady=2
        )

        api_frame = ttk.LabelFrame(win, text="Microsoft Entra API")
        api_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=(6, 2))
        api_frame.columnconfigure(1, weight=1)

        ttk.Label(
            api_frame,
            text=(
                "Uwierzytelnianie: użyj gotowego tokenu Bearer ALBO danych aplikacji.\n"
                "Mapowanie: Aplikacja(klient)=Client ID, Dzierżawa=Tenant ID, "
                "Wartość klucza=Secret Value.\n"
                "Identyfikator obiektu i Identyfikator wpisu tajnego nie są wymagane."
            ),
            justify="left",
            wraplength=620,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 6))

        ttk.Label(api_frame, text="Token Bearer:").grid(
            row=1, column=0, sticky="w", padx=8, pady=2
        )
        token_entry = ttk.Entry(api_frame, textvariable=entra_token_var, show="*")
        token_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=2)

        ttk.Label(api_frame, text="Tenant ID (dzierżawy):").grid(
            row=2, column=0, sticky="w", padx=8, pady=2
        )
        entra_tenant_entry = ttk.Entry(api_frame, textvariable=entra_tenant_var)
        entra_tenant_entry.grid(
            row=2, column=1, sticky="ew", padx=8, pady=2
        )

        ttk.Label(api_frame, text="Client ID (aplikacji):").grid(
            row=3, column=0, sticky="w", padx=8, pady=2
        )
        entra_client_entry = ttk.Entry(api_frame, textvariable=entra_client_var)
        entra_client_entry.grid(
            row=3, column=1, sticky="ew", padx=8, pady=2
        )

        ttk.Label(api_frame, text="Secret Value (wartość klucza):").grid(
            row=4, column=0, sticky="w", padx=8, pady=2
        )
        entra_client_secret_entry = ttk.Entry(
            api_frame, textvariable=entra_client_secret_var, show="*"
        )
        entra_client_secret_entry.grid(
            row=4, column=1, sticky="ew", padx=8, pady=2
        )

        ttk.Label(api_frame, text="Nadawca (UPN/ID):").grid(
            row=5, column=0, sticky="w", padx=8, pady=2
        )
        entra_sender_entry = ttk.Entry(api_frame, textvariable=entra_sender_var)
        entra_sender_entry.grid(
            row=5, column=1, sticky="ew", padx=8, pady=2
        )

        ttk.Label(api_frame, text="Endpoint (opcjonalnie):").grid(
            row=6, column=0, sticky="w", padx=8, pady=2
        )
        entra_endpoint_entry = ttk.Entry(api_frame, textvariable=entra_endpoint_var)
        entra_endpoint_entry.grid(
            row=6, column=1, sticky="ew", padx=8, pady=2
        )

        ttk.Label(api_frame, text="Certyfikat (ID):").grid(
            row=7, column=0, sticky="w", padx=8, pady=2
        )
        secret_key_id_entry = ttk.Entry(
            api_frame, textvariable=secret_key_id_var, state="readonly"
        )
        secret_key_id_entry.grid(row=7, column=1, sticky="ew", padx=8, pady=2)

        cert_btns = ttk.Frame(api_frame)
        cert_btns.grid(row=8, column=1, sticky="w", padx=8, pady=(2, 2))
        generate_cert_btn = ttk.Button(
            cert_btns, text="Generuj certyfikat", command=lambda: None
        )
        generate_cert_btn.pack(side="left")

        fetch_token_btn = ttk.Button(
            api_frame, text="Pobierz token", command=lambda: None
        )
        fetch_token_btn.grid(row=9, column=1, sticky="w", padx=8, pady=(4, 2))

        ttk.Checkbutton(
            api_frame,
            text="Pokaż pola wrażliwe (hasła/tokeny)",
            variable=show_sensitive_var,
        ).grid(row=10, column=1, sticky="w", padx=8, pady=(0, 2))

        ttk.Label(win, text="Temat (prefix):").grid(
            row=4, column=0, sticky="w", padx=10, pady=2
        )
        ttk.Entry(win, textvariable=subject_var).grid(
            row=4, column=1, sticky="ew", padx=10, pady=2
        )

        ttk.Label(win, text="Timeout [s]:").grid(
            row=5, column=0, sticky="w", padx=10, pady=2
        )
        ttk.Entry(win, textvariable=timeout_var, width=10).grid(
            row=5, column=1, sticky="w", padx=10, pady=2
        )

        ttk.Label(
            win,
            text="Odbiorcy (jeden adres w linii, albo rozdzielone przecinkiem):",
        ).grid(row=6, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 2))
        recipients_box = tk.Text(win, height=6, width=46)
        recipients_box.grid(row=7, column=0, columnspan=2, sticky="ew", padx=10, pady=2)
        recipients_box.insert("1.0", mailer.recipients_to_text(cfg["recipients"]))

        btns = ttk.Frame(win)
        btns.grid(row=8, column=0, columnspan=2, sticky="ew", padx=10, pady=(8, 10))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)
        btns.columnconfigure(2, weight=1)
        btns.columnconfigure(3, weight=1)

        smtp_controls = [
            host_entry,
            port_entry,
            security_box,
            username_entry,
            password_entry,
            sender_entry,
        ]
        api_controls = [
            token_entry,
            entra_tenant_entry,
            entra_client_entry,
            entra_client_secret_entry,
            entra_sender_entry,
            entra_endpoint_entry,
            secret_key_id_entry,
            generate_cert_btn,
            fetch_token_btn,
        ]

        def apply_sensitive_visibility(*_args):
            mask = "" if show_sensitive_var.get() else "*"
            password_entry.configure(show=mask)
            token_entry.configure(show=mask)
            entra_client_secret_entry.configure(show=mask)

        def set_controls_state(controls, enabled_state):
            for control in controls:
                if isinstance(control, tk.Text):
                    control.configure(state="normal" if enabled_state else "disabled")
                    continue
                try:
                    if enabled_state:
                        control.state(["!disabled"])
                    else:
                        control.state(["disabled"])
                except Exception:
                    try:
                        control.configure(
                            state=("normal" if enabled_state else "disabled")
                        )
                    except Exception:
                        pass

        def apply_transport_state(*_args):
            use_smtp = transport_var.get() == mailer.TRANSPORT_SMTP
            set_controls_state(smtp_controls, use_smtp)
            set_controls_state(api_controls, not use_smtp)

        transport_var.trace_add("write", apply_transport_state)
        apply_transport_state()
        show_sensitive_var.trace_add("write", apply_sensitive_visibility)
        apply_sensitive_visibility()

        def collect(strict=False, require_recipients=False):
            recipients_value = recipients_box.get("1.0", "end").strip()
            token_value = entra_token_var.get().strip()
            return self._mail_settings_from_inputs(
                transport=transport_var.get(),
                enabled=enabled_var.get(),
                smtp_host=host_var.get(),
                smtp_port=port_var.get(),
                smtp_security=security_var.get(),
                username=username_var.get(),
                password=password_var.get(),
                sender=sender_var.get(),
                entra_token=token_value,
                entra_tenant_id=entra_tenant_var.get(),
                entra_client_id=entra_client_var.get(),
                entra_client_secret=entra_client_secret_var.get(),
                entra_sender=entra_sender_var.get(),
                entra_endpoint=entra_endpoint_var.get(),
                secret_key_id=secret_key_id_var.get(),
                recipients_text=recipients_value,
                subject_prefix=subject_var.get(),
                timeout_seconds=timeout_var.get(),
                strict=strict,
                require_recipients=require_recipients,
            )

        def collect_entra(strict=False):
            recipients_value = recipients_box.get("1.0", "end").strip()
            token_value = entra_token_var.get().strip()
            return self._mail_settings_from_inputs(
                transport=mailer.TRANSPORT_ENTRA_API,
                enabled=enabled_var.get(),
                smtp_host=host_var.get(),
                smtp_port=port_var.get(),
                smtp_security=security_var.get(),
                username=username_var.get(),
                password=password_var.get(),
                sender=sender_var.get(),
                entra_token=token_value,
                entra_tenant_id=entra_tenant_var.get(),
                entra_client_id=entra_client_var.get(),
                entra_client_secret=entra_client_secret_var.get(),
                entra_sender=entra_sender_var.get(),
                entra_endpoint=entra_endpoint_var.get(),
                secret_key_id=secret_key_id_var.get(),
                recipients_text=recipients_value,
                subject_prefix=subject_var.get(),
                timeout_seconds=timeout_var.get(),
                strict=strict,
                require_recipients=False,
            )

        def fetch_token():
            cfg_for_token = collect_entra(strict=False)
            if not cfg_for_token:
                return
            self.set_status("Pobieranie tokenu Entra...")

            def worker():
                try:
                    token = mailer.request_entra_token(cfg_for_token)
                except Exception as exc:
                    logger.exception("Failed to fetch Entra token")
                    self.ui_call(
                        messagebox.showerror,
                        "E-mail",
                        f"Nie udało się pobrać tokenu Entra: {exc}",
                    )
                    self.ui_call(self.set_status, "Błąd pobierania tokenu")
                    return

                def on_success():
                    entra_token_var.set(token)
                    apply_transport_state()
                    self.set_status("Pobrano token Entra")
                    messagebox.showinfo(
                        "E-mail",
                        "Pobrano token Entra API i wstawiono do pola Token Bearer.",
                    )

                self.ui_call(on_success)

            threading.Thread(target=worker, daemon=True).start()

        def generate_certificate():
            cfg_for_cert = collect_entra(strict=True)
            if not cfg_for_cert:
                return
            # Always derive certificate ID from current Tenant+Client inputs.
            cfg_for_cert["secret_key_id"] = ""
            self.set_status("Testowanie Entra przed generowaniem certyfikatu...")

            def worker():
                try:
                    mailer.test_connection(cfg_for_cert)
                except Exception as exc:
                    logger.exception(
                        "Connection test failed before certificate generation"
                    )
                    self.ui_call(
                        messagebox.showerror,
                        "E-mail",
                        "Nie można wygenerować certyfikatu, ponieważ test połączenia "
                        f"Entra zakończył się błędem:\n{exc}",
                    )
                    self.ui_call(self.set_status, "Błąd testu Entra")
                    return

                try:
                    key_id, cert_path, created = mailer.generate_secret_certificate(
                        cfg_for_cert
                    )
                except Exception as exc:
                    logger.exception("Failed to generate mail secret certificate")
                    self.ui_call(
                        messagebox.showerror,
                        "E-mail",
                        f"Nie udało się wygenerować certyfikatu: {exc}",
                    )
                    self.ui_call(self.set_status, "Błąd generowania certyfikatu")
                    return

                def on_success():
                    secret_key_id_var.set(key_id)
                    self.set_status("Certyfikat szyfrowania gotowy")
                    if created:
                        messagebox.showinfo(
                            "E-mail",
                            "Wygenerowano certyfikat szyfrowania.\n"
                            f"Plik: {cert_path}",
                        )
                    else:
                        messagebox.showinfo(
                            "E-mail",
                            "Certyfikat już istnieje i zostanie użyty.\n"
                            f"Plik: {cert_path}",
                        )

                self.ui_call(on_success)

            threading.Thread(target=worker, daemon=True).start()

        fetch_token_btn.configure(command=fetch_token)
        generate_cert_btn.configure(command=generate_certificate)

        def save_only():
            new_cfg = collect(strict=False, require_recipients=False)
            if not new_cfg:
                return
            self.mail_config = new_cfg
            messagebox.showinfo(
                "E-mail",
                "Zapisano ustawienia e-mail. "
                "Użyj 'Zapisz konfigurację', aby zapisać je do pliku config.json.",
            )

        def test_connection():
            test_cfg = collect(strict=True, require_recipients=False)
            if not test_cfg:
                return
            self._run_mail_action_async(
                test_cfg,
                mailer.test_connection,
                "Połączenie działa poprawnie.",
                "Testowanie połączenia...",
                "Połączenie OK",
                "Test połączenia nie powiódł się",
            )

        def send_test_message():
            test_cfg = collect(strict=True, require_recipients=True)
            if not test_cfg:
                return
            self._run_mail_action_async(
                test_cfg,
                mailer.send_test_email,
                "Wysłano testową wiadomość.",
                "Wysyłanie testowej wiadomości...",
                "Wysłano testową wiadomość",
                "Wysyłka testowej wiadomości nie powiodła się",
            )

        def close():
            self.mail_settings_win = None
            win.destroy()

        ttk.Button(btns, text="Test połączenia", command=test_connection).grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Button(btns, text="Wyślij testową wiadomość", command=send_test_message).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )
        ttk.Button(btns, text="Zapisz", command=save_only).grid(
            row=0, column=2, sticky="ew", padx=(6, 0)
        )
        ttk.Button(btns, text="Zamknij", command=close).grid(
            row=0, column=3, sticky="ew", padx=(6, 0)
        )

        win.protocol("WM_DELETE_WINDOW", close)

    def on_generation_complete(self, report):
        self.last_generation_report = deepcopy(report)
        if report.get("status") == "no_changes" and not report.get("warnings"):
            logger.info("Skipping report email: no PDF changes detected.")
            return
        cfg = mailer.normalize_mail_config(getattr(self, "mail_config", {}))
        if not cfg.get("enabled"):
            return
        try:
            cfg = mailer.validate_mail_config(cfg, require_recipients=True)
        except ValueError as exc:
            logger.warning("Mail config is invalid, skipping report: %s", exc)
            messagebox.showwarning(
                "E-mail",
                f"Nie wysłano raportu e-mail: {exc}",
            )
            return

        self.set_status("Wysyłanie raportu e-mail...")

        def worker():
            try:
                mailer.send_generation_report(cfg, report)
            except Exception as exc:
                logger.exception("Failed to send generation report email")
                self.ui_call(
                    messagebox.showwarning,
                    "E-mail",
                    f"Nie udało się wysłać raportu: {exc}",
                )
                self.ui_call(self.set_status, "Błąd wysyłki e-mail")
                return
            self.ui_call(self.set_status, "Wysłano raport e-mail")

        threading.Thread(target=worker, daemon=True).start()

    def toggle_static(self, name, state):
        if state:
            value = self.static_entries[name].get()
            if name not in self.elements:
                element = DraggableElement(self, self.canvas, name, value)
                element.is_image = name in self.image_fields
                element.update_value(value)
                self.elements[name] = element
                self.restack_elements()
            else:
                self.elements[name].is_image = name in self.image_fields
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

    def _set_indicator(self, name, active):
        indicator = self.static_indicators.get(name) or self.column_indicators.get(name)
        if not indicator:
            return
        color = self.highlight_color if active else self.panel_bg
        indicator.configure(bg=color)

    def update_field_highlights(self, names=None):
        if names is None:
            names = [el.name for el in self.selected_elements]
        new_set = set(names)
        for name in self._highlighted_fields - new_set:
            self._set_indicator(name, False)
        for name in new_set - self._highlighted_fields:
            self._set_indicator(name, True)
        self._highlighted_fields = new_set

    def get_formula_tooltip(self, field_name):
        formulas = self.formula_map.get(field_name) or []
        if not formulas:
            return ""
        lines = [f"{field_name}", "Przykładowe formuły:"]
        for formula in formulas[:3]:
            cleaned = formula.replace("\n", " ").strip()
            if len(cleaned) > 160:
                cleaned = cleaned[:157] + "..."
            lines.append(f"- {cleaned}")
        return "\n".join(lines)

    def bind_formula_tooltip(self, widget, field_name):
        if not hasattr(self, "tooltip"):
            return
        self.tooltip.bind(widget, text_func=lambda n=field_name: self.get_formula_tooltip(n))

    def create_static_row(self, name, value=None):
        row = ttk.Frame(self.static_frame)
        if hasattr(self, "add_static_btn"):
            row.pack(fill="x", pady=2, before=self.add_static_btn)
        else:
            row.pack(fill="x", pady=2)
        indicator = tk.Frame(row, width=6, height=18, bg=self.panel_bg)
        indicator.pack(side="left", fill="y", padx=(0, 4))
        indicator.pack_propagate(False)
        var = tk.BooleanVar()
        chk = ttk.Checkbutton(row, variable=var, command=lambda n=name, v=var: self.toggle_static(n, v.get()))
        chk.pack(side="left")
        ttk.Label(row, text=name).pack(side="left")
        entry_var = tk.StringVar(value=value if value is not None else "")
        entry = ttk.Entry(row, textvariable=entry_var, width=15)
        entry.pack(side="left", padx=5)
        entry_var.trace_add("write", lambda *a, f=name: self.update_static_value(f))
        img_var = tk.BooleanVar(value=name in self.image_fields)
        img_chk = ttk.Checkbutton(
            row,
            text="IMG",
            variable=img_var,
            command=lambda n=name, v=img_var: self.set_field_image(n, v.get()),
        )
        img_chk.pack(side="left", padx=2)
        del_btn = ttk.Button(row, text="X", width=2, command=lambda n=name, r=row: self.remove_static_field(n, r))
        del_btn.pack(side="left")
        self.static_vars[name] = var
        self.static_entries[name] = entry_var
        self.static_rows[name] = row
        self.static_indicators[name] = indicator
        self.image_vars[name] = img_var
        if hasattr(self, "tooltip"):
            self.tooltip.bind(chk, text="Włącz pole statyczne")
            self.tooltip.bind(img_chk, text="Traktuj wartość jako obraz")
            self.tooltip.bind(del_btn, text="Usuń pole statyczne")

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
        self.static_indicators.pop(name, None)
        self.image_vars.pop(name, None)
        self.update_field_highlights()
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
        self.update_field_highlights()

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
            "image_fields": sorted(self.image_fields),
            "image_dirs": list(self.image_dirs),
        }
        self.history.append(state)
        if len(self.history) > 50:
            self.history.pop(0)
        self.future.clear()

    def restore_state(self, state):
        self.image_fields = set(state.get("image_fields", []))
        self.image_dirs = list(state.get("image_dirs", []))
        self.refresh_image_dir_list()
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
            if conf.get("is_image"):
                self.image_fields.add(name)
            el.is_image = name in self.image_fields
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
        self.apply_image_field_state()

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
        if self.conditions_win and self.conditions_win.winfo_exists():
            self.conditions_win.lift()
            self.conditions_win.focus_force()
            return

        win = tk.Toplevel(self)
        self.conditions_win = win
        win.title("Warunki")

        # allow the window to expand with resizing
        win.columnconfigure(0, weight=1)
        win.columnconfigure(1, weight=1)
        win.rowconfigure(3, weight=1)

        src_var = tk.StringVar()
        tgt_var = tk.StringVar()
        options = list(self.elements.keys())

        ttk.Label(win, text="Jeśli puste:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(win, values=options, textvariable=src_var).grid(
            row=0, column=1, sticky="ew", padx=5, pady=2
        )
        ttk.Label(win, text="ukryj:").grid(row=1, column=0, sticky="w")
        ttk.Combobox(win, values=options, textvariable=tgt_var).grid(
            row=1, column=1, sticky="ew", padx=5, pady=2
        )

        table = ttk.Frame(win)
        table.grid(row=3, column=0, columnspan=2, pady=5, sticky="nsew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)

        tree = ttk.Treeview(
            table,
            columns=("src", "tgt"),
            show="headings",
            selectmode="extended",
        )
        tree.heading("src", text="Jeśli puste")
        tree.heading("tgt", text="Ukryj")
        tree.column("src", width=240, anchor="w", stretch=True)
        tree.column("tgt", width=240, anchor="w", stretch=True)
        vsb = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        for s, t in self.conditions:
            tree.insert("", "end", values=(s, t))
        def add():
            s = src_var.get()
            t = tgt_var.get()
            if s and t and (s, t) not in self.conditions:
                self.conditions.append((s, t))
                tree.insert("", "end", values=(s, t))
        ttk.Button(win, text="Dodaj", command=add).grid(
            row=2, column=0, columnspan=2, pady=5, sticky="ew"
        )
        def remove():
            sel = tree.selection()
            if not sel:
                return
            for item in sel:
                values = tree.item(item, "values")
                if values:
                    try:
                        self.conditions.remove(tuple(values))
                    except ValueError:
                        pass
                tree.delete(item)
        ttk.Button(win, text="Usuń zaznaczone", command=remove).grid(
            row=4, column=0, columnspan=2, pady=5, sticky="ew"
        )

        def close():
            self.conditions_win = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", close)

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
    def _set_preview_status(self, text):
        if hasattr(self, "preview_status_var") and self.preview_status_var is not None:
            self.preview_status_var.set(text)

    def _preview_animation_tick(self):
        if not self.preview_in_progress:
            return
        dots = "." * (self.preview_animation_step % 4)
        self._set_preview_status(f"Ładowanie{dots}")
        self.preview_animation_step += 1
        self.preview_animation_after = self.after(220, self._preview_animation_tick)

    def _start_preview_animation(self):
        if self.preview_in_progress:
            return
        self.preview_in_progress = True
        self.preview_animation_step = 0
        if hasattr(self, "preview_btn") and self.preview_btn:
            self.preview_btn.state(["disabled"])
        self._preview_animation_tick()

    def _clear_preview_status(self):
        if not self.preview_in_progress:
            self._set_preview_status("")

    def _finish_preview_animation(self, status_text=""):
        self.preview_in_progress = False
        if self.preview_animation_after is not None:
            try:
                self.after_cancel(self.preview_animation_after)
            except Exception:
                pass
            self.preview_animation_after = None
        if hasattr(self, "preview_btn") and self.preview_btn:
            self.preview_btn.state(["!disabled"])
        self._set_preview_status(status_text)
        if status_text:
            self.after(1500, self._clear_preview_status)

    # ------------------------------------------------------------------
    def preview_row(self):
        if not self.dataframes:
            return
        if self.preview_in_progress:
            return
        try:
            idx = int(self.row_var.get()) - 1
        except ValueError:
            messagebox.showerror("Błąd", "Nieprawidłowy numer wiersza")
            return
        self._start_preview_animation()

        def worker():
            try:
                values = {}
                for name in self.elements.keys():
                    if ":" in name:
                        sheet, col = name.split(":", 1)
                        df = self.dataframes.get(sheet)
                        value = ""
                        if df is not None and 0 <= idx < len(df):
                            value = df.iloc[idx].get(col)
                            value = round_numeric_value(value)
                    else:
                        if name in getattr(self, "static_entries", {}):
                            value = self.static_entries[name].get()
                        else:
                            value = name
                    try:
                        if pd.isna(value):
                            value = ""
                    except TypeError:
                        if value is None:
                            value = ""
                    values[name] = value

                hidden = set()
                for src, tgt in self.conditions:
                    if src not in values:
                        continue
                    src_val = values.get(src, "")
                    try:
                        empty = pd.isna(src_val) or src_val == ""
                    except TypeError:
                        empty = src_val == ""
                    if empty:
                        hidden.add(tgt)
                ordered = sorted(self.elements.items(), key=lambda kv: kv[1].layer)
            except Exception as exc:
                logger.exception("Failed to prepare preview row")

                def on_error():
                    self._finish_preview_animation("")
                    messagebox.showerror("Błąd", f"Nie udało się przygotować podglądu: {exc}")

                self.ui_call(on_error)
                return

            def apply_batch(start=0):
                if not self.preview_in_progress:
                    return
                batch_size = 16
                end = min(start + batch_size, len(ordered))
                for name, element in ordered[start:end]:
                    value = values.get(name, "")
                    if name in hidden:
                        value = ""
                    try:
                        element.update_value(value)
                    except Exception:
                        logger.exception("Failed to update preview element %s", name)
                if end < len(ordered):
                    self.after(1, lambda: apply_batch(end))
                else:
                    self._finish_preview_animation("Podgląd gotowy")

            self.ui_call(apply_batch)

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    def save_config(self):
        save_config_func(self)

    def load_config(self, startup=False, path=None):
        load_config_func(self, startup=startup, path=path)
    # ------------------------------------------------------------------
    def generate_pds(self):
        if self.generation_in_progress:
            return
        self._start_generation_ui()
        try:
            started = export_pds(self)
        except Exception:
            self.finish_generation_ui(status="Błąd")
            raise
        if not started and self.generation_in_progress:
            self.finish_generation_ui(status="Gotowe")

    def cancel_generation(self):
        if not self.generation_in_progress:
            return
        if self.cancel_event:
            self.cancel_event.set()
        self.set_status("Anulowanie...")
        if hasattr(self, "cancel_btn") and self.cancel_btn:
            self.cancel_btn.state(["disabled"])

    def ui_call(self, func, *args, **kwargs):
        self.after(0, lambda: func(*args, **kwargs))

    def set_status(self, text):
        if hasattr(self, "status_var") and self.status_var:
            self.status_var.set(f"Akcja: {text}")

    def set_counts(self, total_rows=None, total_tasks=None, processed=None, skipped=None):
        if total_rows is not None and hasattr(self, "rows_var") and self.rows_var:
            self.rows_var.set(f"Wykryto wierszy: {total_rows}")
        if total_tasks is not None and hasattr(self, "progress_info_var") and self.progress_info_var:
            if processed is None:
                processed = 0
            self.progress_info_var.set(
                f"Przetworzono plików: {processed}/{total_tasks}"
            )
        if skipped is not None and hasattr(self, "skip_var") and self.skip_var:
            self.skip_var.set(f"Pomijam: {skipped}")

    def set_progress_indeterminate(self, active=True):
        if not hasattr(self, "progress") or not self.progress:
            return
        if active:
            self.progress.config(mode="indeterminate")
            self.progress.start(10)
        else:
            self.progress.stop()
            self.progress.config(mode="determinate")

    def finish_generation_ui(self, status="Zakończono"):
        self.generation_in_progress = False
        self.cancel_event = None
        if hasattr(self, "cancel_btn") and self.cancel_btn:
            self.cancel_btn.state(["!disabled"])
        self._toggle_generation_buttons(False)
        self.set_progress_indeterminate(False)
        if hasattr(self, "progress") and self.progress:
            self.progress.config(value=0)
        if hasattr(self, "time_label") and self.time_label:
            self.time_label.config(text=status)
        self.set_status(status)

    def _start_generation_ui(self):
        self.generation_in_progress = True
        self.cancel_event = threading.Event()
        self._toggle_generation_buttons(True)
        if hasattr(self, "cancel_btn") and self.cancel_btn:
            self.cancel_btn.state(["!disabled"])
        self.set_status("Przygotowanie danych...")
        self.set_progress_indeterminate(True)
        if hasattr(self, "rows_var") and self.rows_var:
            self.rows_var.set("")
        if hasattr(self, "skip_var") and self.skip_var:
            self.skip_var.set("")
        if hasattr(self, "progress_info_var") and self.progress_info_var:
            self.progress_info_var.set("")
        if hasattr(self, "time_label") and self.time_label:
            self.time_label.config(text="")
        if hasattr(self, "progress") and self.progress:
            self.progress.config(value=0)
        self.update_idletasks()

    def _toggle_generation_buttons(self, running):
        if not hasattr(self, "generate_btn") or not hasattr(self, "cancel_btn"):
            return
        if running:
            self.generate_btn.pack_forget()
            if not self.cancel_btn.winfo_ismapped():
                self.cancel_btn.pack(fill="x")
        else:
            self.cancel_btn.pack_forget()
            if not self.generate_btn.winfo_ismapped():
                self.generate_btn.pack(fill="x")

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
            rel_max_f = getattr(el, "max_font_size", el.font_size) / self.scale
            el.x = rel_x * new_scale
            el.y = rel_y * new_scale
            el.width = rel_w * new_scale
            el.height = rel_h * new_scale
            el.font_size = rel_f * new_scale
            el.max_font_size = rel_max_f * new_scale
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
            rel_max_f = getattr(el, "max_font_size", el.font_size) / self.scale
            el.x = rel_x * new_scale
            el.y = rel_y * new_scale
            el.width = rel_w * new_scale
            el.height = rel_h * new_scale
            el.font_size = rel_f * new_scale
            el.max_font_size = rel_max_f * new_scale
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
            if hasattr(self, "auto_font_var"):
                self.auto_font_var.set(bool(getattr(self.selected_element, "auto_font", True)))
            if hasattr(self, "auto_font_check"):
                self.auto_font_check.state(["!disabled"])
            self.bg_check.state(["!disabled"])
            self.transparent_var.set(not self.selected_element.bg_visible)
            self.layer_entry.configure(state="normal")
            self.layer_var.set(str(int(self.selected_element.layer)))
        else:
            self.font_entry.configure(state="disabled")
            self.font_size_var.set("")
            if hasattr(self, "auto_font_var"):
                self.auto_font_var.set(False)
            if hasattr(self, "auto_font_check"):
                self.auto_font_check.state(["disabled"])
            self.transparent_var.set(False)
            self.bg_check.state(["disabled"])
            self.layer_entry.configure(state="disabled")
            self.layer_var.set("")
        self.update_field_highlights()

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
        el.max_font_size = el.font_size
        el.auto_font = False
        el.apply_font()
        self.font_size_var.set(str(int(el.font_size / self.scale)))
        if hasattr(self, "auto_font_var"):
            self.auto_font_var.set(False)
        self.push_history()

    def decrease_font(self):
        el = self.selected_element
        if not el:
            return
        if el.font_size > self.scale:
            el.font_size -= self.scale
            el.max_font_size = el.font_size
            el.auto_font = False
            el.apply_font()
            self.font_size_var.set(str(int(el.font_size / self.scale)))
            if hasattr(self, "auto_font_var"):
                self.auto_font_var.set(False)
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
        el.max_font_size = el.font_size
        el.auto_font = False
        el.apply_font()
        if hasattr(self, "auto_font_var"):
            self.auto_font_var.set(False)
        self.push_history()

    def toggle_auto_font(self):
        if not self.selected_elements:
            return
        state = bool(self.auto_font_var.get())
        for el in self.selected_elements:
            el.auto_font = state
            if state and not hasattr(el, "max_font_size"):
                el.max_font_size = el.font_size
            el.sync_canvas()
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


    def acquire_excel_lock(self, path):
        current_lock_path = getattr(self, "excel_lock_path", None)
        expected_lock_path = locks._lock_path(path)

        if current_lock_path == expected_lock_path and os.path.exists(current_lock_path):
            return True

        lock_path = locks.acquire_lock(path, os.path.basename(path))
        if not lock_path:
            return False
        self.release_lock("excel_lock_path")
        self.excel_lock_path = lock_path
        return True

    def release_lock(self, attr):
        path = getattr(self, attr, None)
        if path:
            locks.release_lock(path)
        setattr(self, attr, None)

    def on_close(self):
        self.preview_in_progress = False
        if self.preview_animation_after is not None:
            try:
                self.after_cancel(self.preview_animation_after)
            except Exception:
                pass
            self.preview_animation_after = None
        self.release_lock("excel_lock_path")
        self.release_lock("config_lock_path")
        self.destroy()
