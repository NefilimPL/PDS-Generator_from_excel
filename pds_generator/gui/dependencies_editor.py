from types import SimpleNamespace

import tkinter as tk
from tkinter import messagebox, ttk

from ..layout_dependencies import (
    DEPENDENCY_DIRECTION_BELOW,
    DEPENDENCY_DIRECTION_LABELS,
    DEPENDENCY_STATUS_ALREADY_ALIGNED,
    DEPENDENCY_STATUS_APPLIED,
    DEPENDENCY_STATUS_BLOCKED,
    apply_layout_dependencies,
    normalize_dependency,
)


class DependenciesEditor(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Edytor zależności")
        self.geometry("1380x860")
        self.minsize(1100, 700)

        self.direction_labels = dict(DEPENDENCY_DIRECTION_LABELS)
        self.labels_to_direction = {
            label: direction for direction, label in self.direction_labels.items()
        }

        self.current_anchor_names = []
        self.current_mover_names = []
        self.capture_target = None
        self.active_dependency_index = None
        self.element_items = {}
        self.item_to_name = {}
        self.preview_statuses = []
        self.preview_focus_index = None
        self.sel_rect = None
        self.sel_start = None
        self.sel_additive = False
        self._suspend_tree_event = False
        self._suspend_preview_traces = False
        self._refresh_after_id = None
        self._refresh_in_progress = False
        self._refresh_pending = False
        self.simulation_hidden_names = set()
        self.simulation_row_number = None
        self.simulation_error = ""

        self.mode_var = tk.StringVar(value="Tryb wyboru: brak")
        self.anchor_summary_var = tk.StringVar(value="Kotwice: brak")
        self.mover_summary_var = tk.StringVar(value="Przesuwane: brak")
        self.status_var = tk.StringVar(
            value="Podgląd statusu liczy kolizje przy obecnym układzie elementów."
        )
        self.direction_var = tk.StringVar(
            value=self.direction_labels[DEPENDENCY_DIRECTION_BELOW]
        )
        self.gap_var = tk.StringVar(value="0")
        current_row = "1"
        if hasattr(parent, "row_var"):
            try:
                current_row = str(parent.row_var.get() or "1")
            except Exception:
                current_row = "1"
        self.preview_row_var = tk.StringVar(value=current_row)
        self.simulate_row_var = tk.BooleanVar(value=bool(getattr(parent, "dataframes", {})))
        self.show_default_positions_var = tk.BooleanVar(value=False)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self._build_ui()
        self.direction_var.trace_add("write", lambda *_args: self._request_preview_refresh())
        self.gap_var.trace_add("write", lambda *_args: self._request_preview_refresh())
        self.preview_row_var.trace_add("write", lambda *_args: self._request_preview_refresh(150))
        self.simulate_row_var.trace_add("write", lambda *_args: self._request_preview_refresh())
        self.show_default_positions_var.trace_add(
            "write", lambda *_args: self._request_preview_refresh()
        )
        self.refresh_from_parent()
        self.protocol("WM_DELETE_WINDOW", self.close)

    def _request_preview_refresh(self, delay_ms=0):
        if self._suspend_preview_traces:
            return
        if self._refresh_in_progress:
            self._refresh_pending = True
            return
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except tk.TclError:
                pass
            self._refresh_after_id = None
        self._refresh_after_id = self.after(
            delay_ms,
            self._run_requested_preview_refresh,
        )

    def _run_requested_preview_refresh(self):
        self._refresh_after_id = None
        if self._refresh_in_progress:
            self._refresh_pending = True
            return
        self._refresh_preview()

    def _set_preview_values(self, direction_label=None, gap_value=None):
        self._suspend_preview_traces = True
        try:
            if direction_label is not None and self.direction_var.get() != direction_label:
                self.direction_var.set(direction_label)
            if gap_value is not None:
                gap_text = str(gap_value)
                if self.gap_var.get() != gap_text:
                    self.gap_var.set(gap_text)
        finally:
            self._suspend_preview_traces = False

    def _use_current_preview_row(self):
        if not hasattr(self.parent, "row_var"):
            return
        try:
            current_row = str(self.parent.row_var.get() or "1")
        except Exception:
            current_row = "1"
        if self.preview_row_var.get() != current_row:
            self.preview_row_var.set(current_row)

    def _build_ui(self):
        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        header.columnconfigure(0, weight=1)

        ttk.Label(
            header,
            text=(
                "1. Kliknij przycisk wyboru zakresu. 2. Klikaj bloki albo zaznaczaj je "
                "ramką na planszy. 3. Ustaw kierunek i odstęp. 4. Dodaj lub zaktualizuj zależność. "
                "Aby edytować istniejącą, kliknij ją na liście, zmień ustawienia i użyj "
                "\"Zapisz zmiany\"."
            ),
            justify="left",
            style="DialogTitle.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            textvariable=self.mode_var,
            justify="right",
        ).grid(row=0, column=1, sticky="e", padx=(12, 0))

        workspace = ttk.Panedwindow(self, orient="horizontal")
        workspace.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        canvas_shell = ttk.Frame(workspace)
        canvas_shell.columnconfigure(0, weight=1)
        canvas_shell.rowconfigure(0, weight=1)
        workspace.add(canvas_shell, weight=4)

        self.canvas = tk.Canvas(
            canvas_shell,
            bg="#d4dae4",
            highlightthickness=0,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<ButtonPress-1>", self.canvas_button_press)
        self.canvas.bind("<B1-Motion>", self.canvas_drag_select)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_button_release)

        vscroll = ttk.Scrollbar(canvas_shell, orient="vertical", command=self.canvas.yview)
        hscroll = ttk.Scrollbar(canvas_shell, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll.grid(row=1, column=0, sticky="ew")

        right = ttk.Frame(workspace, padding=(8, 0, 0, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)
        workspace.add(right, weight=2)

        current_box = ttk.LabelFrame(right, text="Bieżąca Zależność")
        current_box.grid(row=0, column=0, sticky="ew")
        current_box.columnconfigure(0, weight=1)

        ttk.Label(current_box, textvariable=self.mover_summary_var, justify="left").grid(
            row=0, column=0, sticky="ew", padx=8, pady=(8, 4)
        )
        ttk.Label(current_box, textvariable=self.anchor_summary_var, justify="left").grid(
            row=1, column=0, sticky="ew", padx=8, pady=(0, 8)
        )

        button_row = ttk.Frame(current_box)
        button_row.grid(row=2, column=0, sticky="ew", padx=8)
        button_row.columnconfigure(0, weight=1)
        button_row.columnconfigure(1, weight=1)
        button_row.columnconfigure(2, weight=1)
        button_row.columnconfigure(3, weight=1)
        ttk.Button(
            button_row,
            text="Zaznacz przesuwane",
            command=lambda: self._set_capture_target("mover"),
        ).grid(row=0, column=0, sticky="ew")
        ttk.Button(
            button_row,
            text="Zaznacz kotwice",
            command=lambda: self._set_capture_target("anchor"),
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(
            button_row,
            text="Wyczyść przesuwane",
            command=lambda: self._clear_current_target("mover"),
        ).grid(row=0, column=2, sticky="ew", padx=(8, 0))
        ttk.Button(
            button_row,
            text="Wyczyść kotwice",
            command=lambda: self._clear_current_target("anchor"),
        ).grid(row=0, column=3, sticky="ew", padx=(8, 0))

        settings_row = ttk.Frame(current_box)
        settings_row.grid(row=3, column=0, sticky="ew", padx=8, pady=(8, 8))
        ttk.Label(settings_row, text="Kierunek:").pack(side="left")
        ttk.Combobox(
            settings_row,
            textvariable=self.direction_var,
            values=list(self.labels_to_direction.keys()),
            state="readonly",
            width=12,
        ).pack(side="left", padx=(6, 12))
        ttk.Label(settings_row, text="Min. odstęp (grid):").pack(side="left")
        ttk.Entry(settings_row, textvariable=self.gap_var, width=8).pack(
            side="left", padx=(6, 0)
        )

        simulation_row = ttk.Frame(current_box)
        simulation_row.grid(row=4, column=0, sticky="ew", padx=8, pady=(0, 8))
        ttk.Checkbutton(
            simulation_row,
            text="Symuluj dla wiersza",
            variable=self.simulate_row_var,
        ).pack(side="left")
        ttk.Label(simulation_row, text="Wiersz:").pack(side="left", padx=(10, 0))
        ttk.Entry(
            simulation_row,
            textvariable=self.preview_row_var,
            width=6,
        ).pack(side="left", padx=(6, 0))
        ttk.Button(
            simulation_row,
            text="Aktualny",
            command=self._use_current_preview_row,
        ).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(
            simulation_row,
            text="Pokaż pozycje domyślne",
            variable=self.show_default_positions_var,
        ).pack(side="left", padx=(14, 0))

        ttk.Label(
            current_box,
            textvariable=self.status_var,
            justify="left",
            foreground="#5b6576",
        ).grid(row=5, column=0, sticky="ew", padx=8, pady=(0, 8))

        actions = ttk.Frame(right)
        actions.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        actions.columnconfigure(2, weight=1)
        actions.columnconfigure(3, weight=1)
        actions.columnconfigure(4, weight=1)
        actions.columnconfigure(5, weight=1)
        ttk.Button(actions, text="Nowa", command=self.start_new_dependency).grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Button(actions, text="Dodaj", command=self.add_dependency).grid(
            row=0, column=1, sticky="ew", padx=(8, 0)
        )
        ttk.Button(actions, text="Zapisz zmiany", command=self.update_dependency).grid(
            row=0, column=2, sticky="ew", padx=(8, 0)
        )
        ttk.Button(actions, text="Usuń", command=self.remove_dependency).grid(
            row=0, column=3, sticky="ew", padx=(8, 0)
        )
        ttk.Button(actions, text="Wyżej", command=self.move_dependency_up).grid(
            row=0, column=4, sticky="ew", padx=(8, 0)
        )
        ttk.Button(actions, text="Niżej", command=self.move_dependency_down).grid(
            row=0, column=5, sticky="ew", padx=(8, 0)
        )

        list_box = ttk.LabelFrame(right, text="Lista Zależności")
        list_box.grid(row=2, column=0, sticky="nsew")
        list_box.columnconfigure(0, weight=1)
        list_box.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            list_box,
            columns=("movers", "anchors", "direction", "gap", "status"),
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("movers", text="Przesuwane")
        self.tree.heading("anchors", text="Dosuń do")
        self.tree.heading("direction", text="Kierunek")
        self.tree.heading("gap", text="Grid")
        self.tree.heading("status", text="Status")
        self.tree.column("movers", width=210, stretch=True, anchor="w")
        self.tree.column("anchors", width=210, stretch=True, anchor="w")
        self.tree.column("direction", width=90, stretch=False, anchor="center")
        self.tree.column("gap", width=60, stretch=False, anchor="center")
        self.tree.column("status", width=95, stretch=False, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._load_selected_dependency)

        tree_scroll = ttk.Scrollbar(list_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.grid(row=0, column=1, sticky="ns")

        ttk.Button(right, text="Zamknij", command=self.close).grid(
            row=3, column=0, sticky="ew", pady=(8, 0)
        )

    def refresh_from_parent(self):
        valid_names = set(self.parent.elements.keys())
        self.current_anchor_names = [
            name for name in self.current_anchor_names if name in valid_names
        ]
        self.current_mover_names = [
            name for name in self.current_mover_names if name in valid_names
        ]
        if self.active_dependency_index is not None:
            if self.active_dependency_index < 0 or self.active_dependency_index >= len(
                self.parent.dependencies
            ):
                self.active_dependency_index = None
        self._refresh_preview()

    def _rebuild_canvas_elements(self, layout_elements=None, hidden_names=None):
        self.canvas.delete("all")
        self.element_items = {}
        self.item_to_name = {}
        layout_elements = layout_elements or self.parent.elements
        hidden_names = set(hidden_names or [])

        page_width = self.parent.page_width * self.parent.scale
        page_height = self.parent.page_height * self.parent.scale
        margin = 40
        self.canvas.create_rectangle(
            0,
            0,
            page_width,
            page_height,
            fill="white",
            outline="#69778d",
            width=1,
            tags=("page",),
        )

        for name, element in sorted(
            self.parent.elements.items(), key=lambda item: item[1].layer
        ):
            layout = layout_elements.get(name) or element
            is_hidden = name in hidden_names
            fill = element.bg_color if getattr(element, "bg_visible", True) else "#fbfcfe"
            outline = "#1f2937"
            dash = ()
            label_fill = element.text_color
            label_text = self._short_label(name)
            if is_hidden:
                fill = "#f8fafc"
                outline = "#b8c1cf"
                dash = (4, 2)
                label_fill = "#8b95a5"
                label_text = f"{label_text}\n[ukryte]"
            rect = self.canvas.create_rectangle(
                layout.x,
                layout.y,
                layout.x + layout.width,
                layout.y + layout.height,
                fill=fill,
                outline=outline,
                width=1,
                dash=dash,
                tags=("dep_block",),
            )
            label = self.canvas.create_text(
                layout.x + layout.width / 2,
                layout.y + layout.height / 2,
                text=label_text,
                width=max(20, layout.width - 8),
                justify="center",
                fill=label_fill,
                tags=("dep_block",),
            )
            self.element_items[name] = {
                "rect": rect,
                "label": label,
            }
            self.item_to_name[rect] = name
            self.item_to_name[label] = name

        self.canvas.configure(
            scrollregion=(-margin, -margin, page_width + margin, page_height + margin)
        )
        self._fit_initial_view(page_width, page_height)

    def _fit_initial_view(self, page_width, page_height):
        width = min(max(int(page_width + 20), 480), 920)
        height = min(max(int(page_height + 20), 420), 720)
        self.canvas.configure(width=width, height=height)

    def _short_label(self, name):
        label = self.parent.display_name(name)
        if len(label) > 48:
            return f"{label[:45]}..."
        return label

    def _ordered_names(self, names):
        order = {name: index for index, name in enumerate(self.parent.elements.keys())}
        return sorted(set(names), key=lambda name: order.get(name, 10**9))

    def _format_names(self, names):
        labels = [
            self.parent.display_name(name)
            for name in names
            if name in self.parent.elements
        ]
        if not labels:
            return "brak"
        return ", ".join(labels)

    def _current_dependency(self):
        return normalize_dependency(
            {
                "anchor_names": list(self.current_anchor_names),
                "mover_names": list(self.current_mover_names),
                "direction": self.labels_to_direction.get(
                    self.direction_var.get(), DEPENDENCY_DIRECTION_BELOW
                ),
                "gap_steps": self.gap_var.get(),
            }
        )

    def _validated_current_dependency(self):
        dependency = self._current_dependency()
        if dependency is None:
            messagebox.showerror(
                "Błąd",
                "Najpierw wskaż co najmniej jeden blok w obu zakresach.",
                parent=self,
            )
            return None
        if set(dependency["anchor_names"]).intersection(dependency["mover_names"]):
            messagebox.showerror(
                "Błąd",
                "Ten sam blok nie może być jednocześnie kotwicą i zakresem przesuwanym.",
                parent=self,
            )
            return None
        return dependency

    def _set_capture_target(self, target):
        self.capture_target = target
        if target == "mover":
            self.current_mover_names = []
            self.mode_var.set("Tryb wyboru: przesuwane")
        else:
            self.current_anchor_names = []
            self.mode_var.set("Tryb wyboru: kotwice")
        self._refresh_preview()

    def _clear_current_target(self, target):
        self.capture_target = None
        self.mode_var.set("Tryb wyboru: brak")
        if target == "mover":
            self.current_mover_names = []
        else:
            self.current_anchor_names = []
        self._refresh_preview()

    def start_new_dependency(self):
        self.active_dependency_index = None
        self.capture_target = None
        self.mode_var.set("Tryb wyboru: brak")
        self.current_anchor_names = []
        self.current_mover_names = []
        self._set_preview_values(
            direction_label=self.direction_labels[DEPENDENCY_DIRECTION_BELOW],
            gap_value="0",
        )
        self._refresh_preview()

    def canvas_button_press(self, event):
        if not self.capture_target:
            return
        name = self._element_name_at(event)
        if name:
            additive = bool(event.state & 0x0001)
            self._apply_selection([name], additive=additive)
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.sel_start = (x, y)
        self.sel_additive = bool(event.state & 0x0001)
        if self.sel_rect:
            self.canvas.delete(self.sel_rect)
        self.sel_rect = self.canvas.create_rectangle(
            x,
            y,
            x,
            y,
            outline="#2563eb",
            dash=(3, 2),
            width=2,
        )

    def canvas_drag_select(self, event):
        if not self.sel_start or not self.sel_rect:
            return
        x0, y0 = self.sel_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.canvas.coords(self.sel_rect, x0, y0, x1, y1)

    def canvas_button_release(self, event):
        if not self.sel_start or not self.sel_rect or not self.capture_target:
            return
        x0, y0 = self.sel_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.canvas.delete(self.sel_rect)
        self.sel_rect = None
        self.sel_start = None
        if x0 > x1:
            x0, x1 = x1, x0
        if y0 > y1:
            y0, y1 = y1, y0
        selected = []
        for name, items in self.element_items.items():
            ex0, ey0, ex1, ey1 = self.canvas.coords(items["rect"])
            if ex0 >= x0 and ex1 <= x1 and ey0 >= y0 and ey1 <= y1:
                selected.append(name)
        self._apply_selection(selected, additive=self.sel_additive)

    def _apply_selection(self, names, additive=False):
        ordered = self._ordered_names(names)
        if self.capture_target == "mover":
            current = self.current_mover_names if additive else []
            self.current_mover_names = self._ordered_names(list(current) + ordered)
        elif self.capture_target == "anchor":
            current = self.current_anchor_names if additive else []
            self.current_anchor_names = self._ordered_names(list(current) + ordered)
        self._refresh_preview()

    def _element_name_at(self, event):
        current = self.canvas.find_withtag("current")
        for item in current:
            name = self.item_to_name.get(item)
            if name:
                return name
        return None

    def _build_preview_dependencies(self):
        dependencies = list(self.parent.dependencies)
        draft = self._current_dependency()
        focus_index = None
        if self.active_dependency_index is not None:
            focus_index = self.active_dependency_index
            if draft and 0 <= focus_index < len(dependencies):
                dependencies[focus_index] = draft
        elif draft:
            dependencies.append(draft)
            focus_index = len(dependencies) - 1
        return dependencies, focus_index

    def _clone_layout_elements(self):
        return {
            name: SimpleNamespace(
                x=float(element.x),
                y=float(element.y),
                width=float(element.width),
                height=float(element.height),
            )
            for name, element in self.parent.elements.items()
        }

    def _preview_hidden_names(self):
        self.simulation_error = ""
        self.simulation_row_number = None
        if not self.simulate_row_var.get() or not getattr(self.parent, "dataframes", {}):
            return set()
        try:
            row_number = int(str(self.preview_row_var.get() or "").strip())
        except (TypeError, ValueError):
            self.simulation_error = "Nieprawidłowy numer wiersza do symulacji."
            return set()
        if row_number < 1:
            self.simulation_error = "Numer wiersza symulacji musi być >= 1."
            return set()
        self.simulation_row_number = row_number
        try:
            _values, hidden = self.parent.build_preview_layout_state(idx=row_number - 1)
        except Exception as exc:
            self.simulation_error = str(exc)
            return set()
        return hidden

    def _refresh_preview(self):
        if self._refresh_in_progress:
            self._refresh_pending = True
            return
        self._refresh_in_progress = True
        try:
            self._refresh_preview_now()
        finally:
            self._refresh_in_progress = False
            if self._refresh_pending:
                self._refresh_pending = False
                self._request_preview_refresh()

    def _refresh_preview_now(self):
        self.current_anchor_names = self._ordered_names(self.current_anchor_names)
        self.current_mover_names = self._ordered_names(self.current_mover_names)

        hidden_names = self._preview_hidden_names()
        self.simulation_hidden_names = set(hidden_names)
        preview_dependencies, focus_index = self._build_preview_dependencies()
        simulated_elements = self._clone_layout_elements()
        _, statuses = apply_layout_dependencies(
            simulated_elements,
            dependencies=preview_dependencies,
            hidden_names=hidden_names,
            grid_step=self.parent.grid_size * self.parent.scale,
            return_statuses=True,
        )
        if self.show_default_positions_var.get():
            display_elements = self._clone_layout_elements()
        else:
            display_elements = simulated_elements
        self.preview_statuses = statuses
        self.preview_focus_index = focus_index
        self._rebuild_canvas_elements(
            layout_elements=display_elements,
            hidden_names=hidden_names,
        )
        self._refresh_tree()
        self._redraw_dependency_arrows()
        self._update_block_highlights()
        self._update_status_text()

    def _refresh_tree(self):
        self._suspend_tree_event = True
        try:
            self.tree.delete(*self.tree.get_children())
            statuses = self.preview_statuses[: len(self.parent.dependencies)]
            for index, dependency in enumerate(self.parent.dependencies):
                status = statuses[index] if index < len(statuses) else None
                self.tree.insert(
                    "",
                    "end",
                    iid=str(index),
                    values=(
                        self._format_names(dependency["mover_names"]),
                        self._format_names(dependency["anchor_names"]),
                        self.direction_labels.get(
                            dependency["direction"], dependency["direction"]
                        ),
                        dependency["gap_steps"],
                        self._status_label(status),
                    ),
                )
            if self.active_dependency_index is not None:
                item_id = str(self.active_dependency_index)
                if self.tree.exists(item_id):
                    self.tree.selection_set(item_id)
                    self.tree.focus(item_id)
                else:
                    self.active_dependency_index = None
        finally:
            self._suspend_tree_event = False

    def _status_label(self, status):
        if not status:
            return ""
        if status["status"] == DEPENDENCY_STATUS_APPLIED:
            return "OK"
        if status["status"] == DEPENDENCY_STATUS_ALREADY_ALIGNED:
            return "Już"
        if status["status"] == DEPENDENCY_STATUS_BLOCKED:
            return "Kolizja"
        return "Pominięta"

    def _redraw_dependency_arrows(self):
        self.canvas.delete("dep_arrow")
        for index, status in enumerate(self.preview_statuses):
            anchor_box = status.get("anchor_box")
            mover_box = status.get("mover_box_after") or status.get("mover_box_before")
            if not anchor_box or not mover_box:
                continue
            start_x, start_y, end_x, end_y = self._arrow_points(
                anchor_box,
                mover_box,
                status.get("direction"),
            )
            color = "#8a94a6"
            width = 2
            dash = ()
            if status["status"] == DEPENDENCY_STATUS_BLOCKED:
                color = "#d64545"
                dash = (6, 3)
            if index == self.preview_focus_index:
                color = "#0b66c3" if status["status"] != DEPENDENCY_STATUS_BLOCKED else "#c53f3f"
                width = 3

            self.canvas.create_line(
                start_x,
                start_y,
                end_x,
                end_y,
                arrow="last",
                fill=color,
                width=width,
                dash=dash,
                tags=("dep_arrow",),
            )
            label_x = (start_x + end_x) / 2
            label_y = (start_y + end_y) / 2 - 10
            self.canvas.create_text(
                label_x,
                label_y,
                text=f"{self.direction_labels.get(status['direction'], status['direction'])} / {status['gap_steps']}",
                fill=color,
                font=("TkDefaultFont", 8),
                tags=("dep_arrow",),
            )

    def _arrow_points(self, anchor_box, mover_box, direction):
        anchor_center_x = (anchor_box["left"] + anchor_box["right"]) / 2
        anchor_center_y = (anchor_box["top"] + anchor_box["bottom"]) / 2
        mover_center_x = (mover_box["left"] + mover_box["right"]) / 2
        mover_center_y = (mover_box["top"] + mover_box["bottom"]) / 2

        if direction == "below":
            return anchor_center_x, anchor_box["bottom"], mover_center_x, mover_box["top"]
        if direction == "above":
            return anchor_center_x, anchor_box["top"], mover_center_x, mover_box["bottom"]
        if direction == "right":
            return anchor_box["right"], anchor_center_y, mover_box["left"], mover_center_y
        return anchor_box["left"], anchor_center_y, mover_box["right"], mover_center_y

    def _update_block_highlights(self):
        focus_status = None
        show_blockers = False
        if self.preview_focus_index is not None and 0 <= self.preview_focus_index < len(
            self.preview_statuses
        ):
            focus_status = self.preview_statuses[self.preview_focus_index]
            show_blockers = (
                self.active_dependency_index is not None
                and self.capture_target is None
                and self.preview_focus_index == self.active_dependency_index
            )
        blocked_names = (
            set(focus_status.get("blocked_by", []))
            if focus_status and show_blockers
            else set()
        )

        for name, items in self.element_items.items():
            outline = "#1f2937"
            width = 1
            fill = self.parent.elements[name].bg_color if getattr(
                self.parent.elements[name], "bg_visible", True
            ) else "#fbfcfe"
            dash = ()
            if name in self.simulation_hidden_names:
                fill = "#f8fafc"
                outline = "#b8c1cf"
                dash = (4, 2)

            if name in self.current_anchor_names:
                outline = "#2563eb"
                width = 3 if self.capture_target == "anchor" else 2
            if name in self.current_mover_names:
                outline = "#d97706"
                width = 3 if self.capture_target == "mover" else 2
            if name in blocked_names:
                outline = "#c53f3f"
                width = 3

            self.canvas.itemconfig(
                items["rect"],
                outline=outline,
                width=width,
                fill=fill,
                dash=dash,
            )

    def _update_status_text(self):
        self.mover_summary_var.set(
            f"Przesuwane: {self._format_names(self.current_mover_names)}"
        )
        self.anchor_summary_var.set(
            f"Kotwice: {self._format_names(self.current_anchor_names)}"
        )

        prefix = []
        if self.simulate_row_var.get():
            if self.simulation_error:
                prefix.append(self.simulation_error)
            elif self.simulation_row_number is not None:
                prefix.append(
                    f"Symulacja dla wiersza {self.simulation_row_number}: ukryte {len(self.simulation_hidden_names)} bloków."
                )
        else:
            prefix.append("Symulacja wiersza wyłączona.")
        if self.show_default_positions_var.get():
            prefix.append("Na planszy pokazane są pozycje domyślne.")

        if self.preview_focus_index is not None and 0 <= self.preview_focus_index < len(
            self.preview_statuses
        ):
            status = self.preview_statuses[self.preview_focus_index]
            if status["status"] == DEPENDENCY_STATUS_BLOCKED:
                names = self._format_names(status.get("blocked_by", []))
                prefix.append(
                    f"Nie znaleziono wolnej pozycji bez kolizji z: {names}."
                )
                self.status_var.set(" ".join(prefix))
                return
            if status["status"] == DEPENDENCY_STATUS_APPLIED:
                if status.get("adjusted"):
                    prefix.append(
                        "Zakres zostanie dosunięty do najbliższej wolnej pozycji zgodnej z kierunkiem."
                    )
                else:
                    prefix.append(
                        "Zależność może zostać zastosowana bez dodatkowych przesunięć omijających kolizje."
                    )
                self.status_var.set(" ".join(prefix))
                return
            if status["status"] == DEPENDENCY_STATUS_ALREADY_ALIGNED:
                prefix.append(
                    "Bloki już są ustawione zgodnie z tą zależnością."
                )
                self.status_var.set(" ".join(prefix))
                return
        prefix.append("Podgląd pokazuje możliwy układ końcowy dla bieżących ustawień.")
        self.status_var.set(" ".join(prefix))

    def _load_selected_dependency(self, _event=None):
        if self._suspend_tree_event:
            return
        selection = self.tree.selection()
        if not selection:
            self.active_dependency_index = None
            self._refresh_preview()
            return
        try:
            index = int(selection[0])
        except (TypeError, ValueError):
            return
        if not (0 <= index < len(self.parent.dependencies)):
            return

        dependency = self.parent.dependencies[index]
        direction_label = self.direction_labels.get(
            dependency["direction"],
            self.direction_labels[DEPENDENCY_DIRECTION_BELOW],
        )
        gap_text = str(dependency["gap_steps"])
        anchor_names = list(dependency["anchor_names"])
        mover_names = list(dependency["mover_names"])
        if (
            self.active_dependency_index == index
            and self.capture_target is None
            and self.current_anchor_names == anchor_names
            and self.current_mover_names == mover_names
            and self.direction_var.get() == direction_label
            and self.gap_var.get() == gap_text
        ):
            return
        self.active_dependency_index = index
        self.capture_target = None
        self.mode_var.set("Tryb wyboru: brak")
        self.current_anchor_names = anchor_names
        self.current_mover_names = mover_names
        self._set_preview_values(
            direction_label=direction_label,
            gap_value=gap_text,
        )
        self._request_preview_refresh()

    def add_dependency(self):
        dependency = self._validated_current_dependency()
        if dependency is None:
            return
        self.parent.dependencies.append(dependency)
        self.active_dependency_index = len(self.parent.dependencies) - 1
        self.capture_target = None
        self.mode_var.set("Tryb wyboru: brak")
        self.parent.push_history()
        self._refresh_preview()

    def update_dependency(self):
        dependency = self._validated_current_dependency()
        if dependency is None:
            return
        if self.active_dependency_index is None or not (
            0 <= self.active_dependency_index < len(self.parent.dependencies)
        ):
            messagebox.showerror(
                "Błąd",
                "Najpierw wybierz istniejącą zależność do edycji albo użyj Dodaj.",
                parent=self,
            )
            return
        self.parent.dependencies[self.active_dependency_index] = dependency
        self.capture_target = None
        self.mode_var.set("Tryb wyboru: brak")
        self.parent.push_history()
        self._refresh_preview()

    def remove_dependency(self):
        if self.active_dependency_index is None or not (
            0 <= self.active_dependency_index < len(self.parent.dependencies)
        ):
            return
        self.parent.dependencies.pop(self.active_dependency_index)
        self.start_new_dependency()
        self.parent.push_history()
        self._refresh_preview()

    def move_dependency_up(self):
        if self.active_dependency_index is None or self.active_dependency_index <= 0:
            return
        dependencies = self.parent.dependencies
        index = self.active_dependency_index
        dependencies[index - 1], dependencies[index] = dependencies[index], dependencies[index - 1]
        self.active_dependency_index = index - 1
        self.parent.push_history()
        self._refresh_preview()

    def move_dependency_down(self):
        if self.active_dependency_index is None:
            return
        if self.active_dependency_index >= len(self.parent.dependencies) - 1:
            return
        dependencies = self.parent.dependencies
        index = self.active_dependency_index
        dependencies[index + 1], dependencies[index] = dependencies[index], dependencies[index + 1]
        self.active_dependency_index = index + 1
        self.parent.push_history()
        self._refresh_preview()

    def close(self):
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except tk.TclError:
                pass
            self._refresh_after_id = None
        self.parent.dependencies_win = None
        self.destroy()
