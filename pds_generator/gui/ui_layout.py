import tkinter as tk
from tkinter import ttk

from ..pdf_settings import DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT


def setup_ui(app):
    top_frame = ttk.Frame(app, padding=(8, 8, 8, 4), style="Toolbar.TFrame")
    top_frame.pack(fill="x")

    source_card = ttk.LabelFrame(
        top_frame,
        text="Źródło i strona",
        style="Card.TLabelframe",
    )
    source_card.pack(side="left", fill="x", expand=True, padx=(0, 8))
    source_card.columnconfigure(1, weight=1)

    ttk.Label(source_card, text="Excel:").grid(row=0, column=0, sticky="w", padx=4, pady=2)
    app.path_var = tk.StringVar()
    app.path_entry = ttk.Entry(source_card, textvariable=app.path_var, width=46)
    app.path_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=2)
    browse_btn = ttk.Button(source_card, text="Przeglądaj", command=app.browse_file)
    browse_btn.grid(row=0, column=2, sticky="ew", padx=(6, 4), pady=2)
    ttk.Label(source_card, text="Strona:").grid(
        row=0, column=3, sticky="w", padx=(10, 4), pady=2
    )
    app.size_var = tk.StringVar(value="A4")
    app.size_entry = ttk.Entry(source_card, textvariable=app.size_var, width=8)
    app.size_entry.grid(row=0, column=4, sticky="w", padx=4, pady=2)
    size_btn = ttk.Button(source_card, text="Ustaw", command=app.update_canvas_size)
    size_btn.grid(row=0, column=5, sticky="ew", padx=(6, 4), pady=2)
    app.size_entry.bind("<Return>", lambda e: app.update_canvas_size())

    update_card = ttk.LabelFrame(
        top_frame,
        text="Wersja",
        style="Card.TLabelframe",
    )
    update_card.pack(side="right", fill="y")
    update_frame = ttk.Frame(update_card, style="Panel.TFrame")
    update_frame.pack(fill="x")
    app.update_info_var = tk.StringVar(value=f"Aktualna wersja: {app.version}")
    ttk.Label(update_frame, textvariable=app.update_info_var).pack(
        anchor="w", padx=4, pady=(0, 6)
    )
    buttons_row = ttk.Frame(update_frame, style="Panel.TFrame")
    buttons_row.pack(fill="x")
    app.update_button = tk.Button(
        buttons_row,
        text="Aktualizuj",
        command=app.manual_update,
        padx=8,
        pady=2,
        highlightthickness=0,
    )
    app.update_button_bg = app.update_button.cget("background")
    app.update_button.pack(side="left")
    if app.github_image:
        app.github_button = tk.Button(
            buttons_row,
            image=app.github_image,
            command=app.open_github,
            borderwidth=0,
            padx=6,
            pady=2,
            highlightthickness=0,
        )
    else:
        app.github_button = tk.Button(
            buttons_row,
            text="GitHub",
            command=app.open_github,
            padx=8,
            pady=2,
            highlightthickness=0,
        )
    app.github_button.pack(side="left", padx=(6, 0))

    format_frame = ttk.Frame(app, padding=(8, 0, 8, 5), style="Toolbar.TFrame")
    format_frame.pack(fill="x")
    format_row = ttk.Frame(format_frame, padding=(10, 7), style="Panel.TFrame")
    format_row.pack(fill="x")
    ttk.Label(
        format_row,
        text="Element:",
        style="Panel.TLabel",
    ).pack(side="left", padx=(0, 8))

    bold_btn = ttk.Button(
        format_row,
        text="B",
        width=3,
        style="Compact.TButton",
        command=app.toggle_bold,
    )
    bold_btn.pack(side="left")
    inc_btn = ttk.Button(
        format_row,
        text="A+",
        width=3,
        style="Compact.TButton",
        command=app.increase_font,
    )
    inc_btn.pack(side="left", padx=(3, 0))
    dec_btn = ttk.Button(
        format_row,
        text="A-",
        width=3,
        style="Compact.TButton",
        command=app.decrease_font,
    )
    dec_btn.pack(side="left", padx=(3, 0))
    app.font_size_var = tk.StringVar()
    app.font_entry = ttk.Entry(
        format_row,
        textvariable=app.font_size_var,
        width=4,
        state="disabled",
    )
    app.font_entry.pack(side="left", padx=(6, 0))
    app.font_entry.bind("<Return>", lambda e: app.set_font_size())
    app.auto_font_var = tk.BooleanVar(value=True)
    app.auto_font_check = ttk.Checkbutton(
        format_row,
        text="Auto",
        variable=app.auto_font_var,
        command=app.toggle_auto_font,
    )
    app.auto_font_check.pack(side="left", padx=(6, 0))
    app.auto_font_check.state(["disabled"])
    app.image_auto_zoom_var = tk.BooleanVar(value=False)
    app.image_auto_zoom_check = ttk.Checkbutton(
        format_row,
        text="Zoom IMG",
        variable=app.image_auto_zoom_var,
        command=app.toggle_image_auto_zoom,
    )
    app.image_auto_zoom_check.pack(side="left", padx=(6, 0))
    app.image_auto_zoom_check.state(["disabled"])
    ttk.Separator(format_row, orient="vertical").pack(
        side="left", fill="y", padx=8, pady=2
    )
    text_color_btn = ttk.Button(
        format_row,
        text="Tekst",
        style="Compact.TButton",
        command=app.choose_text_color,
    )
    text_color_btn.pack(side="left")
    bg_color_btn = ttk.Button(
        format_row,
        text="Tło",
        style="Compact.TButton",
        command=app.choose_bg_color,
    )
    bg_color_btn.pack(side="left", padx=(3, 0))
    app.transparent_var = tk.BooleanVar(value=False)
    app.bg_check = ttk.Checkbutton(
        format_row,
        text="Przezr.",
        variable=app.transparent_var,
        command=app.toggle_bg_visible,
    )
    app.bg_check.pack(side="left", padx=(6, 0))
    app.bg_check.state(["disabled"])
    ttk.Separator(format_row, orient="vertical").pack(
        side="left", fill="y", padx=8, pady=2
    )
    align_left_btn = ttk.Button(
        format_row,
        text="L",
        width=3,
        style="Compact.TButton",
        command=lambda: app.set_alignment("left"),
    )
    align_left_btn.pack(side="left")
    align_center_btn = ttk.Button(
        format_row,
        text="C",
        width=3,
        style="Compact.TButton",
        command=lambda: app.set_alignment("center"),
    )
    align_center_btn.pack(side="left", padx=(3, 0))
    align_right_btn = ttk.Button(
        format_row,
        text="P",
        width=3,
        style="Compact.TButton",
        command=lambda: app.set_alignment("right"),
    )
    align_right_btn.pack(side="left", padx=(3, 0))
    center_h_btn = ttk.Button(
        format_row,
        text="X",
        width=3,
        style="Compact.TButton",
        command=app.center_selected_horizontal,
    )
    center_h_btn.pack(side="left", padx=(6, 0))
    center_v_btn = ttk.Button(
        format_row,
        text="Y",
        width=3,
        style="Compact.TButton",
        command=app.center_selected_vertical,
    )
    center_v_btn.pack(side="left", padx=(3, 0))
    ttk.Separator(format_row, orient="vertical").pack(
        side="left", fill="y", padx=8, pady=2
    )
    ttk.Label(format_row, text="Warstwa:", style="Panel.TLabel").pack(side="left")
    app.layer_var = tk.StringVar()
    app.layer_entry = ttk.Entry(
        format_row,
        textvariable=app.layer_var,
        width=4,
        state="disabled",
    )
    app.layer_entry.pack(side="left", padx=(6, 0))
    app.layer_entry.bind("<Return>", lambda e: app.set_layer())
    ttk.Separator(format_row, orient="vertical").pack(
        side="left", fill="y", padx=8, pady=2
    )
    ttk.Label(format_row, text="Źródło:", style="Panel.TLabel").pack(side="left")
    app.value_source_var = tk.StringVar(value="")
    app.value_source_combo = ttk.Combobox(
        format_row,
        textvariable=app.value_source_var,
        values=list(app.VALUE_SOURCE_LABELS.values()),
        width=24,
        state="disabled",
    )
    app.value_source_combo.pack(side="left", padx=(6, 0))
    app.value_source_combo.bind("<<ComboboxSelected>>", app.on_value_source_changed)

    workspace = ttk.Panedwindow(app, orient="horizontal")
    workspace.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    app.canvas_container = tk.Frame(workspace, bg="#d4dae4")
    app.canvas_container.pack_propagate(False)
    app.canvas_container.bind("<Configure>", app.schedule_resize_canvas)
    app.canvas = tk.Canvas(
        app.canvas_container,
        bg="#d4dae4",
        highlightthickness=0,
    )
    app.canvas.pack(fill="both", expand=True)
    app.canvas.bind("<ButtonPress-1>", app.canvas_button_press)
    app.canvas.bind("<B1-Motion>", app.canvas_drag_select)
    app.canvas.bind("<ButtonRelease-1>", app.canvas_button_release)
    app.canvas.bind("<Control-MouseWheel>", app.ctrl_zoom)
    app.canvas.bind("<Control-Button-4>", lambda e: app.ctrl_zoom(e, 120))
    app.canvas.bind("<Control-Button-5>", lambda e: app.ctrl_zoom(e, -120))
    app.canvas.bind("<ButtonPress-2>", app.start_pan)
    app.canvas.bind("<B2-Motion>", app.pan_canvas)
    app.canvas.configure(
        scrollregion=(
            -app.margin,
            -app.margin,
            app.page_width + app.margin,
            app.page_height + app.margin,
        )
    )

    zoom_frame = ttk.Frame(app.canvas_container, padding=(6, 4), style="Panel.TFrame")
    zoom_frame.place(relx=1.0, rely=1.0, anchor="se", x=-8, y=-8)
    ttk.Button(
        zoom_frame,
        text="Dopasuj",
        style="Compact.TButton",
        command=app.fit_to_window,
    ).pack(
        side="right"
    )
    app.zoom_var = tk.StringVar(value="100%")
    ttk.Label(zoom_frame, textvariable=app.zoom_var, style="Panel.TLabel").pack(
        side="right", padx=(0, 8)
    )

    right_container = ttk.Frame(workspace, width=296, style="Toolbar.TFrame")
    workspace.add(app.canvas_container, weight=5)
    workspace.add(right_container, weight=1)

    app.right_canvas = tk.Canvas(
        right_container,
        width=296,
        bg=app.cget("background"),
        highlightthickness=0,
        bd=0,
    )
    right_scroll = ttk.Scrollbar(
        right_container,
        orient="vertical",
        command=app.right_canvas.yview,
    )
    app.right_canvas.configure(yscrollcommand=right_scroll.set)
    right_scroll.pack(side="right", fill="y")
    app.right_canvas.pack(side="left", fill="both", expand=True)

    right_frame = ttk.Frame(
        app.right_canvas,
        padding=(2, 2, 2, 10),
        style="Toolbar.TFrame",
    )
    right_window = app.right_canvas.create_window((0, 0), window=right_frame, anchor="nw")

    def sync_right_panel(_event=None):
        app.right_canvas.configure(scrollregion=app.right_canvas.bbox("all"))

    def fit_right_panel_width(event):
        pending = getattr(app, "_right_panel_resize_after_id", None)
        if pending is not None:
            try:
                app.after_cancel(pending)
            except tk.TclError:
                pass

        width = max(event.width, 1)

        def apply_width():
            app._right_panel_resize_after_id = None
            app.right_canvas.itemconfigure(right_window, width=width)
            sync_right_panel()

        app._right_panel_resize_after_id = app.after(120, apply_width)

    right_frame.bind("<Configure>", sync_right_panel)
    app.right_canvas.bind("<Configure>", fit_right_panel_width)
    app.right_canvas.bind(
        "<Enter>",
        lambda e: app.right_canvas.bind_all("<MouseWheel>", app._on_mousewheel),
    )
    app.right_canvas.bind(
        "<Leave>",
        lambda e: app.right_canvas.unbind_all("<MouseWheel>"),
    )
    app.right_canvas.bind(
        "<Button-4>",
        lambda e: app.right_canvas.yview_scroll(-1, "units"),
    )
    app.right_canvas.bind(
        "<Button-5>",
        lambda e: app.right_canvas.yview_scroll(1, "units"),
    )
    app.panel_bg = app.right_canvas.cget("background") or app.cget("background")

    columns_section = ttk.LabelFrame(
        right_frame,
        text="Kolumny z Excela",
        style="Card.TLabelframe",
    )
    columns_section.pack(fill="x", pady=(0, 8))
    app.columns_frame = ttk.Frame(columns_section, style="Panel.TFrame")
    app.columns_frame.pack(fill="x")
    app.columns_vars = {}
    column_actions = ttk.Frame(columns_section, style="Panel.TFrame")
    column_actions.pack(fill="x", pady=(8, 0))
    app.tracking_btn = ttk.Button(
        column_actions,
        text="Śledzenie zmian",
        command=app.open_tracking_settings,
    )
    app.tracking_btn.pack(side="left", fill="x", expand=True)
    app.mail_settings_btn = ttk.Button(
        column_actions,
        text="Konfiguracja e-mail",
        command=app.open_mail_settings,
    )
    app.mail_settings_btn.pack(side="left", fill="x", expand=True, padx=(8, 0))

    images_section = ttk.LabelFrame(
        right_frame,
        text="Foldery obrazów",
        style="Card.TLabelframe",
    )
    images_section.pack(fill="x", pady=(0, 8))
    app.image_dirs_list = tk.Listbox(images_section, height=3)
    app.image_dirs_list.pack(fill="x", pady=(0, 6))
    img_btns = ttk.Frame(images_section, style="Panel.TFrame")
    img_btns.pack(fill="x")
    add_img_btn = ttk.Button(img_btns, text="Dodaj folder", command=app.add_image_dir)
    add_img_btn.pack(side="left", fill="x", expand=True)
    remove_img_btn = ttk.Button(
        img_btns,
        text="Usuń zaznaczone",
        command=app.remove_image_dir,
    )
    remove_img_btn.pack(side="left", fill="x", expand=True, padx=(8, 0))
    app.image_index_btn = ttk.Button(
        images_section,
        text="Aktualizuj indeks obrazów",
        command=app.rebuild_image_index,
    )
    app.image_index_btn.pack(fill="x", pady=(8, 2))
    app.image_index_status_var = tk.StringVar(value="")
    app.image_index_status_label = ttk.Label(
        images_section,
        textvariable=app.image_index_status_var,
        anchor="w",
        justify="left",
    )
    app.image_index_status_label.pack(fill="x")

    pdf_section = ttk.LabelFrame(
        right_frame,
        text="PDF",
        style="Card.TLabelframe",
    )
    pdf_section.pack(fill="x", pady=(0, 8))
    ttk.Label(pdf_section, text="Kompresja obrazów:").pack(anchor="w")
    compression_row = ttk.Frame(pdf_section, style="Panel.TFrame")
    compression_row.pack(fill="x", pady=(6, 0))
    app.pdf_image_compression_var = tk.IntVar(
        value=getattr(
            app,
            "pdf_image_compression_percent",
            DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
        )
    )
    app.pdf_image_compression_scale = ttk.Scale(
        compression_row,
        from_=0,
        to=100,
        orient="horizontal",
        variable=app.pdf_image_compression_var,
        command=app.on_pdf_image_compression_changed,
    )
    app.pdf_image_compression_scale.pack(side="left", fill="x", expand=True)
    app.pdf_image_compression_value_var = tk.StringVar(value="")
    ttk.Label(
        compression_row,
        textvariable=app.pdf_image_compression_value_var,
        width=5,
        anchor="e",
    ).pack(side="left", padx=(8, 0))
    app.pdf_image_compression_hint_var = tk.StringVar(value="")
    ttk.Label(
        pdf_section,
        textvariable=app.pdf_image_compression_hint_var,
        anchor="w",
        justify="left",
    ).pack(fill="x", pady=(6, 0))
    ttk.Label(
        pdf_section,
        text="0% = oryginał, 100% = mniejszy plik",
        justify="left",
    ).pack(fill="x", pady=(4, 0))
    app.set_pdf_image_compression_percent(
        getattr(
            app,
            "pdf_image_compression_percent",
            DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
        )
    )

    static_section = ttk.LabelFrame(
        right_frame,
        text="Pola statyczne",
        style="Card.TLabelframe",
    )
    static_section.pack(fill="x", pady=(0, 8))
    app.static_frame = ttk.Frame(static_section, style="Panel.TFrame")
    app.static_frame.pack(fill="x")
    app.static_vars = {}
    app.static_entries = {}
    app.static_rows = {}
    for field in app.DEFAULT_STATIC_FIELDS:
        app.create_static_row(field, "")
    app.add_static_btn = ttk.Button(
        app.static_frame,
        text="Dodaj pole",
        command=app.add_static_field,
    )
    app.add_static_btn.pack(fill="x", pady=(8, 0))

    preview_section = ttk.LabelFrame(
        right_frame,
        text="Podgląd danych",
        style="Card.TLabelframe",
    )
    preview_section.pack(fill="x", pady=(0, 8))
    preview_frame = ttk.Frame(preview_section, style="Panel.TFrame")
    preview_frame.pack(fill="x")
    ttk.Label(preview_frame, text="Numer wiersza:").pack(side="left")
    app.row_var = tk.StringVar(value="1")
    row_entry = ttk.Entry(preview_frame, textvariable=app.row_var, width=5)
    row_entry.pack(side="left", padx=(6, 0))
    row_entry.bind("<Return>", lambda e: app.preview_row())
    app.preview_btn = ttk.Button(
        preview_frame,
        text="Podgląd",
        command=app.preview_row,
    )
    app.preview_btn.pack(side="left", padx=(8, 0))
    app.preview_status_var = tk.StringVar(value="")
    app.preview_status_label = ttk.Label(
        preview_section,
        textvariable=app.preview_status_var,
        justify="left",
    )
    app.preview_status_label.pack(fill="x", pady=(8, 0))

    groups_section = ttk.LabelFrame(
        right_frame,
        text="Grupy",
        style="Card.TLabelframe",
    )
    groups_section.pack(fill="x", pady=(0, 8))
    grp_container = ttk.Frame(groups_section, style="Panel.TFrame")
    grp_container.pack(fill="x")
    app.groups_list = tk.Listbox(grp_container, height=5)
    app.groups_list.pack(side="left", fill="both", expand=True)
    grp_scroll = ttk.Scrollbar(
        grp_container,
        orient="vertical",
        command=app.groups_list.yview,
    )
    grp_scroll.pack(side="right", fill="y")
    app.groups_list.configure(yscrollcommand=grp_scroll.set)
    app.groups_list.bind("<Double-1>", lambda e: app.edit_selected_group())
    remove_group_btn = ttk.Button(
        groups_section,
        text="Usuń grupę",
        command=app.remove_group,
    )
    remove_group_btn.pack(fill="x", pady=(8, 0))

    actions_section = ttk.LabelFrame(
        right_frame,
        text="Akcje",
        style="Card.TLabelframe",
    )
    actions_section.pack(fill="x", pady=(0, 8))
    save_btn = ttk.Button(
        actions_section,
        text="Zapisz konfigurację",
        command=app.save_config,
    )
    save_btn.pack(fill="x")
    conditions_btn = ttk.Button(
        actions_section,
        text="Warunki",
        command=app.open_conditions,
    )
    conditions_btn.pack(fill="x", pady=(8, 0))
    dependencies_btn = ttk.Button(
        actions_section,
        text="Edytor zależności",
        command=app.open_dependencies_editor,
    )
    dependencies_btn.pack(fill="x", pady=(8, 0))
    add_group_btn = ttk.Button(
        actions_section,
        text="Dodaj grupę",
        command=app.add_group,
    )
    add_group_btn.pack(fill="x", pady=(8, 0))
    gen_frame = ttk.Frame(actions_section, style="Panel.TFrame")
    gen_frame.pack(fill="x", pady=(8, 0))
    app.generate_btn = ttk.Button(
        gen_frame,
        text="Generuj PDS",
        style="Action.TButton",
        command=app.generate_pds,
    )
    app.generate_btn.pack(fill="x")
    app.cancel_btn = ttk.Button(
        gen_frame,
        text="Anuluj",
        command=app.cancel_generation,
    )
    app.cancel_btn.pack(fill="x", pady=(6, 0))
    app.cancel_btn.pack_forget()

    status_section = ttk.LabelFrame(
        right_frame,
        text="Status",
        style="Card.TLabelframe",
    )
    status_section.pack(fill="x")
    app.status_var = tk.StringVar(value="Akcja: Gotowe")
    app.status_label = ttk.Label(
        status_section,
        textvariable=app.status_var,
        anchor="w",
        justify="left",
    )
    app.status_label.pack(fill="x")
    app.progress = ttk.Progressbar(
        status_section,
        orient="horizontal",
        mode="determinate",
    )
    app.progress.pack(fill="x", pady=(8, 0))
    app.progress_info_var = tk.StringVar(value="")
    app.progress_info_label = ttk.Label(
        status_section,
        textvariable=app.progress_info_var,
        anchor="w",
        justify="left",
    )
    app.progress_info_label.pack(fill="x", pady=(6, 0))
    app.skip_var = tk.StringVar(value="")
    app.skip_label = ttk.Label(
        status_section,
        textvariable=app.skip_var,
        anchor="w",
        justify="left",
    )
    app.skip_label.pack(fill="x")
    app.rows_var = tk.StringVar(value="")
    app.rows_label = ttk.Label(
        status_section,
        textvariable=app.rows_var,
        anchor="w",
        justify="left",
    )
    app.rows_label.pack(fill="x")
    app.time_label = ttk.Label(status_section, text="")
    app.time_label.pack(anchor="w", pady=(4, 0))

    app.draw_grid()
    app.bind_all("<Delete>", app.delete_selected)

    if hasattr(app, "tooltip"):
        app.tooltip.bind(app.path_entry, text="Ścieżka do pliku Excel")
        app.tooltip.bind(browse_btn, text="Wybierz plik Excel")
        app.tooltip.bind(app.size_entry, text="Rozmiar strony (np. A4, B5)")
        app.tooltip.bind(size_btn, text="Ustaw rozmiar strony")
        app.tooltip.bind(bold_btn, text="Pogrubienie")
        app.tooltip.bind(inc_btn, text="Zwiększ rozmiar czcionki")
        app.tooltip.bind(dec_btn, text="Zmniejsz rozmiar czcionki")
        app.tooltip.bind(app.font_entry, text="Rozmiar czcionki (pt)")
        app.tooltip.bind(
            app.auto_font_check, text="Auto dopasuj rozmiar czcionki"
        )
        app.tooltip.bind(
            app.image_auto_zoom_check,
            text="Dla obrazów: przytnij białe marginesy i powiększ bez ucinania produktu",
        )
        app.tooltip.bind(text_color_btn, text="Kolor tekstu")
        app.tooltip.bind(bg_color_btn, text="Kolor tła")
        app.tooltip.bind(app.bg_check, text="Przezroczyste tło")
        app.tooltip.bind(align_left_btn, text="Wyrównaj tekst do lewej")
        app.tooltip.bind(align_center_btn, text="Wyśrodkuj tekst")
        app.tooltip.bind(align_right_btn, text="Wyrównaj tekst do prawej")
        app.tooltip.bind(center_h_btn, text="Wyśrodkuj element w poziomie")
        app.tooltip.bind(center_v_btn, text="Wyśrodkuj element w pionie")
        app.tooltip.bind(
            app.layer_entry,
            text="Warstwa elementu (wyższa = na wierzchu)",
        )
        app.tooltip.bind(
            app.value_source_combo,
            text="Źródło wartości elementu: domyślne lub data utworzenia/modyfikacji pliku Excel",
        )
        app.tooltip.bind(
            app.tracking_btn,
            text="Wybierz kolumny do śledzenia zmian",
        )
        app.tooltip.bind(
            app.mail_settings_btn,
            text="Ustaw SMTP i odbiorców raportów",
        )
        app.tooltip.bind(add_img_btn, text="Dodaj katalog z obrazami")
        app.tooltip.bind(
            remove_img_btn,
            text="Usuń zaznaczony katalog z listy",
        )
        app.tooltip.bind(
            app.image_index_btn,
            text="Przeskanuj katalogi i odśwież indeks obrazów",
        )
        app.tooltip.bind(
            app.pdf_image_compression_scale,
            text="Ustaw procent kompresji obrazów w PDF",
        )
        app.tooltip.bind(app.add_static_btn, text="Dodaj nowe pole statyczne")
        app.tooltip.bind(
            app.preview_btn,
            text="Podgląd wiersza z Excela oraz wierny podgląd PDF bezpośrednio na stronie",
        )
        app.tooltip.bind(remove_group_btn, text="Usuń wybraną grupę")
        app.tooltip.bind(save_btn, text="Zapisz konfigurację do pliku")
        app.tooltip.bind(conditions_btn, text="Edytuj warunki widoczności")
        app.tooltip.bind(
            dependencies_btn,
            text="Skonfiguruj łączniki dosuwające zakresy bloków po ukryciu elementów",
        )
        app.tooltip.bind(add_group_btn, text="Dodaj nową grupę pól")
        app.tooltip.bind(app.generate_btn, text="Wygeneruj PDS")
        app.tooltip.bind(app.cancel_btn, text="Przerwij generowanie")
