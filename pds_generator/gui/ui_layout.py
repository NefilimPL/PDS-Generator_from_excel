import tkinter as tk
from tkinter import ttk


def setup_ui(app):
    top_frame = ttk.Frame(app)
    top_frame.pack(fill="x", padx=5, pady=5)

    ttk.Label(top_frame, text="Plik Excel:").pack(side="left")
    app.path_var = tk.StringVar()
    app.path_entry = ttk.Entry(top_frame, textvariable=app.path_var, width=60)
    app.path_entry.pack(side="left", padx=5)
    ttk.Button(top_frame, text="Przeglądaj", command=app.browse_file).pack(side="left")

    ttk.Label(top_frame, text="Rozmiar strony:").pack(side="left", padx=(20, 0))
    app.size_var = tk.StringVar(value="A4")
    app.size_entry = ttk.Entry(top_frame, textvariable=app.size_var, width=10)
    app.size_entry.pack(side="left")
    ttk.Button(top_frame, text="Ustaw", command=app.update_canvas_size).pack(side="left", padx=2)
    app.size_entry.bind("<Return>", lambda e: app.update_canvas_size())

    update_frame = ttk.Frame(top_frame)
    update_frame.pack(side="right")
    app.update_info_var = tk.StringVar(
        value=f"Ostatnia aktualizacja: {app.last_update} ({app.version})"
    )
    ttk.Label(update_frame, textvariable=app.update_info_var).pack(side="left", padx=5)
    app.update_button = tk.Button(
        update_frame, text="UPDATE NOW", command=app.manual_update
    )
    app.update_button_bg = app.update_button.cget("background")
    if app.github_image:
        app.github_button = tk.Button(
            update_frame,
            image=app.github_image,
            command=app.open_github,
            borderwidth=0,
        )
    else:
        app.github_button = tk.Button(
            update_frame, text="GitHub", command=app.open_github
        )
    app.github_button.pack(side="left", padx=5)

    format_frame = ttk.Frame(app)
    format_frame.pack(fill="x", padx=5)
    ttk.Button(format_frame, text="B", command=app.toggle_bold).pack(side="left")
    ttk.Button(format_frame, text="A+", command=app.increase_font).pack(side="left", padx=2)
    ttk.Button(format_frame, text="A-", command=app.decrease_font).pack(side="left")
    app.font_size_var = tk.StringVar()
    app.font_entry = ttk.Entry(format_frame, textvariable=app.font_size_var, width=4, state="disabled")
    app.font_entry.pack(side="left", padx=5)
    app.font_entry.bind("<Return>", lambda e: app.set_font_size())
    ttk.Button(format_frame, text="Kolor", command=app.choose_text_color).pack(side="left", padx=2)
    ttk.Button(format_frame, text="Tło", command=app.choose_bg_color).pack(side="left", padx=2)
    app.transparent_var = tk.BooleanVar(value=False)
    app.bg_check = ttk.Checkbutton(
        format_frame,
        text="Przezroczyste",
        variable=app.transparent_var,
        command=app.toggle_bg_visible,
    )
    app.bg_check.pack(side="left", padx=2)
    app.bg_check.state(["disabled"])
    ttk.Button(format_frame, text="L", command=lambda: app.set_alignment("left")).pack(side="left", padx=2)
    ttk.Button(format_frame, text="C", command=lambda: app.set_alignment("center")).pack(side="left", padx=2)
    ttk.Button(format_frame, text="R", command=lambda: app.set_alignment("right")).pack(side="left", padx=2)
    ttk.Button(format_frame, text="Środek H", command=app.center_selected_horizontal).pack(side="left", padx=2)
    ttk.Button(format_frame, text="Środek V", command=app.center_selected_vertical).pack(side="left", padx=2)
    ttk.Label(format_frame, text="Warstwa:").pack(side="left", padx=(5, 0))
    app.layer_var = tk.StringVar()
    app.layer_entry = ttk.Entry(format_frame, textvariable=app.layer_var, width=4, state="disabled")
    app.layer_entry.pack(side="left", padx=2)
    app.layer_entry.bind("<Return>", lambda e: app.set_layer())
    app.canvas_container = tk.Frame(app, bg="#b0b0b0")
    app.canvas_container.pack(side="left", fill="both", expand=True, padx=5, pady=5)
    app.canvas_container.pack_propagate(False)
    app.canvas_container.bind("<Configure>", app.resize_canvas)
    app.canvas = tk.Canvas(
        app.canvas_container,
        bg="#b0b0b0",
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
    app.canvas.configure(scrollregion=(-app.margin, -app.margin, app.page_width + app.margin, app.page_height + app.margin))

    zoom_frame = ttk.Frame(app.canvas_container)
    zoom_frame.place(relx=1.0, rely=1.0, anchor="se", x=-5, y=-5)
    ttk.Button(zoom_frame, text="Dopasuj", command=app.fit_to_window).pack(side="right")
    app.zoom_var = tk.StringVar(value="100%")
    ttk.Label(zoom_frame, textvariable=app.zoom_var).pack(side="right", padx=5)

    right_container = ttk.Frame(app)
    right_container.pack(side="right", fill="y", padx=5, pady=5)
    app.right_canvas = tk.Canvas(right_container, width=300)
    right_scroll = ttk.Scrollbar(right_container, orient="vertical", command=app.right_canvas.yview)
    app.right_canvas.configure(yscrollcommand=right_scroll.set)
    right_scroll.pack(side="right", fill="y")
    app.right_canvas.pack(side="left", fill="both", expand=True)
    right_frame = ttk.Frame(app.right_canvas)
    app.right_canvas.create_window((0,0), window=right_frame, anchor="nw")
    right_frame.bind("<Configure>", lambda e: app.right_canvas.configure(scrollregion=app.right_canvas.bbox("all")))
    app.right_canvas.bind("<Enter>", lambda e: app.right_canvas.bind_all("<MouseWheel>", app._on_mousewheel))
    app.right_canvas.bind("<Leave>", lambda e: app.right_canvas.unbind_all("<MouseWheel>"))
    app.right_canvas.bind("<Button-4>", lambda e: app.right_canvas.yview_scroll(-1, "units"))
    app.right_canvas.bind("<Button-5>", lambda e: app.right_canvas.yview_scroll(1, "units"))

    # Dynamic column checkboxes
    ttk.Label(right_frame, text="Kolumny z Excela:").pack(anchor="w")
    app.columns_frame = ttk.Frame(right_frame)
    app.columns_frame.pack(fill="y", expand=True)
    app.columns_vars = {}

    # Static field checkboxes
    ttk.Label(right_frame, text="Pola statyczne:").pack(anchor="w", pady=(10, 0))
    app.static_frame = ttk.Frame(right_frame)
    app.static_frame.pack(fill="x")
    app.static_vars = {}
    app.static_entries = {}
    app.static_rows = {}
    for field in app.DEFAULT_STATIC_FIELDS:
        app.create_static_row(field, "")
    app.add_static_btn = ttk.Button(app.static_frame, text="Dodaj pole", command=app.add_static_field)
    app.add_static_btn.pack(fill="x", pady=5)

    # Row preview controls
    preview_frame = ttk.Frame(right_frame)
    preview_frame.pack(fill="x", pady=(10, 0))
    ttk.Label(preview_frame, text="Numer wiersza:").pack(side="left")
    app.row_var = tk.StringVar(value="1")
    ttk.Entry(preview_frame, textvariable=app.row_var, width=6).pack(side="left")
    ttk.Button(preview_frame, text="Podgląd", command=app.preview_row).pack(side="left", padx=5)

    # Group list
    ttk.Label(right_frame, text="Grupy:").pack(anchor="w", pady=(10, 0))
    grp_container = ttk.Frame(right_frame)
    grp_container.pack(fill="x")
    app.groups_list = tk.Listbox(grp_container, height=5)
    app.groups_list.pack(side="left", fill="both", expand=True)
    grp_scroll = ttk.Scrollbar(grp_container, orient="vertical", command=app.groups_list.yview)
    grp_scroll.pack(side="right", fill="y")
    app.groups_list.configure(yscrollcommand=grp_scroll.set)
    app.groups_list.bind("<Double-1>", lambda e: app.edit_selected_group())
    ttk.Button(right_frame, text="Usuń grupę", command=app.remove_group).pack(fill="x", pady=(5, 0))

    # Buttons
    button_frame = ttk.Frame(right_frame)
    button_frame.pack(fill="x", pady=(20, 0))
    ttk.Button(button_frame, text="Zapisz konfigurację", command=app.save_config).pack(fill="x")
    ttk.Button(button_frame, text="Warunki", command=app.open_conditions).pack(fill="x", pady=5)
    ttk.Button(button_frame, text="Dodaj grupę", command=app.add_group).pack(fill="x", pady=5)
    ttk.Button(button_frame, text="Generuj PDS", command=app.generate_pds).pack(fill="x", pady=5)

    # Progress bar
    app.progress = ttk.Progressbar(right_frame, orient="horizontal", mode="determinate")
    app.progress.pack(fill="x", pady=(20, 0))
    app.time_label = ttk.Label(right_frame, text="")
    app.time_label.pack()
    app.draw_grid()
    app.bind_all("<Delete>", app.delete_selected)
