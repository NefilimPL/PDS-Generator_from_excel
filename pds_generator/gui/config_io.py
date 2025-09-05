import json
import os
import logging
import shutil
from tkinter import messagebox

from ..elements import DraggableElement
from ..groups import GroupArea

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pds_generator")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
OLD_CONFIG_FILE = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "config.json")
)

logger = logging.getLogger(__name__)


def _ensure_config_dir():
    os.makedirs(CONFIG_DIR, exist_ok=True)


def _excel_config_path(excel_path):
    if not excel_path:
        return None
    return os.path.join(os.path.dirname(excel_path), "config.json")


def _acquire_lock(path):
    lock = f"{path}.lock"
    if os.path.exists(lock):
        messagebox.showerror("Błąd", f"Plik {os.path.basename(path)} jest używany na innym komputerze")
        return None
    try:
        with open(lock, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
    except OSError:
        logger.exception("Failed to create lock %s", lock)
        messagebox.showerror("Błąd", f"Nie można utworzyć blokady dla {path}")
        return None
    return lock


def _release_lock(path):
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            logger.exception("Failed to remove lock %s", path)


def save_config(app):
    if not app.excel_path:
        messagebox.showerror("Błąd", "Najpierw wybierz plik Excel")
        return
    config = {
        "excel_path": app.excel_path,
        "page_width": app.page_width,
        "page_height": app.page_height,
        "elements": [el.to_dict() for el in app.elements.values()],
        "static_fields": {name: var.get() for name, var in app.static_entries.items()},
        "conditions": app.conditions,
        "groups": [g.to_dict() for g in app.groups.values()],
        "ignore_updates": getattr(app, "ignore_updates", False),
        "update_test": getattr(app, "update_test", False),
    }
    cfg_path = _excel_config_path(app.excel_path)
    if not cfg_path:
        messagebox.showerror("Błąd", "Brak ścieżki do konfiguracji")
        return
    lock = _acquire_lock(cfg_path)
    if not lock:
        return
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except OSError:
        logger.exception("Failed to save config to %s", cfg_path)
        messagebox.showerror("Błąd", f"Nie można zapisać konfiguracji do {cfg_path}")
    finally:
        _release_lock(lock)
    _ensure_config_dir()
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except OSError:
        logger.exception("Failed to save backup config to %s", CONFIG_FILE)
    messagebox.showinfo("Zapisano", f"Zapisano konfigurację do {cfg_path}")


def load_config(app, startup=False, path=None):
    excel_path = path or app.excel_path
    if not excel_path and startup and os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                excel_path = json.load(f).get("excel_path")
        except Exception:
            excel_path = None

    cfg_path = None
    loaded_from_backup = False
    if excel_path:
        candidate = _excel_config_path(excel_path)
        if candidate and os.path.exists(candidate):
            cfg_path = candidate
        else:
            loaded_from_backup = True
    if cfg_path is None:
        cfg_path = CONFIG_FILE if os.path.exists(CONFIG_FILE) else None
    if cfg_path is None and os.path.exists(OLD_CONFIG_FILE):
        cfg_path = OLD_CONFIG_FILE
    if cfg_path is None or not os.path.exists(cfg_path):
        return

    if getattr(app, "config_lock_path", None):
        _release_lock(app.config_lock_path)
        app.config_lock_path = None
    lock = _acquire_lock(cfg_path)
    if not lock:
        return
    app.config_lock_path = lock

    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except (OSError, json.JSONDecodeError):
        logger.exception("Failed to load config from %s", cfg_path)
        _release_lock(lock)
        app.config_lock_path = None
        return

    if cfg_path == OLD_CONFIG_FILE:
        try:
            _ensure_config_dir()
            shutil.move(OLD_CONFIG_FILE, CONFIG_FILE)
            cfg_path = CONFIG_FILE
        except OSError:
            logger.exception("Failed to migrate config to %s", CONFIG_FILE)
    if cfg_path != CONFIG_FILE:
        _ensure_config_dir()
        try:
            shutil.copy(cfg_path, CONFIG_FILE)
        except OSError:
            logger.exception("Failed to update backup config from %s", cfg_path)

    app.ignore_updates = config.get("ignore_updates", False)
    app.update_test = config.get("update_test", False)
    excel_cfg = config.get("excel_path")
    if startup and excel_cfg and os.path.exists(excel_cfg):
        if not getattr(app, "excel_lock_path", None):
            if not app.acquire_excel_lock(excel_cfg):
                return
        app.excel_path = excel_cfg
        app.image_cache = {}
        app.path_var.set(excel_cfg)
        app.load_excel(excel_cfg)
    if path and excel_cfg and excel_cfg != path:
        return
    if loaded_from_backup:
        messagebox.showwarning(
            "Kopia zapasowa",
            "Otworzono konfigurację z kopii zapasowej. Prosimy o zapisanie konfiguracji.",
        )

    app.page_width = config.get("page_width", app.page_width)
    app.page_height = config.get("page_height", app.page_height)
    set_name = None
    for n, sz in app.PAGE_SIZES.items():
        if abs(sz[0] - app.page_width) < 1 and abs(sz[1] - app.page_height) < 1:
            set_name = n
            break
    if set_name:
        app.size_var.set(set_name)
    else:
        app.size_var.set(f"{int(app.page_width)}x{int(app.page_height)}")
    app.resize_canvas()
    for name, val in config.get("static_fields", {}).items():
        if name not in app.static_vars:
            app.create_static_row(name, val)
        else:
            app.static_entries[name].set(val)
    app.conditions = config.get("conditions", [])
    for elconf in config.get("elements", []):
        name = elconf["name"]
        if name not in app.elements:
            element = DraggableElement(app, app.canvas, name, elconf.get("text", name))
            element.x = elconf.get("x", element.x) * app.scale
            element.y = elconf.get("y", element.y) * app.scale
            element.width = elconf.get("width", element.width) * app.scale
            element.height = elconf.get("height", element.height) * app.scale
            element.font_size = elconf.get("font_size", element.font_size) * app.scale
            element.bold = elconf.get("bold", element.bold)
            element.text_color = elconf.get("text_color", element.text_color)
            element.bg_color = elconf.get("bg_color", element.bg_color)
            element.bg_visible = elconf.get("bg_visible", element.bg_visible)
            element.align = elconf.get("align", element.align)
            element.auto_font = elconf.get("auto_font", element.auto_font)
            element.layer = elconf.get("layer", element.layer)
            element.sync_canvas()
            app.elements[name] = element
            if name in app.columns_vars:
                app.columns_vars[name].set(True)
            if name in app.static_vars:
                app.static_vars[name].set(True)
                app.static_entries[name].set(elconf.get("text", ""))
    for gconf in config.get("groups", []):
        group = GroupArea(app, app.canvas, gconf.get("name", f"Group{len(app.groups)+1}"))
        group.x = gconf.get("x", group.x) * app.scale
        group.y = gconf.get("y", group.y) * app.scale
        group.width = gconf.get("width", group.width) * app.scale
        group.height = gconf.get("height", group.height) * app.scale
        group.sync_canvas()
        group.field_pos = {k: (v[0], v[1]) for k, v in gconf.get("field_pos", {}).items()}
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
        app.groups[group.name] = group
        if hasattr(app, "groups_list"):
            app.groups_list.insert("end", group.name)
    app.restack_elements()
    app.push_history()
