import json
import os
import logging
import shutil
from tkinter import messagebox

from .. import app_paths
from ..elements import DraggableElement
from ..groups import GroupArea
from ..layout_dependencies import normalize_dependencies
from ..pdf_settings import DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT
from ..value_sources import (
    FILE_DATE_KIND_MODIFIED,
    VALUE_SOURCE_DEFAULT,
    normalize_file_date_kind,
    normalize_value_source,
)
from . import locks
from . import mailer

CONFIG_DIR = app_paths.get_app_storage_dir()
CONFIG_FILE = app_paths.get_backup_config_path()
LEGACY_CONFIG_FILE = app_paths.get_legacy_backup_config_path()
OLD_CONFIG_FILE = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "config.json")
)

logger = logging.getLogger(__name__)


def _ensure_config_dir():
    os.makedirs(CONFIG_DIR, exist_ok=True)


def _candidate_backup_files():
    candidates = []
    for candidate in (CONFIG_FILE, LEGACY_CONFIG_FILE):
        if candidate and candidate not in candidates:
            candidates.append(candidate)
    return candidates


def _excel_config_path(excel_path):
    if not excel_path:
        return None
    return os.path.join(os.path.dirname(excel_path), "config.json")


def _acquire_lock(path):
    return locks.acquire_lock(path, os.path.basename(path))


def _release_lock(path):
    try:
        locks.release_lock(path)
    except Exception:
        logger.exception("Failed to remove lock %s", path)


def _normalize_path(path):
    if not path:
        return ""
    return os.path.normcase(os.path.abspath(os.path.normpath(str(path))))


def _same_path(path_a, path_b):
    if not path_a or not path_b:
        return False
    return _normalize_path(path_a) == _normalize_path(path_b)


def _apply_element_config(app, element, elconf):
    scale = app.scale if app.scale else 1.0
    default_text = getattr(element, "text", element.name)
    element.text = str(elconf.get("text", default_text) or "")
    element.x = elconf.get("x", element.x / scale) * scale
    element.y = elconf.get("y", element.y / scale) * scale
    element.width = elconf.get("width", element.width / scale) * scale
    element.height = elconf.get("height", element.height / scale) * scale
    element.font_size = elconf.get("font_size", element.font_size / scale) * scale
    element.max_font_size = (
        elconf.get(
            "max_font_size",
            elconf.get("font_size", element.font_size / scale),
        )
        * scale
    )
    element.bold = elconf.get("bold", element.bold)
    element.text_color = elconf.get("text_color", element.text_color)
    element.bg_color = elconf.get("bg_color", element.bg_color)
    element.bg_visible = elconf.get("bg_visible", element.bg_visible)
    element.align = elconf.get("align", element.align)
    element.auto_font = elconf.get("auto_font", element.auto_font)
    element.layer = elconf.get("layer", element.layer)
    element.value_source = normalize_value_source(
        elconf.get("value_source", getattr(element, "value_source", VALUE_SOURCE_DEFAULT))
    )
    element.file_date_kind = normalize_file_date_kind(
        elconf.get(
            "file_date_kind",
            getattr(element, "file_date_kind", FILE_DATE_KIND_MODIFIED),
        )
    )
    if elconf.get("is_image"):
        app.image_fields.add(element.name)
    element.is_image = element.name in app.image_fields
    element.sync_canvas()


def _clear_groups(app):
    for group in list(app.groups.values()):
        for item in (group.rect, group.handle) + tuple(
            getattr(group, "preview_items", [])
        ):
            app.canvas.delete(item)
    app.groups = {}
    if hasattr(app, "groups_list"):
        app.groups_list.delete(0, "end")


def save_config(app):
    if not app.excel_path:
        messagebox.showerror("Błąd", "Najpierw wybierz plik Excel")
        return
    try:
        mail_cfg = mailer.sanitize_mail_config_for_storage(
            getattr(app, "mail_config", {})
        )
    except Exception as exc:
        logger.exception("Failed to sanitize mail config for storage")
        messagebox.showerror(
            "Błąd",
            f"Nie można zapisać konfiguracji e-mail: {exc}",
        )
        return
    config = {
        "excel_path": app.excel_path,
        "page_width": app.page_width,
        "page_height": app.page_height,
        "pdf_image_compression_percent": getattr(
            app,
            "pdf_image_compression_percent",
            DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
        ),
        "elements": [el.to_dict() for el in app.elements.values()],
        "static_fields": {name: var.get() for name, var in app.static_entries.items()},
        "image_fields": sorted(getattr(app, "image_fields", set())),
        "image_dirs": list(getattr(app, "image_dirs", [])),
        "conditions": app.conditions,
        "dependencies": normalize_dependencies(getattr(app, "dependencies", [])),
        "groups": [g.to_dict() for g in app.groups.values()],
        "tracking_excluded": sorted(getattr(app, "tracking_excluded", set())),
        "ignore_updates": getattr(app, "ignore_updates", False),
        "update_test": getattr(app, "update_test", False),
        "mail": mail_cfg,
    }
    cfg_path = _excel_config_path(app.excel_path)
    if not cfg_path:
        messagebox.showerror("Błąd", "Brak ścieżki do konfiguracji")
        return
    expected_lock = f"{cfg_path}.lock"
    current_lock = getattr(app, "config_lock_path", None)
    if (
        not current_lock
        or current_lock != expected_lock
        or not os.path.exists(current_lock)
    ):
        if current_lock and current_lock != expected_lock:
            _release_lock(current_lock)
        lock = _acquire_lock(cfg_path)
        if not lock:
            return
        app.config_lock_path = lock
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except OSError:
        logger.exception("Failed to save config to %s", cfg_path)
        messagebox.showerror("Błąd", f"Nie można zapisać konfiguracji do {cfg_path}")
    _ensure_config_dir()
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except OSError:
        logger.exception("Failed to save backup config to %s", CONFIG_FILE)
    messagebox.showinfo("Zapisano", f"Zapisano konfigurację do {cfg_path}")


def load_config(app, startup=False, path=None):
    excel_path = path or app.excel_path
    if not excel_path and startup:
        for backup_path in _candidate_backup_files():
            if not os.path.exists(backup_path):
                continue
            try:
                with open(backup_path, "r", encoding="utf-8") as f:
                    excel_path = json.load(f).get("excel_path")
            except Exception:
                excel_path = None
            if excel_path:
                break

    cfg_path = None
    loaded_from_backup = False
    if excel_path:
        candidate = _excel_config_path(excel_path)
        if candidate and os.path.exists(candidate):
            cfg_path = candidate
        else:
            loaded_from_backup = True
    if cfg_path is None:
        for backup_path in _candidate_backup_files():
            if os.path.exists(backup_path):
                cfg_path = backup_path
                break
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
    if cfg_path == LEGACY_CONFIG_FILE and LEGACY_CONFIG_FILE != CONFIG_FILE:
        try:
            _ensure_config_dir()
            shutil.move(LEGACY_CONFIG_FILE, CONFIG_FILE)
            cfg_path = CONFIG_FILE
        except OSError:
            logger.exception("Failed to migrate legacy backup config to %s", CONFIG_FILE)
    if cfg_path != CONFIG_FILE:
        _ensure_config_dir()
        try:
            shutil.copy(cfg_path, CONFIG_FILE)
        except OSError:
            logger.exception("Failed to update backup config from %s", cfg_path)

    excel_cfg = config.get("excel_path")
    selected_cfg_path = _excel_config_path(path) if path else None
    loaded_selected_excel_config = bool(
        selected_cfg_path and _same_path(cfg_path, selected_cfg_path)
    )
    if (
        path
        and excel_cfg
        and not _same_path(excel_cfg, path)
        and not loaded_selected_excel_config
    ):
        logger.info(
            "Skipping config load due to excel path mismatch: selected=%s config=%s",
            path,
            excel_cfg,
        )
        _release_lock(lock)
        app.config_lock_path = None
        if loaded_from_backup:
            messagebox.showwarning(
                "Kopia zapasowa",
                "Nie znaleziono config.json dla wybranego pliku Excel.\n"
                "Dostępna kopia zapasowa dotyczy innego pliku, więc nie została załadowana.",
            )
        return
    if path and excel_cfg and not _same_path(excel_cfg, path):
        logger.info(
            "Loading config from selected Excel directory despite excel_path mismatch: "
            "selected=%s config=%s",
            path,
            excel_cfg,
        )

    app.ignore_updates = config.get("ignore_updates", False)
    app.update_test = config.get("update_test", False)
    app.tracking_excluded = set(config.get("tracking_excluded", []))
    app.mail_config = mailer.load_mail_config(config.get("mail", {}))
    app.image_fields = set(config.get("image_fields", []))
    app.image_dirs = list(config.get("image_dirs", []))
    if hasattr(app, "set_pdf_image_compression_percent"):
        app.set_pdf_image_compression_percent(
            config.get(
                "pdf_image_compression_percent",
                getattr(
                    app,
                    "pdf_image_compression_percent",
                    DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
                ),
            )
        )
    else:
        app.pdf_image_compression_percent = config.get(
            "pdf_image_compression_percent",
            getattr(
                app,
                "pdf_image_compression_percent",
                DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
            ),
        )
    app.image_cache = {}
    if hasattr(app, "_invalidate_image_index"):
        app._invalidate_image_index()
    if hasattr(app, "refresh_image_dir_list"):
        app.refresh_image_dir_list()
    if startup and excel_cfg and os.path.exists(excel_cfg):
        if not getattr(app, "excel_lock_path", None):
            if not app.acquire_excel_lock(excel_cfg):
                _release_lock(lock)
                app.config_lock_path = None
                return
        app.excel_path = excel_cfg
        app.image_cache = {}
        app.path_var.set(excel_cfg)
        app.load_excel(excel_cfg)
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
    app.dependencies = normalize_dependencies(config.get("dependencies", []))
    for var in app.columns_vars.values():
        var.set(False)
    for var in app.static_vars.values():
        var.set(False)

    element_configs = []
    target_element_names = set()
    for raw_conf in config.get("elements", []):
        if not isinstance(raw_conf, dict):
            continue
        name = str(raw_conf.get("name") or "").strip()
        if not name:
            continue
        elconf = dict(raw_conf)
        elconf["name"] = name
        element_configs.append(elconf)
        target_element_names.add(name)

    if hasattr(app, "_suspend_dependency_prune"):
        app._suspend_dependency_prune = True
    try:
        for name in list(app.elements.keys()):
            if name not in target_element_names:
                app.remove_element(name)

        for elconf in element_configs:
            name = elconf["name"]
            if name not in app.elements:
                app.elements[name] = DraggableElement(
                    app,
                    app.canvas,
                    name,
                    elconf.get("text", name),
                )
            element = app.elements[name]
            _apply_element_config(app, element, elconf)
            if name in app.columns_vars:
                app.columns_vars[name].set(True)
            if name in app.static_vars:
                app.static_vars[name].set(True)
                if normalize_value_source(
                    getattr(element, "value_source", VALUE_SOURCE_DEFAULT)
                ) == VALUE_SOURCE_DEFAULT:
                    app.static_entries[name].set(element.text)
    finally:
        if hasattr(app, "_suspend_dependency_prune"):
            app._suspend_dependency_prune = False

    _clear_groups(app)
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
                "max_font_size": fc.get("max_font_size", fc.get("font_size", 12)),
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
    if hasattr(app, "_prune_dependencies"):
        app._prune_dependencies()
    app.restack_elements()
    if hasattr(app, "apply_image_field_state"):
        app.apply_image_field_state()
    if hasattr(app, "refresh_value_source_elements"):
        app.refresh_value_source_elements()
    app.push_history()
