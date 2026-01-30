import logging
import os
import time
import threading
import re
from concurrent.futures import ProcessPoolExecutor, as_completed
from contextlib import suppress
from io import BytesIO
from types import SimpleNamespace

import pandas as pd
import requests
from PIL import Image
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from tkinter import messagebox

from ..number_format import round_numeric_value

from .excel_tracking import (
    TRACKING_COLUMN,
    update_tracking_column,
    update_tracking_cache,
)

logger = logging.getLogger(__name__)


def _ui_call(app, func, *args, **kwargs):
    try:
        if hasattr(app, "ui_call"):
            app.ui_call(func, *args, **kwargs)
        else:
            func(*args, **kwargs)
    except Exception:
        logger.exception("Failed to update UI")


def _ui_status(app, text):
    if hasattr(app, "set_status"):
        _ui_call(app, app.set_status, text)


def _ui_counts(app, total_rows=None, total_tasks=None, processed=None, skipped=None):
    if hasattr(app, "set_counts"):
        _ui_call(
            app,
            app.set_counts,
            total_rows=total_rows,
            total_tasks=total_tasks,
            processed=processed,
            skipped=skipped,
        )


def _ui_progress_mode(app, indeterminate):
    if hasattr(app, "set_progress_indeterminate"):
        _ui_call(app, app.set_progress_indeterminate, indeterminate)


def _ui_finish(app, status):
    if hasattr(app, "finish_generation_ui"):
        _ui_call(app, app.finish_generation_ui, status)


def _is_cancelled(app):
    cancel_event = getattr(app, "cancel_event", None)
    return cancel_event is not None and cancel_event.is_set()


def to_reportlab_color(value):
    try:
        return colors.HexColor(value)
    except (ValueError, TypeError):
        return colors.toColor(value)


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r"[^\w\s-]", "", str(name))
    cleaned = re.sub(r"\s+", "_", cleaned).strip("_")
    return cleaned


def _first_data_column(df):
    for col in df.columns:
        if col != TRACKING_COLUMN:
            return col
    return None


def _get_filename_value(df, name_column, idx):
    if df is None or df.empty or not name_column:
        return ""
    if idx >= len(df):
        return ""
    row = df.iloc[idx]
    value = row.get(name_column, "")
    if pd.isna(value):
        return ""
    return value


def draw_pdf_element(app, c, element, value, x, y):
    if isinstance(value, str) and value.lower().startswith("http"):
        try:
            resp = requests.get(value, timeout=5)
            img = Image.open(BytesIO(resp.content))
            c.drawImage(
                ImageReader(img),
                x,
                y,
                width=element.width / app.scale,
                height=element.height / app.scale,
            )
            return
        except (requests.RequestException, OSError):
            logger.exception("Failed to load remote image %s", value)
    if isinstance(value, str):
        local_path = app.find_local_image(value)
        if local_path:
            try:
                img = Image.open(local_path)
                c.drawImage(
                    ImageReader(img),
                    x,
                    y,
                    width=element.width / app.scale,
                    height=element.height / app.scale,
                )
                return
            except OSError:
                logger.exception("Failed to load local image %s", local_path)
    if element.bg_visible:
        c.setFillColor(to_reportlab_color(element.bg_color))
        c.rect(
            x,
            y,
            element.width / app.scale,
            element.height / app.scale,
            fill=1,
            stroke=0,
        )
    c.setFillColor(to_reportlab_color(element.text_color))
    font_name = "Helvetica-Bold" if element.bold else "Helvetica"
    font_size = element.font_size / app.scale
    c.setFont(font_name, font_size)
    ascent = pdfmetrics.getAscent(font_name, font_size)
    descent = pdfmetrics.getDescent(font_name, font_size)
    text_y = y + (element.height / app.scale - (ascent - descent)) / 2 - descent
    if element.align == "center":
        c.drawCentredString(
            x + (element.width / app.scale) / 2,
            text_y,
            str(value),
        )
    elif element.align == "right":
        c.drawRightString(
            x + (element.width / app.scale),
            text_y,
            str(value),
        )
    else:
        c.drawString(x, text_y, str(value))


_render_context = None
_render_proxy = None


class RenderAppProxy:
    def __init__(self, scale, excel_dir):
        self.scale = scale
        self.excel_dir = excel_dir or ""
        self._image_cache = {}

    def find_local_image(self, filename):
        if not filename or not self.excel_dir:
            return None
        key = str(filename).lower()
        if key in self._image_cache:
            return self._image_cache[key]
        if os.path.isabs(filename):
            path = filename if os.path.exists(filename) else None
        else:
            candidate = os.path.join(self.excel_dir, str(filename))
            if os.path.exists(candidate):
                path = candidate
            else:
                path = None
                for root, _dirs, files in os.walk(self.excel_dir):
                    for current in files:
                        if current.lower() == key:
                            path = os.path.join(root, current)
                            break
                    if path:
                        break
        self._image_cache[key] = path
        return path


def _init_render_context(context):
    global _render_context, _render_proxy
    _render_context = context
    _render_proxy = RenderAppProxy(context["scale"], context["excel_dir"])


def _collect_element_specs(app):
    elements = {}
    order = []
    for name, element in sorted(app.elements.items(), key=lambda kv: kv[1].layer):
        elements[name] = {
            "x": element.x,
            "y": element.y,
            "width": element.width,
            "height": element.height,
            "font_size": element.font_size,
            "bold": element.bold,
            "text_color": element.text_color,
            "bg_color": element.bg_color,
            "bg_visible": element.bg_visible,
            "align": element.align,
            "auto_font": getattr(element, "auto_font", True),
            "layer": element.layer,
        }
        order.append(name)
    return elements, order


def _collect_group_specs(app):
    groups = []
    for group in app.groups.values():
        groups.append(
            {
                "x": group.x,
                "y": group.y,
                "width": group.width,
                "height": group.height,
                "fields": list(group.fields),
                "field_pos": {k: (pos[0], pos[1]) for k, pos in group.field_pos.items()},
                "field_conf": {k: dict(conf) for k, conf in group.field_conf.items()},
                "conditions": [tuple(cond) for cond in group.conditions],
            }
        )
    return groups


def _collect_static_entries(app):
    return {name: var.get() for name, var in getattr(app, "static_entries", {}).items()}


def render_single_pdf(task):
    if _render_context is None or _render_proxy is None:
        raise RuntimeError("Render context not initialised")

    context = _render_context
    idx = task["idx"]
    target_path = task["pdf_path"]
    tmp_path = f"{target_path}.{os.getpid()}.tmp"

    scale = context["scale"]
    page_width = context["page_width"]
    page_height = context["page_height"]
    element_specs = context["elements"]
    element_order = context["element_order"]
    groups = context["groups"]
    conditions = context["conditions"]
    static_entries = context["static_entries"]
    elements = {name: SimpleNamespace(**spec) for name, spec in element_specs.items()}

    needed = set(elements.keys())
    for group in groups:
        needed.update(group["fields"])
    needed.update(static_entries.keys())

    row_values = task.get("row_values", {})
    values = {}
    for name in needed:
        if ":" in name:
            value = row_values.get(name, "")
        else:
            value = static_entries.get(name, "")
        if pd.isna(value):
            value = ""
        values[name] = value

    group_field_names = {fname for group in groups for fname in group["fields"]}

    hidden = set()
    for src, tgt in conditions:
        if src in group_field_names or tgt in group_field_names:
            continue
        if pd.isna(values.get(src)) or values.get(src) == "":
            hidden.add(tgt)

    c = pdf_canvas.Canvas(tmp_path, pagesize=(page_width, page_height))
    success = False
    try:
        for group in groups:
            g_hidden = set()
            for src, tgt in group["conditions"]:
                if src not in group["fields"] or tgt not in group["fields"]:
                    continue
                if pd.isna(values.get(src)) or values.get(src) == "":
                    g_hidden.add(tgt)
            positions = group["field_pos"]
            columns = {}
            for fname in group["fields"]:
                if fname in hidden or fname in g_hidden:
                    continue
                val = values.get(fname, "")
                if val == "":
                    continue
                conf = group["field_conf"].get(fname, {})
                el = elements.get(fname)
                if not conf and not el:
                    continue
                width = conf.get("width", el.width / scale if el else 0)
                height = conf.get("height", el.height / scale if el else 0)
                x0, y0 = positions.get(fname, (0, 0))
                columns.setdefault(x0, []).append((y0, fname, width, height, conf, el, val))

            placed = []
            for x0 in sorted(columns):
                col_items = columns[x0]
                col_items.sort(key=lambda t: t[0])
                cur_y = 0
                for _, fname, width, height, conf, el, val in col_items:
                    y = cur_y
                    while True:
                        overlap = False
                        for px, py, pw, ph in placed:
                            if (
                                x0 < px + pw
                                and x0 + width > px
                                and y < py + ph
                                and y + height > py
                            ):
                                y = py + ph
                                overlap = True
                                break
                        if not overlap:
                            break
                    if y + height > group["height"] / scale:
                        continue
                    font_size = conf.get(
                        "font_size", el.font_size / scale if el else 12
                    )
                    dummy = SimpleNamespace(
                        width=width * scale,
                        height=height * scale,
                        font_size=font_size * scale,
                        bold=conf.get("bold", el.bold if el else False),
                        text_color=conf.get(
                            "text_color", el.text_color if el else "black"
                        ),
                        bg_color=conf.get("bg_color", el.bg_color if el else "white"),
                        bg_visible=conf.get(
                            "bg_visible", el.bg_visible if el else True
                        ),
                        align=conf.get("align", el.align if el else "left"),
                        auto_font=conf.get("auto_font", el.auto_font if el else True),
                    )
                    x_pdf = group["x"] / scale + x0
                    y_pdf = page_height - (group["y"] / scale + y + height)
                    draw_pdf_element(_render_proxy, c, dummy, val, x_pdf, y_pdf)
                    placed.append((x0, y, width, height))
                    cur_y = y + height

        for name in element_order:
            element = elements.get(name)
            if not element or name in hidden:
                continue
            val = values.get(name, "")
            x = element.x / scale
            y = page_height - (element.y / scale) - (element.height / scale)
            draw_pdf_element(_render_proxy, c, element, val, x, y)

        c.showPage()
        c.save()
        final_path = target_path
        try:
            os.replace(tmp_path, final_path)
        except OSError:
            logger.exception("Failed to replace %s, trying alternative name", final_path)
            base, ext = os.path.splitext(final_path)
            alt_path = f"{base}_{idx + 1}_{int(time.time())}{ext}"
            try:
                os.replace(tmp_path, alt_path)
            except OSError:
                logger.exception("Failed to move temp PDF to %s", alt_path)
                raise
            else:
                final_path = alt_path
        success = True
        return {"idx": idx, "pdf_path": final_path, "name": task.get("name")}
    finally:
        if not success:
            with suppress(OSError):
                os.remove(tmp_path)


def generate_pds(app):
    def finish_now(status):
        if hasattr(app, "finish_generation_ui"):
            app.finish_generation_ui(status)
        else:
            _ui_finish(app, status)

    if not app.excel_path or not app.dataframes:
        messagebox.showerror("Błąd", "Brak danych do generowania")
        finish_now("Brak danych")
        return False

    _, first_df = next(iter(app.dataframes.items()))
    total_rows = len(first_df)
    if total_rows == 0:
        messagebox.showinfo("Info", "Brak wierszy w pliku Excel")
        finish_now("Brak wierszy")
        return False

    _ui_counts(app, total_rows=total_rows)

    element_specs, element_order = _collect_element_specs(app)
    group_specs = _collect_group_specs(app)
    conditions = [tuple(cond) for cond in app.conditions]

    dynamic_fields = set()

    def collect_dynamic(name):
        if isinstance(name, str) and ":" in name:
            dynamic_fields.add(name)

    for name in element_specs.keys():
        collect_dynamic(name)
    for group in group_specs:
        for fname in group["fields"]:
            collect_dynamic(fname)
        for src, tgt in group["conditions"]:
            collect_dynamic(src)
            collect_dynamic(tgt)
    for src, tgt in conditions:
        collect_dynamic(src)
        collect_dynamic(tgt)

    sheet_fields = {}
    excluded_fields = getattr(app, "tracking_excluded", set())
    for field in dynamic_fields:
        if field in excluded_fields:
            continue
        sheet, col = field.split(":", 1)
        if col == TRACKING_COLUMN:
            continue
        sheet_fields.setdefault(sheet, set()).add(col)

    cache_rows = None
    cache_changed_cols = {}
    excel_rows = None
    tracking_mode = "none"

    try:
        cache_rows, cache_changed_cols = update_tracking_cache(
            app.excel_path, app.dataframes, total_rows, sheet_fields=sheet_fields
        )
    except Exception:
        logger.exception("Failed to update tracking cache")
        cache_rows = None
        cache_changed_cols = {}

    try:
        excel_rows = update_tracking_column(
            app.excel_path, app.dataframes, total_rows, sheet_fields=sheet_fields
        )
        tracking_mode = "excel"
    except Exception:
        logger.exception("Failed to update tracking column")
        if cache_rows is None:
            messagebox.showwarning(
                "Uwaga",
                "Nie udało się zaktualizować kolumny kontrolnej w Excelu "
                "(sprawdź, czy plik nie jest otwarty). "
                "Pliki PDF zostaną wygenerowane ponownie.",
            )
        else:
            tracking_mode = "cache"
            messagebox.showinfo(
                "Info",
                "Nie udało się zaktualizować kolumny kontrolnej w Excelu "
                "(sprawdź, czy plik nie jest otwarty). "
                "Używam lokalnego pliku śledzenia, więc wygenerują się "
                "tylko zmienione PDF-y.",
            )

    if excel_rows is None:
        changed_rows = cache_rows
        tracking_mode = "cache" if cache_rows is not None else tracking_mode
    else:
        changed_rows = excel_rows
        if cache_rows is not None and len(excel_rows) == total_rows and len(cache_rows) < total_rows:
            logger.info(
                "Excel tracking indicates all rows changed; using cache (%s/%s).",
                len(cache_rows),
                total_rows,
            )
            changed_rows = cache_rows
            tracking_mode = "cache"

    tracked_columns = sum(len(cols) for cols in sheet_fields.values())
    cache_count = len(cache_rows) if cache_rows is not None else None
    excel_count = len(excel_rows) if excel_rows is not None else None
    if changed_rows is None:
        logger.info(
            "Tracking mode=%s; tracked_columns=%s; cache_rows=%s; excel_rows=%s; generating all rows=%s",
            tracking_mode,
            tracked_columns,
            cache_count,
            excel_count,
            total_rows,
        )
    else:
        logger.info(
            "Tracking mode=%s; tracked_columns=%s; cache_rows=%s; excel_rows=%s; changed_rows=%s/%s",
            tracking_mode,
            tracked_columns,
            cache_count,
            excel_count,
            len(changed_rows),
            total_rows,
        )
        if len(changed_rows) >= max(1, total_rows - 1) and cache_changed_cols:
            changed_preview = []
            for sheet, cols in cache_changed_cols.items():
                for col in cols:
                    changed_preview.append(f"{sheet}:{col}")
                    if len(changed_preview) >= 6:
                        break
                if len(changed_preview) >= 6:
                    break
            logger.info(
                "Most rows changed; changed columns (sample): %s",
                ", ".join(changed_preview),
            )

    output_dir = os.path.join(os.path.dirname(app.excel_path), "PDS")
    os.makedirs(output_dir, exist_ok=True)

    page_width = app.page_width
    page_height = app.page_height

    static_entries = _collect_static_entries(app)

    sheet_columns = {
        sheet: {col for col in df.columns if col != TRACKING_COLUMN}
        for sheet, df in app.dataframes.items()
    }

    name_column = _first_data_column(first_df)
    filename_counters = {}
    tasks = []
    for idx in range(total_rows):
        first_val = _get_filename_value(first_df, name_column, idx)
        filename = sanitize_filename(first_val) or f"pds_{idx + 1}"
        count = filename_counters.get(filename, 0)
        filename_counters[filename] = count + 1
        if count:
            unique_name = f"{filename}_{count + 1}"
        else:
            unique_name = filename
        pdf_path = os.path.join(output_dir, f"{unique_name}.pdf")
        if os.path.exists(pdf_path) and changed_rows is not None and idx not in changed_rows:
            continue
        row_values = {}
        for sheet, columns in sheet_fields.items():
            df = app.dataframes.get(sheet)
            if df is None or idx >= len(df):
                for col in columns:
                    row_values[f"{sheet}:{col}"] = ""
                continue
            row_series = df.iloc[idx]
            available = sheet_columns.get(sheet, set())
            for col in columns:
                if col in available:
                    value = row_series[col]
                else:
                    value = ""
                if pd.isna(value):
                    value = ""
                else:
                    value = round_numeric_value(value)
                row_values[f"{sheet}:{col}"] = value
        tasks.append(
            {
                "idx": idx,
                "pdf_path": pdf_path,
                "name": unique_name,
                "row_values": row_values,
            }
        )

    if not tasks:
        messagebox.showinfo(
            "Info", "Brak zmian - wszystkie pliki PDF są aktualne."
        )
        _ui_counts(app, total_rows=total_rows, total_tasks=0, processed=0, skipped=total_rows)
        finish_now("Brak zmian")
        return False

    skipped_rows = total_rows - len(tasks) if changed_rows is not None else 0
    _ui_counts(
        app,
        total_rows=total_rows,
        total_tasks=len(tasks),
        processed=0,
        skipped=skipped_rows,
    )
    _ui_progress_mode(app, False)
    _ui_status(app, "Generowanie PDF...")

    if _is_cancelled(app):
        finish_now("Anulowano")
        return False

    worker_payload = {
        "static_entries": static_entries,
        "elements": element_specs,
        "element_order": element_order,
        "groups": group_specs,
        "conditions": conditions,
        "scale": app.scale,
        "page_width": page_width,
        "page_height": page_height,
        "excel_dir": os.path.dirname(app.excel_path),
        "tasks": tasks,
        "output_dir": output_dir,
        "total_tasks": len(tasks),
        "total_rows": total_rows,
        "skipped_rows": skipped_rows,
    }

    def worker(payload):
        start_time = time.time()
        context = {
            "scale": payload["scale"],
            "page_width": payload["page_width"],
            "page_height": payload["page_height"],
            "elements": payload["elements"],
            "element_order": payload["element_order"],
            "groups": payload["groups"],
            "conditions": payload["conditions"],
            "static_entries": payload["static_entries"],
            "excel_dir": payload["excel_dir"],
        }

        tasks_local = payload["tasks"]
        total = payload["total_tasks"]
        failures = []
        completed = 0
        cancelled = False
        max_workers = max(1, min(len(tasks_local), os.cpu_count() or 1))

        executor = None
        future_map = {}
        try:
            executor = ProcessPoolExecutor(
                max_workers=max_workers,
                initializer=_init_render_context,
                initargs=(context,),
            )
            future_map = {
                executor.submit(render_single_pdf, task): task for task in tasks_local
            }
            for future in as_completed(future_map):
                if _is_cancelled(app):
                    cancelled = True
                    break
                task = future_map[future]
                try:
                    future.result()
                except Exception as exc:  # pragma: no cover - defensive logging
                    failures.append((task["idx"], task.get("name", ""), str(exc)))
                    logger.exception(
                        "Failed to render PDF for row %s", task["idx"] + 1, exc_info=exc
                    )
                completed += 1
                progress = completed / total * 100
                elapsed = time.time() - start_time
                remaining = (elapsed / completed) * (total - completed) if completed else 0
                remaining_seconds = max(0, int(remaining))
                app.progress.after(
                    0, lambda p=progress: app.progress.config(value=p)
                )
                _ui_counts(
                    app,
                    total_tasks=total,
                    processed=completed,
                )
                app.time_label.after(
                    0,
                    lambda r=remaining_seconds: app.time_label.config(
                        text=f"Pozostały czas: {r} s"
                    ),
                )
        except Exception as exc:  # pragma: no cover - executor level failure
            failures.append((-1, "", str(exc)))
            logger.exception("PDF generation failed", exc_info=exc)
        finally:
            if executor is not None:
                if cancelled:
                    for fut in future_map:
                        fut.cancel()
                    try:
                        executor.shutdown(wait=False, cancel_futures=True)
                    except TypeError:
                        try:
                            executor.shutdown(wait=False)
                        except TypeError:
                            executor.shutdown()
                else:
                    executor.shutdown()
            def finish():
                if cancelled:
                    if hasattr(app, "finish_generation_ui"):
                        app.finish_generation_ui("Anulowano")
                    else:
                        _ui_finish(app, "Anulowano")
                    return
                if failures:
                    if hasattr(app, "finish_generation_ui"):
                        app.finish_generation_ui("Błąd")
                    else:
                        _ui_finish(app, "Błąd")
                    failed_rows = [idx for idx, _name, _err in failures if idx >= 0]
                    if failed_rows:
                        rows_text = ", ".join(str(idx + 1) for idx in failed_rows)
                        message = (
                            "Wystąpiły błędy podczas generowania wierszy: "
                            f"{rows_text}. Sprawdź logi."
                        )
                    else:
                        message = "Wystąpił błąd podczas generowania plików PDF. Sprawdź logi."
                    messagebox.showerror("Błąd", message)
                else:
                    if hasattr(app, "finish_generation_ui"):
                        app.finish_generation_ui("Zakończono")
                    else:
                        _ui_finish(app, "Zakończono")
                    messagebox.showinfo(
                        "Zakończono", f"Pliki zapisane w {payload['output_dir']}"
                    )

            app.after(0, finish)

    threading.Thread(target=worker, args=(worker_payload,), daemon=True).start()
    return True
