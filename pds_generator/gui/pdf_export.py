import logging
import os
import time
import threading
import re
import math
import hashlib
from concurrent.futures import ProcessPoolExecutor, as_completed
from contextlib import suppress
from io import BytesIO
from types import SimpleNamespace

import pandas as pd
import requests
from PIL import Image
from openpyxl import load_workbook


def _patch_md5_usedforsecurity_compat():
    """Compatibility shim for Python builds without hashlib `usedforsecurity` kwarg."""
    md5_func = getattr(hashlib, "md5", None)
    if md5_func is None:
        return
    try:
        md5_func(b"", usedforsecurity=False)
        return
    except TypeError as exc:
        if "usedforsecurity" not in str(exc):
            return
    except Exception:
        return

    def _md5_compat(data=b"", *args, **kwargs):
        kwargs.pop("usedforsecurity", None)
        return md5_func(data, *args, **kwargs)

    hashlib.md5 = _md5_compat


_patch_md5_usedforsecurity_compat()

from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from tkinter import messagebox

from ..number_format import round_numeric_value
from .. import image_index as image_index_utils

from .excel_tracking import (
    TRACKING_COLUMN,
    update_tracking_column,
    update_tracking_cache,
)

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


def _to_float(value):
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _extract_rgb(color_obj):
    if not color_obj:
        return None
    rgb = getattr(color_obj, "rgb", None)
    if not rgb:
        return None
    rgb = str(rgb).strip().lstrip("#")
    if len(rgb) == 8:
        rgb = rgb[2:]
    if len(rgb) != 6:
        return None
    try:
        int(rgb, 16)
    except ValueError:
        return None
    return rgb.upper()


def _is_red_color(color_obj):
    rgb = _extract_rgb(color_obj)
    if rgb:
        r = int(rgb[0:2], 16)
        g = int(rgb[2:4], 16)
        b = int(rgb[4:6], 16)
        return r >= 170 and g <= 110 and b <= 110
    indexed = getattr(color_obj, "indexed", None)
    return indexed in {10}


def _cell_marked_red(cell):
    fill = getattr(cell, "fill", None)
    if fill and getattr(fill, "patternType", None) not in (None, "none"):
        if _is_red_color(getattr(fill, "fgColor", None)) or _is_red_color(
            getattr(fill, "start_color", None)
        ):
            return True
    font = getattr(cell, "font", None)
    if font and _is_red_color(getattr(font, "color", None)):
        return True
    return False


def _collect_red_nonpositive_issues(app, row_indices, sheet_fields_all):
    if not row_indices or not sheet_fields_all:
        return []

    task_rows = sorted({idx for idx in row_indices if isinstance(idx, int) and idx >= 0})
    if not task_rows:
        return []

    issues = set()
    try:
        wb = load_workbook(app.excel_path, data_only=False, read_only=True)
    except Exception:
        logger.exception("Failed to open workbook for red-cell validation")
        return ["Nie udało się sprawdzić pól oznaczonych na czerwono w Excelu."]

    for sheet_name, columns in sheet_fields_all.items():
        if sheet_name not in wb.sheetnames:
            continue
        df = app.dataframes.get(sheet_name)
        if df is None or df.empty:
            continue
        ws = wb[sheet_name]
        try:
            header_row = next(ws.iter_rows(min_row=1, max_row=1))
        except StopIteration:
            continue
        header_indices = {}
        for col_idx, header_cell in enumerate(header_row, start=1):
            header_value = header_cell.value
            if header_value is None:
                continue
            header_indices[str(header_value)] = col_idx

        for col_name in columns:
            col_idx = header_indices.get(col_name)
            if not col_idx:
                continue
            for row_idx in task_rows:
                if row_idx < 0 or row_idx >= len(df):
                    continue
                excel_row = row_idx + 2
                cell = ws.cell(row=excel_row, column=col_idx)
                if not _cell_marked_red(cell):
                    continue
                value = df.iloc[row_idx].get(col_name)
                number = _to_float(value)
                if number is None or number > 0:
                    continue
                shown = round_numeric_value(value)
                issues.add(
                    f"Wiersz {row_idx + 1}: {sheet_name}:{col_name} ma wartość {shown} (<= 0) i jest oznaczone na czerwono."
                )
    return sorted(issues)


def _is_http_value(value):
    text = str(value or "").strip().lower()
    return text.startswith("http://") or text.startswith("https://")


def _check_remote_image_exists(url, timeout=3):
    response = None
    try:
        response = requests.head(url, allow_redirects=True, timeout=timeout)
        status = response.status_code
        if status in {405} or status >= 400:
            with suppress(Exception):
                response.close()
            response = requests.get(url, stream=True, timeout=timeout)
            status = response.status_code
        return 200 <= status < 400
    except requests.RequestException:
        return False
    finally:
        if response is not None:
            with suppress(Exception):
                response.close()


def _collect_missing_image_issues(app, row_indices):
    image_fields = set(getattr(app, "image_fields", set()) or [])
    if not row_indices or not image_fields:
        return []

    data_rows = sorted({idx for idx in row_indices if isinstance(idx, int) and idx >= 0})
    if not data_rows:
        return []

    issues = set()
    remote_cache = {}
    max_remote_checks = 30

    dynamic_fields = []
    for field in image_fields:
        if ":" not in field:
            continue
        sheet, col = field.split(":", 1)
        dynamic_fields.append((field, sheet, col))

    for row_idx in data_rows:
        row_no = row_idx + 1
        for field, sheet, col in dynamic_fields:
            df = app.dataframes.get(sheet)
            if df is None or row_idx >= len(df):
                continue
            value = df.iloc[row_idx].get(col)
            try:
                if pd.isna(value):
                    continue
            except Exception:
                pass
            text = str(value or "").strip()
            if not text:
                continue
            if _is_http_value(text):
                if text not in remote_cache and len(remote_cache) < max_remote_checks:
                    remote_cache[text] = _check_remote_image_exists(text)
                if text in remote_cache and not remote_cache[text]:
                    issues.add(
                        f"Wiersz {row_no}: nie można pobrać obrazu z URL dla pola {field}: {text}"
                    )
                continue
            if not app.find_local_image(text):
                issues.add(
                    f"Wiersz {row_no}: nie znaleziono pliku obrazu dla pola {field}: {text}"
                )

    static_entries = _collect_static_entries(app)
    for field in image_fields:
        if ":" in field:
            continue
        value = static_entries.get(field, "")
        text = str(value or "").strip()
        if not text:
            continue
        if _is_http_value(text):
            if text not in remote_cache and len(remote_cache) < max_remote_checks:
                remote_cache[text] = _check_remote_image_exists(text)
            if text in remote_cache and not remote_cache[text]:
                issues.add(
                    f"Pole statyczne {field}: nie można pobrać obrazu z URL: {text}"
                )
            continue
        if not app.find_local_image(text):
            issues.add(
                f"Pole statyczne {field}: nie znaleziono pliku obrazu: {text}"
            )

    if len(remote_cache) >= max_remote_checks:
        issues.add(
            "Pominięto część zdalnych sprawdzeń URL obrazów (limit 30 unikalnych linków na uruchomienie)."
        )
    return sorted(issues)


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


def _notify_generation_complete(app, report):
    if hasattr(app, "on_generation_complete"):
        _ui_call(app, app.on_generation_complete, report)


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


def _wrap_text_lines(text, font_name, font_size, max_width):
    if text is None:
        return [""]
    text = str(text)
    if not text:
        return [""]
    lines = []
    for para in text.splitlines():
        if not para:
            lines.append("")
            continue
        words = para.split()
        if not words:
            lines.append("")
            continue
        current = words[0]
        for word in words[1:]:
            candidate = f"{current} {word}"
            if pdfmetrics.stringWidth(candidate, font_name, font_size) <= max_width:
                current = candidate
                continue
            lines.append(current)
            if pdfmetrics.stringWidth(word, font_name, font_size) <= max_width:
                current = word
            else:
                part = ""
                for ch in word:
                    candidate_part = f"{part}{ch}"
                    if part and pdfmetrics.stringWidth(candidate_part, font_name, font_size) > max_width:
                        lines.append(part)
                        part = ch
                    else:
                        part = candidate_part
                current = part
        lines.append(current)
    return lines


def _fit_text_lines(text, font_name, max_font_size, box_width, box_height, pad=2):
    max_size = max(1, int(round(max_font_size)))
    max_width = max(1, box_width - pad * 2)
    max_height = max(1, box_height - pad * 2)
    first_shrink = 3
    second_shrink = 3
    min_size_stage1 = max(1, max_size - first_shrink)
    min_size_stage2 = max(1, max_size - first_shrink - second_shrink)

    def split_lines(value):
        if value is None:
            return [""]
        value = str(value)
        if not value:
            return [""]
        lines = value.splitlines()
        return lines if lines else [""]

    def lines_fit_no_wrap(lines, size):
        line_height = pdfmetrics.getAscent(font_name, size) - pdfmetrics.getDescent(font_name, size)
        if line_height * len(lines) > max_height:
            return False
        for line in lines:
            if pdfmetrics.stringWidth(line, font_name, size) > max_width:
                return False
        return True

    raw_lines = split_lines(text)
    for size in range(max_size, min_size_stage1 - 1, -1):
        if lines_fit_no_wrap(raw_lines, size):
            return size, raw_lines

    fallback_size = min_size_stage1
    fallback_lines = _wrap_text_lines(text, font_name, fallback_size, max_width)
    for size in range(max_size, min_size_stage1 - 1, -1):
        lines = _wrap_text_lines(text, font_name, size, max_width)
        if lines_fit_no_wrap(lines, size):
            return size, lines
        fallback_size = size
        fallback_lines = lines

    for size in range(min_size_stage1 - 1, min_size_stage2 - 1, -1):
        lines = _wrap_text_lines(text, font_name, size, max_width)
        if lines_fit_no_wrap(lines, size):
            return size, lines
        fallback_size = size
        fallback_lines = lines

    return max(1, fallback_size), fallback_lines


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
    value_str = value if isinstance(value, str) else str(value)
    if getattr(element, "is_image", False) and value_str:
        if value_str.lower().startswith("http"):
            try:
                resp = requests.get(value_str, timeout=5)
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
                logger.exception("Failed to load remote image %s", value_str)
        local_path = app.find_local_image(value_str)
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
    box_width = element.width / app.scale
    box_height = element.height / app.scale
    base_font_size = element.font_size / app.scale
    max_font_size = getattr(element, "max_font_size", element.font_size) / app.scale
    auto_fit = getattr(element, "auto_font", True)
    pad = 2 if auto_fit else 0
    if auto_fit:
        font_size, lines = _fit_text_lines(
            value_str, font_name, max_font_size, box_width, box_height, pad=pad
        )
    else:
        font_size = base_font_size
        lines = value_str.splitlines() or [value_str]
    c.setFont(font_name, font_size)
    ascent = pdfmetrics.getAscent(font_name, font_size)
    descent = pdfmetrics.getDescent(font_name, font_size)
    line_height = ascent - descent
    total_height = line_height * len(lines)
    bottom = y + (box_height - total_height) / 2
    baseline_last = bottom - descent
    for idx, line in enumerate(lines):
        baseline = baseline_last + line_height * (len(lines) - 1 - idx)
        if element.align == "center":
            c.drawCentredString(x + box_width / 2, baseline, line)
        elif element.align == "right":
            c.drawRightString(x + box_width - pad, baseline, line)
        else:
            c.drawString(x + pad, baseline, line)


_render_context = None
_render_proxy = None


class RenderAppProxy:
    def __init__(self, scale, excel_dir, image_dirs=None, image_index_path=""):
        self.scale = scale
        self.excel_dir = excel_dir or ""
        self.image_dirs = list(image_dirs or [])
        self._image_cache = {}
        self._image_index_data = None
        roots = image_index_utils.normalize_roots([self.excel_dir] + self.image_dirs)
        if roots:
            self._image_index_data = image_index_utils.load_index_for_roots(
                roots,
                index_path=image_index_path or None,
            )

    def find_local_image(self, filename):
        if not filename:
            return None
        if not self.excel_dir and not self.image_dirs:
            return None
        name = str(filename).strip()
        if not name:
            return None
        key = name.lower()
        if key in self._image_cache:
            return self._image_cache[key]

        roots = [self.excel_dir] + list(self.image_dirs or [])
        seen = set()
        search_roots = []
        for root in roots:
            if not root:
                continue
            root = os.path.abspath(root)
            if root not in seen:
                seen.add(root)
                search_roots.append(root)

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
        if path is None and self._image_index_data:
            path = image_index_utils.find_in_index(self._image_index_data, name)
        if path is None and stem and not self._image_index_data:
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
        self._image_cache[key] = path
        return path


def _init_render_context(context):
    global _render_context, _render_proxy
    _render_context = context
    _render_proxy = RenderAppProxy(
        context["scale"],
        context["excel_dir"],
        context.get("image_dirs", []),
        context.get("image_index_path", ""),
    )


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
            "max_font_size": getattr(element, "max_font_size", element.font_size),
            "bold": element.bold,
            "text_color": element.text_color,
            "bg_color": element.bg_color,
            "bg_visible": element.bg_visible,
            "align": element.align,
            "auto_font": getattr(element, "auto_font", True),
            "layer": element.layer,
            "is_image": getattr(element, "is_image", False),
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
    image_fields = set(context.get("image_fields", []))
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
                        max_font_size=conf.get(
                            "max_font_size",
                            (el.max_font_size / scale if el and hasattr(el, "max_font_size") else font_size),
                        )
                        * scale,
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
                        is_image=(
                            conf.get("is_image")
                            if conf is not None and "is_image" in conf
                            else (el.is_image if el and hasattr(el, "is_image") else fname in image_fields)
                        ),
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
    report = {
        "status": "started",
        "excel_path": getattr(app, "excel_path", ""),
        "output_dir": "",
        "total_rows": 0,
        "total_tasks": 0,
        "processed_rows": 0,
        "skipped_rows": 0,
        "new_pdfs": [],
        "updated_pdfs": [],
        "errors": [],
        "warnings": [],
    }

    def finish_now(status_label, status_code):
        if hasattr(app, "finish_generation_ui"):
            app.finish_generation_ui(status_label)
        else:
            _ui_finish(app, status_label)
        report["status"] = status_code
        _notify_generation_complete(app, report)

    if not app.excel_path or not app.dataframes:
        report["errors"].append("Brak danych do generowania.")
        messagebox.showerror("Błąd", "Brak danych do generowania")
        finish_now("Brak danych", "no_data")
        return False

    _, first_df = next(iter(app.dataframes.items()))
    total_rows = len(first_df)
    report["total_rows"] = total_rows
    if total_rows == 0:
        messagebox.showinfo("Info", "Brak wierszy w pliku Excel")
        finish_now("Brak wierszy", "no_rows")
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

    sheet_fields_all = {}
    sheet_fields_tracking = {}
    excluded_fields = getattr(app, "tracking_excluded", set())
    for field in dynamic_fields:
        sheet, col = field.split(":", 1)
        if col == TRACKING_COLUMN:
            continue
        sheet_fields_all.setdefault(sheet, set()).add(col)
        if field in excluded_fields:
            continue
        sheet_fields_tracking.setdefault(sheet, set()).add(col)

    cache_rows = None
    cache_changed_cols = {}
    excel_rows = None
    tracking_mode = "none"

    try:
        cache_rows, cache_changed_cols = update_tracking_cache(
            app.excel_path, app.dataframes, total_rows, sheet_fields=sheet_fields_tracking
        )
    except Exception:
        logger.exception("Failed to update tracking cache")
        report["errors"].append(
            "Nie udało się zaktualizować lokalnego cache śledzenia zmian."
        )
        cache_rows = None
        cache_changed_cols = {}

    try:
        excel_rows = update_tracking_column(
            app.excel_path, app.dataframes, total_rows, sheet_fields=sheet_fields_tracking
        )
        tracking_mode = "excel"
    except Exception:
        logger.exception("Failed to update tracking column")
        report["errors"].append(
            "Nie udało się zaktualizować kolumny kontrolnej w Excelu."
        )
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

    tracked_columns = sum(len(cols) for cols in sheet_fields_tracking.values())
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
    report["output_dir"] = output_dir

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
        existed_before = os.path.exists(pdf_path)
        if existed_before and changed_rows is not None and idx not in changed_rows:
            continue
        row_values = {}
        for sheet, columns in sheet_fields_all.items():
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
                "existed_before": existed_before,
            }
        )

    if not tasks:
        report["total_tasks"] = 0
        report["skipped_rows"] = total_rows
        messagebox.showinfo(
            "Info", "Brak zmian - wszystkie pliki PDF są aktualne."
        )
        if report["warnings"]:
            messagebox.showwarning(
                "Uwaga",
                "Brak zmian w PDF, ale wykryto ostrzeżenia jakości danych: "
                f"{len(report['warnings'])}.",
            )
        _ui_counts(app, total_rows=total_rows, total_tasks=0, processed=0, skipped=total_rows)
        finish_now("Brak zmian", "no_changes")
        return False

    skipped_rows = total_rows - len(tasks) if changed_rows is not None else 0
    report["total_tasks"] = len(tasks)
    report["skipped_rows"] = skipped_rows
    _ui_counts(
        app,
        total_rows=total_rows,
        total_tasks=len(tasks),
        processed=0,
        skipped=skipped_rows,
    )
    _ui_progress_mode(app, False)
    _ui_status(app, "Sprawdzanie danych i generowanie PDF...")

    if _is_cancelled(app):
        finish_now("Anulowano", "cancelled")
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
        "image_fields": sorted(getattr(app, "image_fields", set())),
        "image_dirs": list(getattr(app, "image_dirs", [])),
        "tasks": tasks,
        "output_dir": output_dir,
        "total_tasks": len(tasks),
        "total_rows": total_rows,
        "skipped_rows": skipped_rows,
        "image_index_path": image_index_utils.get_default_index_path(),
    }

    def worker(payload):
        start_time = time.time()
        roots = image_index_utils.normalize_roots(
            [payload["excel_dir"]] + list(payload.get("image_dirs", []))
        )
        if roots:
            _ui_status(app, "Aktualizacja indeksu obrazów...")
            try:
                image_index_utils.ensure_index(
                    roots,
                    index_path=payload.get("image_index_path") or None,
                    max_age_hours=24,
                    force_rebuild=False,
                )
            except Exception:
                logger.exception("Failed to build/load image index for generation")
            _ui_status(app, "Sprawdzanie danych i generowanie PDF...")

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
            "image_fields": payload.get("image_fields", []),
            "image_dirs": payload.get("image_dirs", []),
            "image_index_path": payload.get("image_index_path", ""),
        }

        tasks_local = payload["tasks"]
        total = payload["total_tasks"]
        failures = []
        completed = 0
        cancelled = False
        new_files = []
        updated_files = []
        max_workers = max(1, min(len(tasks_local), os.cpu_count() or 1))

        validation_warnings = []
        validation_row_indices = [task.get("idx") for task in tasks_local]
        try:
            validation_warnings.extend(
                _collect_red_nonpositive_issues(app, validation_row_indices, sheet_fields_all)
            )
        except Exception:
            logger.exception("Failed while validating red-marked numeric fields")
            validation_warnings.append(
                "Nie udało się sprawdzić pól oznaczonych na czerwono."
            )
        try:
            validation_warnings.extend(_collect_missing_image_issues(app, validation_row_indices))
        except Exception:
            logger.exception("Failed while validating image references")
            validation_warnings.append(
                "Nie udało się sprawdzić ścieżek i linków obrazów."
            )
        if validation_warnings:
            report["warnings"] = sorted(set(validation_warnings))
            sample = ", ".join(report["warnings"][:3])
            logger.warning(
                "Detected data quality warnings: %s (sample: %s)",
                len(report["warnings"]),
                sample,
            )
            for warning_text in report["warnings"][:200]:
                logger.warning("Data warning: %s", warning_text)

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
                    result = future.result()
                    final_pdf = result.get("pdf_path", task["pdf_path"])
                    if task.get("existed_before"):
                        updated_files.append(final_pdf)
                    else:
                        new_files.append(final_pdf)
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
                report["processed_rows"] = completed
                report["new_pdfs"] = sorted(new_files)
                report["updated_pdfs"] = sorted(updated_files)
                for idx, name, err in failures:
                    if idx >= 0:
                        if name:
                            report["errors"].append(
                                f"Wiersz {idx + 1} ({name}): {err}"
                            )
                        else:
                            report["errors"].append(f"Wiersz {idx + 1}: {err}")
                    else:
                        report["errors"].append(f"Błąd wykonawcy: {err}")

                if cancelled:
                    report["status"] = "cancelled"
                    if hasattr(app, "finish_generation_ui"):
                        app.finish_generation_ui("Anulowano")
                    else:
                        _ui_finish(app, "Anulowano")
                    _notify_generation_complete(app, report)
                    return
                if failures:
                    report["status"] = "error"
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
                    report["status"] = "success"
                    if hasattr(app, "finish_generation_ui"):
                        app.finish_generation_ui("Zakończono")
                    else:
                        _ui_finish(app, "Zakończono")
                    messagebox.showinfo(
                        "Zakończono", f"Pliki zapisane w {payload['output_dir']}"
                    )
                    if report["warnings"]:
                        messagebox.showwarning(
                            "Uwaga",
                            "Wykryto ostrzeżenia jakości danych: "
                            f"{len(report['warnings'])}. "
                            "Szczegóły są dostępne w raporcie e-mail i logach.",
                        )
                _notify_generation_complete(app, report)

            app.after(0, finish)

    threading.Thread(target=worker, args=(worker_payload,), daemon=True).start()
    return True
