import logging
import os
import time
import threading
import re
from io import BytesIO
from types import SimpleNamespace

import pandas as pd
import requests
from PIL import Image
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
import tkinter as tk
from tkinter import messagebox

logger = logging.getLogger(__name__)


def to_reportlab_color(value):
    try:
        return colors.HexColor(value)
    except (ValueError, TypeError):
        return colors.toColor(value)


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r"[^\w\s-]", "", str(name))
    cleaned = re.sub(r"\s+", "_", cleaned).strip("_")
    return cleaned


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
    c.setFont(
        "Helvetica-Bold" if element.bold else "Helvetica",
        element.font_size / app.scale,
    )
    if element.align == "center":
        c.drawCentredString(
            x + (element.width / app.scale) / 2,
            y + (element.height / app.scale) / 2,
            str(value),
        )
    elif element.align == "right":
        c.drawRightString(
            x + (element.width / app.scale),
            y + (element.height / app.scale) / 2,
            str(value),
        )
    else:
        c.drawString(x, y + (element.height / app.scale) / 2, str(value))


def generate_pds(app):
    if not app.excel_path or not app.dataframes:
        messagebox.showerror("Błąd", "Brak danych do generowania")
        return
    first_df = next(iter(app.dataframes.values()))
    total_rows = len(first_df)
    if total_rows == 0:
        messagebox.showinfo("Info", "Brak wierszy w pliku Excel")
        return
    output_dir = os.path.join(os.path.dirname(app.excel_path), "PDS")
    os.makedirs(output_dir, exist_ok=True)

    page_width = app.page_width
    page_height = app.page_height

    def worker():
        start_time = time.time()
        for idx in range(total_rows):
            first_val = first_df.iloc[idx, 0] if first_df.shape[1] else ""
            filename = sanitize_filename(first_val) or f"pds_{idx+1}"
            pdf_path = os.path.join(output_dir, f"{filename}.pdf")
            tmp_path = pdf_path + ".tmp"
            c = pdf_canvas.Canvas(tmp_path, pagesize=(page_width, page_height))
            needed = set(app.elements.keys())
            for g in app.groups.values():
                needed.update(g.fields)
            needed.update(app.static_entries.keys())
            values = {}
            for name in needed:
                if ":" in name:
                    sheet, col = name.split(":", 1)
                    df = app.dataframes.get(sheet)
                    value = df.iloc[idx].get(col, "") if df is not None else ""
                else:
                    value = app.static_entries.get(name, tk.StringVar()).get()
                if pd.isna(value):
                    value = ""
                values[name] = value
            group_field_names = {fname for g in app.groups.values() for fname in g.fields}

            hidden = set()
            for src, tgt in app.conditions:
                if src in group_field_names or tgt in group_field_names:
                    continue
                if pd.isna(values.get(src)) or values.get(src) == "":
                    hidden.add(tgt)
            for group in app.groups.values():
                g_hidden = set()
                for src, tgt in group.conditions:
                    if src not in group.fields or tgt not in group.fields:
                        continue
                    if pd.isna(values.get(src)) or values.get(src) == "":
                        g_hidden.add(tgt)
                positions = group.field_pos
                columns = {}
                for fname in group.fields:
                    if fname in hidden or fname in g_hidden:
                        continue
                    val = values.get(fname, "")
                    if val == "":
                        continue
                    conf = group.field_conf.get(fname, {})
                    el = app.elements.get(fname)
                    if not conf and not el:
                        continue
                    width = conf.get("width", el.width if el else 0)
                    height = conf.get("height", el.height if el else 0)
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
                        if y + height > group.height:
                            continue
                        dummy = SimpleNamespace(
                            width=width,
                            height=height,
                            font_size=conf.get("font_size", el.font_size if el else 12),
                            bold=conf.get("bold", el.bold if el else False),
                            text_color=conf.get("text_color", el.text_color if el else "black"),
                            bg_color=conf.get("bg_color", el.bg_color if el else "white"),
                            bg_visible=conf.get("bg_visible", el.bg_visible if el else True),
                            align=conf.get("align", el.align if el else "left"),
                            auto_font=conf.get("auto_font", el.auto_font if el else True),
                        )
                        x_pdf = (group.x + x0) / app.scale
                        y_pdf = page_height - (group.y + y + height) / app.scale
                        draw_pdf_element(app, c, dummy, val, x_pdf, y_pdf)
                        placed.append((x0, y, width, height))
                        cur_y = y + height
            for name, element in sorted(app.elements.items(), key=lambda kv: kv[1].layer):
                if name in hidden:
                    continue
                val = values.get(name, "")
                x = element.x / app.scale
                y = page_height - (element.y / app.scale) - (element.height / app.scale)
                draw_pdf_element(app, c, element, val, x, y)
            c.showPage()
            c.save()
            try:
                os.replace(tmp_path, pdf_path)
            except OSError:
                logger.exception("Failed to replace %s, trying alternative name", pdf_path)
                alt_path = pdf_path.replace(
                    ".pdf", f"_{int(time.time())}.pdf"
                )
                try:
                    os.replace(tmp_path, alt_path)
                except OSError:
                    logger.exception("Failed to move temp PDF to %s", alt_path)
            progress = (idx + 1) / total_rows * 100
            elapsed = time.time() - start_time
            remaining = (elapsed / (idx + 1)) * (total_rows - idx - 1)
            app.progress.after(0, lambda p=progress: app.progress.config(value=p))
            app.time_label.after(0, lambda r=remaining: app.time_label.config(text=f"Pozostały czas: {int(r)} s"))
        app.progress.after(0, lambda: app.progress.config(value=0))
        app.time_label.after(0, lambda: app.time_label.config(text="Zakończono"))
        messagebox.showinfo("Zakończono", f"Pliki zapisane w {output_dir}")

    threading.Thread(target=worker, daemon=True).start()
