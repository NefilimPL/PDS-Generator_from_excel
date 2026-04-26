from __future__ import annotations

import logging
import os
from collections import OrderedDict
from functools import lru_cache
from io import BytesIO
from types import SimpleNamespace

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from reportlab.graphics import renderPM
from reportlab.graphics.shapes import Drawing, Image as DrawingImage, Rect, String
from reportlab.pdfbase import pdfmetrics

from .. import image_index as image_index_utils
from ..layout_dependencies import DEFAULT_GRID_SIZE, apply_layout_dependencies
from ..text_layout import fit_text_lines, pdf_font_name
from .pdf_export import (
    DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
    RenderAppProxy,
    _background_rgb_for_element,
    _collect_dependencies,
    _collect_element_specs,
    _collect_group_specs,
    _prepare_pdf_image,
    to_reportlab_color,
)

logger = logging.getLogger(__name__)

DEFAULT_PREVIEW_DPI = 108

_WINDOWS_FONT_DIR = os.path.join(
    os.environ.get("WINDIR", r"C:\Windows"),
    "Fonts",
)
_PREVIEW_FONT_CANDIDATES = {
    False: (
        os.path.join(_WINDOWS_FONT_DIR, "arial.ttf"),
        os.path.join(_WINDOWS_FONT_DIR, "ARIAL.TTF"),
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "arial.ttf",
        "LiberationSans-Regular.ttf",
        "DejaVuSans.ttf",
    ),
    True: (
        os.path.join(_WINDOWS_FONT_DIR, "arialbd.ttf"),
        os.path.join(_WINDOWS_FONT_DIR, "ARIALBD.TTF"),
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "arialbd.ttf",
        "LiberationSans-Bold.ttf",
        "DejaVuSans-Bold.ttf",
    ),
}


def _is_empty_value(value):
    try:
        return value == "" or pd.isna(value)
    except TypeError:
        return value == ""


def _image_has_transparency(image):
    if image.mode in {"RGBA", "LA"}:
        alpha = image.getchannel("A")
        alpha_min, _alpha_max = alpha.getextrema()
        return alpha_min < 255
    if image.mode == "P":
        return image.info.get("transparency") is not None
    return False


def _flatten_preview_image(image, background_rgb):
    rgba_image = image.convert("RGBA")
    base = Image.new(
        "RGBA",
        rgba_image.size,
        (
            int(background_rgb[0]),
            int(background_rgb[1]),
            int(background_rgb[2]),
            255,
        ),
    )
    base.alpha_composite(rgba_image)
    return base.convert("RGB")


def _prepared_image_to_preview_source(prepared_image, background_rgb):
    source_image = getattr(prepared_image, "image", None)
    if source_image is not None:
        prepared = source_image.copy()
    else:
        buffer = getattr(prepared_image, "buffer", None)
        if buffer is None:
            return None
        if hasattr(buffer, "getvalue"):
            data = buffer.getvalue()
        else:
            position = None
            try:
                position = buffer.tell()
                buffer.seek(0)
                data = buffer.read()
            finally:
                if position is not None:
                    try:
                        buffer.seek(position)
                    except Exception:
                        pass
        with Image.open(BytesIO(data)) as opened:
            opened.load()
            prepared = opened.copy()

    if _image_has_transparency(prepared):
        return _flatten_preview_image(prepared, background_rgb)
    if prepared.mode not in {"RGB", "L"}:
        return prepared.convert("RGB")
    return prepared


def _preview_color_to_rgb(value):
    color = to_reportlab_color(value)
    return (
        int(round(float(getattr(color, "red", 0.0)) * 255.0)),
        int(round(float(getattr(color, "green", 0.0)) * 255.0)),
        int(round(float(getattr(color, "blue", 0.0)) * 255.0)),
    )


@lru_cache(maxsize=64)
def _load_preview_font(bold, size_px):
    size_px = max(1, int(round(size_px)))
    for candidate in _PREVIEW_FONT_CANDIDATES[bool(bold)]:
        try:
            return ImageFont.truetype(candidate, size_px)
        except OSError:
            continue
    return ImageFont.load_default()


def _draw_preview_element(app_proxy, drawing, element, value, x, y):
    value_str = value if isinstance(value, str) else str(value)
    box_width = element.width / app_proxy.scale
    box_height = element.height / app_proxy.scale
    if getattr(element, "is_image", False) and value_str:
        try:
            prepared_image = _prepare_pdf_image(
                app_proxy,
                value_str,
                box_width,
                box_height,
                auto_zoom=getattr(element, "image_auto_zoom", False),
                background_rgb=_background_rgb_for_element(element),
            )
        except Exception:
            logger.exception("Failed to prepare image %s for preview", value_str)
        else:
            if prepared_image is not None:
                source_image = _prepared_image_to_preview_source(
                    prepared_image,
                    _background_rgb_for_element(element),
                )
                if source_image is not None:
                    drawing.add(DrawingImage(x, y, box_width, box_height, source_image))
                    return
    if element.bg_visible:
        drawing.add(
            Rect(
                x,
                y,
                box_width,
                box_height,
                fillColor=to_reportlab_color(element.bg_color),
                strokeColor=None,
            )
        )
    font_name = pdf_font_name(getattr(element, "bold", False))
    base_font_size = element.font_size / app_proxy.scale
    max_font_size = getattr(element, "max_font_size", element.font_size) / app_proxy.scale
    auto_fit = getattr(element, "auto_font", True)
    pad = 2 if auto_fit else 0
    if auto_fit:
        font_size, lines = fit_text_lines(
            value_str,
            font_name,
            max_font_size,
            box_width,
            box_height,
            pad=pad,
        )
    else:
        font_size = base_font_size
        lines = value_str.splitlines() or [value_str]
    ascent = pdfmetrics.getAscent(font_name, font_size)
    descent = pdfmetrics.getDescent(font_name, font_size)
    line_height = ascent - descent
    total_height = line_height * len(lines)
    bottom = y + (box_height - total_height) / 2
    baseline_last = bottom - descent
    fill_color = to_reportlab_color(element.text_color)
    for idx, line in enumerate(lines):
        baseline = baseline_last + line_height * (len(lines) - 1 - idx)
        text_anchor = "start"
        anchor_x = x + pad
        if element.align == "center":
            text_anchor = "middle"
            anchor_x = x + box_width / 2
        elif element.align == "right":
            text_anchor = "end"
            anchor_x = x + box_width - pad
        drawing.add(
            String(
                anchor_x,
                baseline,
                line,
                fontName=font_name,
                fontSize=font_size,
                fillColor=fill_color,
                textAnchor=text_anchor,
            )
        )


def _draw_preview_element_with_pillow(app_proxy, image, page_height, element, value, dpi):
    value_str = value if isinstance(value, str) else str(value)
    box_width = element.width / app_proxy.scale
    box_height = element.height / app_proxy.scale
    scale_factor = float(dpi) / 72.0
    x = float(element._preview_x)
    y = float(element._preview_y)
    left = int(round(x * scale_factor))
    top = int(round((page_height - (y + box_height)) * scale_factor))
    width_px = max(1, int(round(box_width * scale_factor)))
    height_px = max(1, int(round(box_height * scale_factor)))

    if getattr(element, "is_image", False) and value_str:
        try:
            prepared_image = _prepare_pdf_image(
                app_proxy,
                value_str,
                box_width,
                box_height,
                auto_zoom=getattr(element, "image_auto_zoom", False),
                background_rgb=_background_rgb_for_element(element),
            )
        except Exception:
            logger.exception("Failed to prepare image %s for preview", value_str)
        else:
            if prepared_image is not None:
                source_image = _prepared_image_to_preview_source(
                    prepared_image,
                    _background_rgb_for_element(element),
                )
                if source_image is not None:
                    rendered = source_image.resize((width_px, height_px), Image.LANCZOS)
                    if _image_has_transparency(rendered):
                        rendered_rgba = rendered.convert("RGBA")
                        image.paste(rendered_rgba, (left, top), rendered_rgba)
                    else:
                        image.paste(rendered.convert("RGB"), (left, top))
                    return

    draw = ImageDraw.Draw(image)
    if element.bg_visible:
        draw.rectangle(
            (left, top, left + width_px - 1, top + height_px - 1),
            fill=_preview_color_to_rgb(element.bg_color),
        )

    font_name = pdf_font_name(getattr(element, "bold", False))
    base_font_size = element.font_size / app_proxy.scale
    max_font_size = getattr(element, "max_font_size", element.font_size) / app_proxy.scale
    auto_fit = getattr(element, "auto_font", True)
    pad = 2 if auto_fit else 0
    if auto_fit:
        font_size, lines = fit_text_lines(
            value_str,
            font_name,
            max_font_size,
            box_width,
            box_height,
            pad=pad,
        )
    else:
        font_size = base_font_size
        lines = value_str.splitlines() or [value_str]

    ascent = pdfmetrics.getAscent(font_name, font_size)
    descent = pdfmetrics.getDescent(font_name, font_size)
    line_height = ascent - descent
    total_height = line_height * len(lines)
    bottom = y + (box_height - total_height) / 2
    baseline_last = bottom - descent
    pad_px = pad * scale_factor
    fill_color = _preview_color_to_rgb(element.text_color)
    font = _load_preview_font(getattr(element, "bold", False), font_size * scale_factor)

    for idx, line in enumerate(lines):
        baseline = baseline_last + line_height * (len(lines) - 1 - idx)
        line_width_px = pdfmetrics.stringWidth(line, font_name, font_size) * scale_factor
        if element.align == "center":
            text_x = left + (width_px - line_width_px) / 2.0
        elif element.align == "right":
            text_x = left + width_px - pad_px - line_width_px
        else:
            text_x = left + pad_px
        text_y = (page_height - (baseline + ascent)) * scale_factor
        draw.text(
            (int(round(text_x)), int(round(text_y))),
            line,
            font=font,
            fill=fill_color,
        )


def _build_preview_context(app, values, hidden_names=None):
    scale = getattr(app, "scale", 1.0) or 1.0
    page_width = float(getattr(app, "page_width", 0) or 0)
    page_height = float(getattr(app, "page_height", 0) or 0)
    find_local_image = getattr(app, "find_local_image", None)
    if callable(find_local_image):
        prepared_cache = getattr(app, "_prepared_pdf_image_cache", None)
        if prepared_cache is None:
            prepared_cache = OrderedDict()
            try:
                setattr(app, "_prepared_pdf_image_cache", prepared_cache)
            except Exception:
                pass
        app_proxy = SimpleNamespace(
            scale=scale,
            find_local_image=find_local_image,
            pdf_image_compression_percent=getattr(
                app,
                "pdf_image_compression_percent",
                DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
            ),
            _prepared_pdf_image_cache=prepared_cache,
        )
    else:
        app_proxy = RenderAppProxy(
            scale,
            os.path.dirname(getattr(app, "excel_path", "") or ""),
            getattr(app, "image_dirs", []),
            image_index_utils.get_default_index_path(),
            getattr(
                app,
                "pdf_image_compression_percent",
                DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
            ),
        )

    element_specs, element_order = _collect_element_specs(app)
    elements = {name: SimpleNamespace(**spec) for name, spec in element_specs.items()}
    hidden = set(hidden_names or [])
    apply_layout_dependencies(
        elements,
        hidden_names=hidden,
        dependencies=_collect_dependencies(app),
        grid_step=getattr(app, "grid_size", DEFAULT_GRID_SIZE) * scale,
    )
    return {
        "scale": scale,
        "page_width": page_width,
        "page_height": page_height,
        "app_proxy": app_proxy,
        "elements": elements,
        "element_order": element_order,
        "groups": _collect_group_specs(app),
        "image_fields": set(getattr(app, "image_fields", set()) or []),
        "hidden": hidden,
        "values": values,
    }


def _iter_preview_layout_items(context):
    scale = context["scale"]
    page_height = context["page_height"]
    elements = context["elements"]
    hidden = context["hidden"]
    values = context["values"]
    image_fields = context["image_fields"]

    for group in context["groups"]:
        g_hidden = set()
        for src, tgt in group["conditions"]:
            if src not in group["fields"] or tgt not in group["fields"]:
                continue
            if _is_empty_value(values.get(src, "")):
                g_hidden.add(tgt)
        positions = group["field_pos"]
        columns = {}
        for fname in group["fields"]:
            if fname in hidden or fname in g_hidden:
                continue
            val = values.get(fname, "")
            if _is_empty_value(val):
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
            col_items.sort(key=lambda item: item[0])
            cur_y = 0
            for _orig_y, fname, width, height, conf, el, val in col_items:
                y0 = cur_y
                while True:
                    overlap = False
                    for px, py, pw, ph in placed:
                        if (
                            x0 < px + pw
                            and x0 + width > px
                            and y0 < py + ph
                            and y0 + height > py
                        ):
                            y0 = py + ph
                            overlap = True
                            break
                    if not overlap:
                        break
                if y0 + height > group["height"] / scale:
                    continue
                font_size = conf.get(
                    "font_size",
                    el.font_size / scale if el else 12,
                )
                dummy = SimpleNamespace(
                    width=width * scale,
                    height=height * scale,
                    font_size=font_size * scale,
                    max_font_size=conf.get(
                        "max_font_size",
                        (
                            el.max_font_size / scale
                            if el and hasattr(el, "max_font_size")
                            else font_size
                        ),
                    )
                    * scale,
                    bold=conf.get("bold", el.bold if el else False),
                    text_color=conf.get(
                        "text_color",
                        el.text_color if el else "black",
                    ),
                    bg_color=conf.get("bg_color", el.bg_color if el else "white"),
                    bg_visible=conf.get(
                        "bg_visible",
                        el.bg_visible if el else True,
                    ),
                    align=conf.get("align", el.align if el else "left"),
                    auto_font=conf.get("auto_font", el.auto_font if el else True),
                    is_image=(
                        conf.get("is_image")
                        if conf is not None and "is_image" in conf
                        else (
                            el.is_image
                            if el and hasattr(el, "is_image")
                            else fname in image_fields
                        )
                    ),
                    image_auto_zoom=(
                        conf.get("image_auto_zoom")
                        if conf is not None and "image_auto_zoom" in conf
                        else (
                            el.image_auto_zoom
                            if el and hasattr(el, "image_auto_zoom")
                            else False
                        )
                    ),
                )
                dummy._preview_x = group["x"] / scale + x0
                dummy._preview_y = page_height - (group["y"] / scale + y0 + height)
                yield dummy, val, dummy._preview_x, dummy._preview_y
                placed.append((x0, y0, width, height))
                cur_y = y0 + height

    for name in context["element_order"]:
        element = elements.get(name)
        if not element or name in hidden:
            continue
        value = values.get(name, "")
        element._preview_x = element.x / scale
        element._preview_y = page_height - (element.y / scale) - (element.height / scale)
        yield element, value, element._preview_x, element._preview_y


def _render_preview_with_reportlab(context, dpi):
    drawing = Drawing(context["page_width"], context["page_height"])
    drawing.add(
        Rect(
            0,
            0,
            context["page_width"],
            context["page_height"],
            fillColor=to_reportlab_color("white"),
            strokeColor=None,
        )
    )
    for element, value, x, y in _iter_preview_layout_items(context):
        _draw_preview_element(context["app_proxy"], drawing, element, value, x, y)

    preview_image = renderPM.drawToPIL(
        drawing,
        dpi=int(max(24, round(float(dpi)))),
        bg=0xFFFFFF,
        backendFmt="RGBA",
    )
    preview_image.load()
    return preview_image


def _render_preview_with_pillow(context, dpi):
    scale_factor = float(dpi) / 72.0
    width_px = max(1, int(round(context["page_width"] * scale_factor)))
    height_px = max(1, int(round(context["page_height"] * scale_factor)))
    preview_image = Image.new("RGB", (width_px, height_px), (255, 255, 255))
    for element, value, _x, _y in _iter_preview_layout_items(context):
        _draw_preview_element_with_pillow(
            context["app_proxy"],
            preview_image,
            context["page_height"],
            element,
            value,
            dpi,
        )
    return preview_image


def render_pdf_preview_image(app, values, hidden_names=None, dpi=DEFAULT_PREVIEW_DPI):
    context = _build_preview_context(app, values, hidden_names=hidden_names)
    target_dpi = float(dpi or DEFAULT_PREVIEW_DPI)
    try:
        return _render_preview_with_reportlab(context, target_dpi)
    except Exception:
        logger.exception("Falling back to Pillow PDF preview renderer")
        return _render_preview_with_pillow(context, target_dpi)
