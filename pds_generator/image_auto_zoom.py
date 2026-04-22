from __future__ import annotations

from collections import deque

from PIL import Image, ImageOps

WHITE_BACKGROUND_THRESHOLD = 242
ALPHA_BACKGROUND_THRESHOLD = 16
ANALYSIS_MAX_DIMENSION = 256
SAFE_MARGIN_RATIO = 0.04


def _has_transparency(image):
    if image.mode in {"RGBA", "LA"}:
        alpha = image.getchannel("A")
        alpha_min, _alpha_max = alpha.getextrema()
        return alpha_min < 255
    if image.mode == "P":
        return image.info.get("transparency") is not None
    return False


def _analysis_copy(image, max_dimension=ANALYSIS_MAX_DIMENSION):
    prepared = ImageOps.exif_transpose(image).convert("RGBA")
    longest_edge = max(prepared.size)
    if longest_edge <= max_dimension:
        return prepared
    scale = float(max_dimension) / float(longest_edge)
    resized = (
        max(1, int(round(prepared.width * scale))),
        max(1, int(round(prepared.height * scale))),
    )
    return prepared.resize(resized, Image.LANCZOS)


def _is_background_pixel(pixel, white_threshold, alpha_threshold):
    r, g, b, a = pixel
    return a <= alpha_threshold or (
        r >= white_threshold and g >= white_threshold and b >= white_threshold
    )


def compute_safe_content_bbox(
    image,
    white_threshold=WHITE_BACKGROUND_THRESHOLD,
    alpha_threshold=ALPHA_BACKGROUND_THRESHOLD,
    margin_ratio=SAFE_MARGIN_RATIO,
):
    prepared = ImageOps.exif_transpose(image)
    if prepared.width <= 1 or prepared.height <= 1:
        return None

    analysis = _analysis_copy(
        prepared,
        max_dimension=ANALYSIS_MAX_DIMENSION,
    )
    width, height = analysis.size
    pixels = analysis.load()
    background = bytearray(width * height)

    for y in range(height):
        row_offset = y * width
        for x in range(width):
            if _is_background_pixel(
                pixels[x, y],
                white_threshold=white_threshold,
                alpha_threshold=alpha_threshold,
            ):
                background[row_offset + x] = 1

    visited = bytearray(width * height)
    queue = deque()

    def seed(x, y):
        idx = y * width + x
        if background[idx] and not visited[idx]:
            visited[idx] = 1
            queue.append((x, y))

    for x in range(width):
        seed(x, 0)
        seed(x, height - 1)
    for y in range(height):
        seed(0, y)
        seed(width - 1, y)

    while queue:
        x, y = queue.popleft()
        for nx, ny in (
            (x - 1, y),
            (x + 1, y),
            (x, y - 1),
            (x, y + 1),
        ):
            if nx < 0 or ny < 0 or nx >= width or ny >= height:
                continue
            idx = ny * width + nx
            if background[idx] and not visited[idx]:
                visited[idx] = 1
                queue.append((nx, ny))

    row_threshold = max(1, width // 100)
    col_threshold = max(1, height // 100)

    def row_has_content(y):
        row_offset = y * width
        return sum(1 for x in range(width) if not visited[row_offset + x]) >= row_threshold

    def col_has_content(x):
        return sum(1 for y in range(height) if not visited[y * width + x]) >= col_threshold

    top = next((y for y in range(height) if row_has_content(y)), None)
    bottom = next((y for y in range(height - 1, -1, -1) if row_has_content(y)), None)
    left = next((x for x in range(width) if col_has_content(x)), None)
    right = next((x for x in range(width - 1, -1, -1) if col_has_content(x)), None)

    if None in {top, bottom, left, right}:
        return None

    scale_x = float(prepared.width) / float(width)
    scale_y = float(prepared.height) / float(height)
    content_left = max(0, int(left * scale_x))
    content_top = max(0, int(top * scale_y))
    content_right = min(prepared.width, int((right + 1) * scale_x))
    content_bottom = min(prepared.height, int((bottom + 1) * scale_y))

    content_width = max(1, content_right - content_left)
    content_height = max(1, content_bottom - content_top)
    margin_x = max(2, int(round(content_width * margin_ratio)))
    margin_y = max(2, int(round(content_height * margin_ratio)))

    crop_left = max(0, content_left - margin_x)
    crop_top = max(0, content_top - margin_y)
    crop_right = min(prepared.width, content_right + margin_x)
    crop_bottom = min(prepared.height, content_bottom + margin_y)

    if (
        crop_left <= 1
        and crop_top <= 1
        and crop_right >= prepared.width - 1
        and crop_bottom >= prepared.height - 1
    ):
        return None

    if crop_right - crop_left <= 0 or crop_bottom - crop_top <= 0:
        return None

    return crop_left, crop_top, crop_right, crop_bottom


def _pad_to_aspect(image, target_width, target_height):
    target_ratio = float(target_width) / float(target_height)
    image_ratio = float(image.width) / float(image.height)
    if abs(image_ratio - target_ratio) < 1e-6:
        return image.convert("RGBA")

    if image_ratio > target_ratio:
        canvas_width = image.width
        canvas_height = max(1, int(round(float(canvas_width) / target_ratio)))
    else:
        canvas_height = image.height
        canvas_width = max(1, int(round(float(canvas_height) * target_ratio)))

    canvas = Image.new("RGBA", (canvas_width, canvas_height), (255, 255, 255, 0))
    offset_x = (canvas_width - image.width) // 2
    offset_y = (canvas_height - image.height) // 2
    paste_image = image.convert("RGBA")
    canvas.paste(paste_image, (offset_x, offset_y), paste_image)
    return canvas


def build_auto_zoom_image(image, target_width, target_height, resize_to_target=False):
    prepared = ImageOps.exif_transpose(image)
    prepared.load()
    crop_box = compute_safe_content_bbox(prepared)
    if crop_box is None:
        return None

    cropped = prepared.crop(crop_box)
    padded = _pad_to_aspect(cropped, target_width, target_height)
    if not resize_to_target:
        return padded
    return padded.resize(
        (max(1, int(target_width)), max(1, int(target_height))),
        Image.LANCZOS,
    )


def render_image_to_box(image, width, height, auto_zoom=False):
    target_size = (max(1, int(width)), max(1, int(height)))
    prepared = ImageOps.exif_transpose(image)
    prepared.load()

    if not auto_zoom:
        return prepared.resize(target_size, Image.LANCZOS)

    zoomed = build_auto_zoom_image(
        prepared,
        target_size[0],
        target_size[1],
        resize_to_target=True,
    )
    if zoomed is None:
        return prepared.resize(target_size, Image.LANCZOS)
    return zoomed
