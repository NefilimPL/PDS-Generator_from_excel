from PIL import Image, ImageDraw

from pds_generator.image_auto_zoom import compute_safe_content_bbox, render_image_to_box


def _foreground_bbox(image, threshold=242):
    rgba = image.convert("RGBA")
    left = top = right = bottom = None
    for y in range(rgba.height):
        for x in range(rgba.width):
            r, g, b, a = rgba.getpixel((x, y))
            if a <= 16:
                continue
            if r >= threshold and g >= threshold and b >= threshold:
                continue
            if left is None or x < left:
                left = x
            if right is None or x > right:
                right = x
            if top is None or y < top:
                top = y
            if bottom is None or y > bottom:
                bottom = y
    if left is None:
        return None
    return left, top, right, bottom


def test_compute_safe_content_bbox_detects_centered_product():
    image = Image.new("RGB", (400, 300), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((115, 120, 285, 205), fill=(166, 111, 66))

    bbox = compute_safe_content_bbox(image)

    assert bbox is not None
    left, top, right, bottom = bbox
    assert 90 <= left <= 120
    assert 105 <= top <= 125
    assert 280 <= right <= 310
    assert 200 <= bottom <= 220


def test_render_image_to_box_auto_zoom_enlarges_foreground_without_touching_edges():
    image = Image.new("RGB", (400, 300), "white")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((120, 125, 280, 205), radius=6, fill=(166, 111, 66))

    plain = render_image_to_box(image, 220, 220, auto_zoom=False)
    zoomed = render_image_to_box(image, 220, 220, auto_zoom=True)

    plain_bbox = _foreground_bbox(plain)
    zoomed_bbox = _foreground_bbox(zoomed)

    assert plain_bbox is not None
    assert zoomed_bbox is not None
    plain_width = plain_bbox[2] - plain_bbox[0]
    plain_height = plain_bbox[3] - plain_bbox[1]
    zoomed_width = zoomed_bbox[2] - zoomed_bbox[0]
    zoomed_height = zoomed_bbox[3] - zoomed_bbox[1]

    assert zoomed_width > plain_width
    assert zoomed_height > plain_height
    assert zoomed_bbox[0] > 0
    assert zoomed_bbox[1] > 0
    assert zoomed_bbox[2] < zoomed.width - 1
    assert zoomed_bbox[3] < zoomed.height - 1
    assert zoomed.mode == "RGBA"
    assert zoomed.getpixel((0, 0))[3] == 0
    assert zoomed.getpixel((zoomed.width - 1, zoomed.height - 1))[3] == 0
