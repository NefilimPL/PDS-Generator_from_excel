from types import SimpleNamespace

from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas

from pds_generator.gui.pdf_export import _compress_image_for_pdf, draw_pdf_element
from pds_generator.pdf_settings import (
    DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
    blend_pdf_image_target_size,
    get_pdf_image_compression_profile,
)


def _photo_like_image(size=(900, 600)):
    width, height = size
    image = Image.new("RGB", size)
    pixels = [
        (
            (x * 29 + y * 13) % 256,
            (x * 7 + y * 19) % 256,
            (x * 17 + y * 31) % 256,
        )
        for y in range(height)
        for x in range(width)
    ]
    image.putdata(pixels)
    return image


def test_compress_image_for_pdf_downscales_and_prefers_jpeg_for_photos():
    image = _photo_like_image()
    profile = get_pdf_image_compression_profile(DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT)
    expected_size = blend_pdf_image_target_size(
        900,
        600,
        profile["target_dpi"] * 2,
        profile["target_dpi"],
        profile["compression_percent"],
    )

    buffer, encoded_format, pixel_size = _compress_image_for_pdf(image, 144, 72)

    assert encoded_format == "JPEG"
    assert pixel_size == expected_size

    with Image.open(buffer) as saved:
        assert saved.format == "JPEG"
        assert saved.size == expected_size


def test_compress_image_for_pdf_preserves_transparency():
    image = Image.new("RGBA", (800, 800), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rectangle((80, 80, 720, 720), fill=(20, 120, 220, 180))
    profile = get_pdf_image_compression_profile(DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT)
    expected_size = blend_pdf_image_target_size(
        800,
        800,
        profile["target_dpi"],
        profile["target_dpi"],
        profile["compression_percent"],
    )

    buffer, encoded_format, pixel_size = _compress_image_for_pdf(image, 72, 72)

    assert encoded_format == "PNG"
    assert pixel_size == expected_size

    with Image.open(buffer) as saved:
        assert saved.format == "PNG"
        assert saved.size == expected_size
        assert saved.getchannel("A").getextrema()[0] < 255


def test_draw_pdf_element_embeds_smaller_image_payload(tmp_path):
    image_path = tmp_path / "photo.png"
    pdf_path = tmp_path / "out.pdf"
    _photo_like_image().save(image_path, format="PNG")

    class DummyApp:
        scale = 1.0

        def __init__(self, path):
            self.path = str(path)

        def find_local_image(self, filename):
            if filename == "photo":
                return self.path
            return None

    element = SimpleNamespace(
        is_image=True,
        width=144,
        height=96,
        font_size=12,
        max_font_size=12,
        bg_visible=False,
        text_color="black",
        bold=False,
        align="left",
        auto_font=True,
    )

    app = DummyApp(image_path)
    pdf = canvas.Canvas(str(pdf_path), pagesize=(200, 200))
    draw_pdf_element(app, pdf, element, "photo", 20, 20)
    draw_pdf_element(app, pdf, element, "photo", 20, 120)
    pdf.showPage()
    pdf.save()

    assert pdf_path.stat().st_size < image_path.stat().st_size
    assert len(app._prepared_pdf_image_cache) == 1
