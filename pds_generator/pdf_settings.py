DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT = 35
MIN_PDF_IMAGE_COMPRESSION_PERCENT = 0
MAX_PDF_IMAGE_COMPRESSION_PERCENT = 100

MIN_PDF_IMAGE_TARGET_DPI = 80
MAX_PDF_IMAGE_TARGET_DPI = 220

MIN_PDF_IMAGE_JPEG_QUALITY = 30
MAX_PDF_IMAGE_JPEG_QUALITY = 92


def normalize_pdf_image_compression_percent(
    value,
    default=DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
):
    try:
        percent = int(round(float(str(value).strip().replace(",", "."))))
    except (TypeError, ValueError):
        percent = int(default)
    return max(
        MIN_PDF_IMAGE_COMPRESSION_PERCENT,
        min(MAX_PDF_IMAGE_COMPRESSION_PERCENT, percent),
    )


def get_pdf_image_compression_profile(value):
    compression_percent = normalize_pdf_image_compression_percent(value)
    if compression_percent == 0:
        return {
            "compression_percent": 0,
            "passthrough": True,
            "target_dpi": None,
            "jpeg_quality": None,
        }
    ratio = compression_percent / 100.0
    target_dpi = int(
        round(
            MAX_PDF_IMAGE_TARGET_DPI
            - (MAX_PDF_IMAGE_TARGET_DPI - MIN_PDF_IMAGE_TARGET_DPI) * ratio
        )
    )
    jpeg_quality = int(
        round(
            MAX_PDF_IMAGE_JPEG_QUALITY
            - (MAX_PDF_IMAGE_JPEG_QUALITY - MIN_PDF_IMAGE_JPEG_QUALITY) * ratio
        )
    )
    return {
        "compression_percent": compression_percent,
        "passthrough": False,
        "target_dpi": target_dpi,
        "jpeg_quality": jpeg_quality,
    }


def blend_pdf_image_target_size(
    original_width,
    original_height,
    target_width,
    target_height,
    compression_percent,
):
    compression_percent = normalize_pdf_image_compression_percent(compression_percent)
    original_width = max(1, int(original_width))
    original_height = max(1, int(original_height))
    target_width = max(1, int(target_width))
    target_height = max(1, int(target_height))
    if compression_percent == 0:
        return original_width, original_height

    target_ratio = float(target_width) / float(target_height)
    original_ratio = float(original_width) / float(original_height)
    if original_ratio > target_ratio:
        source_width = original_width
        source_height = max(1, int(round(float(source_width) / target_ratio)))
    else:
        source_height = original_height
        source_width = max(1, int(round(float(source_height) * target_ratio)))

    target_width = min(source_width, target_width)
    target_height = min(source_height, target_height)

    strength = compression_percent / 100.0
    blended_width = source_width - (source_width - target_width) * strength
    blended_height = source_height - (source_height - target_height) * strength
    return max(1, int(round(blended_width))), max(1, int(round(blended_height)))
