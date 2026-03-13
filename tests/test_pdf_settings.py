import unittest

from pds_generator.pdf_settings import (
    DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
    blend_pdf_image_target_size,
    get_pdf_image_compression_profile,
    normalize_pdf_image_compression_percent,
)


class PdfSettingsTests(unittest.TestCase):
    def test_normalize_pdf_image_compression_percent_clamps_range(self):
        self.assertEqual(normalize_pdf_image_compression_percent(-10), 0)
        self.assertEqual(normalize_pdf_image_compression_percent(120), 100)

    def test_get_pdf_image_compression_profile_zero_is_passthrough(self):
        profile = get_pdf_image_compression_profile(0)

        self.assertEqual(profile["compression_percent"], 0)
        self.assertTrue(profile["passthrough"])
        self.assertIsNone(profile["target_dpi"])
        self.assertIsNone(profile["jpeg_quality"])

    def test_normalize_pdf_image_compression_percent_uses_default_for_invalid_input(self):
        self.assertEqual(
            normalize_pdf_image_compression_percent("abc"),
            DEFAULT_PDF_IMAGE_COMPRESSION_PERCENT,
        )

    def test_get_pdf_image_compression_profile_lowers_quality_and_dpi_for_higher_compression(self):
        low = get_pdf_image_compression_profile(10)
        high = get_pdf_image_compression_profile(80)

        self.assertEqual(low["compression_percent"], 10)
        self.assertEqual(high["compression_percent"], 80)
        self.assertFalse(low["passthrough"])
        self.assertFalse(high["passthrough"])
        self.assertGreater(low["target_dpi"], high["target_dpi"])
        self.assertGreater(low["jpeg_quality"], high["jpeg_quality"])

    def test_blend_pdf_image_target_size_is_gradual(self):
        zero = blend_pdf_image_target_size(4000, 3000, 800, 600, 0)
        one = blend_pdf_image_target_size(4000, 3000, 800, 600, 1)
        ten = blend_pdf_image_target_size(4000, 3000, 800, 600, 10)
        hundred = blend_pdf_image_target_size(4000, 3000, 800, 600, 100)

        self.assertEqual(zero, (4000, 3000))
        self.assertEqual(one, (3968, 2976))
        self.assertEqual(ten, (3680, 2760))
        self.assertEqual(hundred, (800, 600))


if __name__ == "__main__":
    unittest.main()
