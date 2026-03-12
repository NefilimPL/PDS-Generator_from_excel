from pds_generator.text_layout import fit_text_lines, pdf_font_name


def test_pdf_font_name_matches_weight():
    assert pdf_font_name(False) == "Helvetica"
    assert pdf_font_name(True) == "Helvetica-Bold"


def test_fit_text_lines_is_deterministic_for_same_input():
    args = ("Sonoma OAK Light", "Helvetica", 12, 120, 28)
    first = fit_text_lines(*args)
    second = fit_text_lines(*args)

    assert first == second


def test_fit_text_lines_shrinks_when_box_is_narrower():
    wide_size, wide_lines = fit_text_lines(
        "Sonoma OAK Light",
        "Helvetica",
        12,
        140,
        28,
    )
    narrow_size, narrow_lines = fit_text_lines(
        "Sonoma OAK Light",
        "Helvetica",
        12,
        80,
        28,
    )

    assert narrow_size <= wide_size
    assert len(narrow_lines) >= len(wide_lines)
