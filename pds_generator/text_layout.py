from reportlab.pdfbase import pdfmetrics


DEFAULT_FONT_FAMILY = "Helvetica"


def pdf_font_name(bold=False):
    return "Helvetica-Bold" if bold else "Helvetica"


def wrap_text_lines(text, font_name, font_size, max_width):
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
                    if (
                        part
                        and pdfmetrics.stringWidth(
                            candidate_part, font_name, font_size
                        )
                        > max_width
                    ):
                        lines.append(part)
                        part = ch
                    else:
                        part = candidate_part
                current = part
        lines.append(current)
    return lines


def fit_text_lines(text, font_name, max_font_size, box_width, box_height, pad=2):
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
        line_height = pdfmetrics.getAscent(font_name, size) - pdfmetrics.getDescent(
            font_name, size
        )
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
    fallback_lines = wrap_text_lines(text, font_name, fallback_size, max_width)
    for size in range(max_size, min_size_stage1 - 1, -1):
        lines = wrap_text_lines(text, font_name, size, max_width)
        if lines_fit_no_wrap(lines, size):
            return size, lines
        fallback_size = size
        fallback_lines = lines

    for size in range(min_size_stage1 - 1, min_size_stage2 - 1, -1):
        lines = wrap_text_lines(text, font_name, size, max_width)
        if lines_fit_no_wrap(lines, size):
            return size, lines
        fallback_size = size
        fallback_lines = lines

    return max(1, fallback_size), fallback_lines
