import datetime as dt
import math
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

DEFAULT_DECIMALS = 3

_NUMERIC_PATTERN = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$")


def _parse_decimal(value, allow_strings=True, require_separator=False):
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return None
    if isinstance(value, Decimal):
        return value, ".", False
    if isinstance(value, str):
        if not allow_strings:
            return None
        raw = value.strip()
        if not raw:
            return None
        if require_separator and "." not in raw and "," not in raw:
            return None
        raw = raw.replace(" ", "").replace("\u00a0", "")
        sign = ""
        if raw and raw[0] in "+-":
            sign = raw[0]
            raw = raw[1:]
        if not raw:
            return None
        if not any(ch.isdigit() for ch in raw):
            return None
        if raw[0] in ".,":
            raw = "0" + raw
        sep = "."
        if "." in raw and "," in raw:
            if raw.rfind(".") > raw.rfind(","):
                sep = "."
                raw = raw.replace(",", "")
            else:
                sep = ","
                raw = raw.replace(".", "")
            raw = raw.replace(sep, ".")
        elif "," in raw:
            sep = ","
            if raw.count(",") > 1:
                parts = raw.split(",")
                raw = "".join(parts[:-1]) + "." + parts[-1]
            else:
                raw = raw.replace(",", ".")
        elif "." in raw:
            sep = "."
            if raw.count(".") > 1:
                parts = raw.split(".")
                raw = "".join(parts[:-1]) + "." + parts[-1]
        else:
            sep = "."
        raw = sign + raw
        if not _NUMERIC_PATTERN.match(raw):
            return None
        try:
            return Decimal(raw), sep, True
        except InvalidOperation:
            return None
    try:
        dec = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    return dec, ".", False


def _format_decimal(value, decimals, sep):
    fmt = f"{{0:.{decimals}f}}".format(value)
    fmt = fmt.rstrip("0").rstrip(".")
    if sep == ",":
        fmt = fmt.replace(".", ",")
    return fmt


def analyze_numeric_value(value, decimals=DEFAULT_DECIMALS, allow_strings=True, require_separator=True):
    parsed = _parse_decimal(
        value, allow_strings=allow_strings, require_separator=require_separator
    )
    if parsed is None:
        return None
    dec, sep, _is_string = parsed
    quant = Decimal(1).scaleb(-decimals)
    rounded = dec.quantize(quant, rounding=ROUND_HALF_UP)
    if rounded == 0:
        rounded = Decimal(0)
    formatted = _format_decimal(rounded, decimals, sep)
    formatted_dot = _format_decimal(rounded, decimals, ".")
    should_mark = rounded == 0
    changed = rounded != dec
    if rounded == rounded.to_integral_value():
        numeric = int(rounded)
    else:
        numeric = float(rounded)
    return {
        "rounded": rounded,
        "formatted": formatted,
        "formatted_dot": formatted_dot,
        "numeric": numeric,
        "should_mark": should_mark,
        "changed": changed,
    }


def format_numeric_value(value, decimals=DEFAULT_DECIMALS):
    analysis = analyze_numeric_value(value, decimals=decimals)
    if analysis is None:
        return value
    return analysis["formatted"]


def round_numeric_value(value, decimals=DEFAULT_DECIMALS):
    analysis = analyze_numeric_value(value, decimals=decimals)
    if analysis is None:
        return value
    return analysis["numeric"]
