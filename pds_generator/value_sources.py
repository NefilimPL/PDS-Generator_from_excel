from __future__ import annotations

import datetime as dt
import os

VALUE_SOURCE_DEFAULT = "default"
VALUE_SOURCE_FILE_DATE = "file_date"

FILE_DATE_KIND_MODIFIED = "modified"
FILE_DATE_KIND_CREATED = "created"

FILE_DATE_FORMAT = "%d.%m.%Y"


def normalize_value_source(value):
    if str(value or "").strip().lower() == VALUE_SOURCE_FILE_DATE:
        return VALUE_SOURCE_FILE_DATE
    return VALUE_SOURCE_DEFAULT


def normalize_file_date_kind(value):
    if str(value or "").strip().lower() == FILE_DATE_KIND_CREATED:
        return FILE_DATE_KIND_CREATED
    return FILE_DATE_KIND_MODIFIED


def _element_config_value(element, key, default=None):
    if element is None:
        return default
    if isinstance(element, dict):
        return element.get(key, default)
    return getattr(element, key, default)


def resolve_file_date_value(path, kind=FILE_DATE_KIND_MODIFIED):
    if not path:
        return ""
    file_date_kind = normalize_file_date_kind(kind)
    try:
        stat_result = os.stat(path)
        if file_date_kind == FILE_DATE_KIND_CREATED:
            timestamp = getattr(stat_result, "st_birthtime", None)
            if timestamp is None:
                timestamp = stat_result.st_ctime
        else:
            timestamp = stat_result.st_mtime
    except OSError:
        return ""
    return dt.datetime.fromtimestamp(timestamp).strftime(FILE_DATE_FORMAT)


def resolve_element_value(default_value, element=None, file_path=""):
    if normalize_value_source(_element_config_value(element, "value_source")) == VALUE_SOURCE_FILE_DATE:
        return resolve_file_date_value(
            file_path,
            _element_config_value(element, "file_date_kind"),
        )
    return default_value
