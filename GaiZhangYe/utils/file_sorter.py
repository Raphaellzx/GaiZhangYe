"""Compatibility shim — re-export sorting helpers from core.basic.file_processor."""
from GaiZhangYe.core.basic.file_processor import (
    windows_natural_sort_key,
    sort_files_windows_style,
    sort_dicts_by_name_windows_style,
)

__all__ = [
    "windows_natural_sort_key",
    "sort_files_windows_style",
    "sort_dicts_by_name_windows_style",
]
