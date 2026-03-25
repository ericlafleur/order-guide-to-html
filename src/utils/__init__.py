"""Utility functions for order guide processing."""

from .constants import (
    STATUS_LABELS,
    PROPULSION_KEYWORDS,
    VEHICLE_TYPE_SUV_KEYWORDS,
    DRIVE_TYPE_PATTERNS,
    CATEGORY_SEQUENCE,
    CATEGORY_ORDER,
)
from .text_utils import (
    normalize_text,
    slugify,
    htmlize_text,
    unique_preserve_order,
    parse_footnote_map,
    split_main_notes_and_bullets,
    sentence_chunks,
    chunk_list,
    compact_text,
)

__all__ = [
    # Constants
    "STATUS_LABELS",
    "PROPULSION_KEYWORDS",
    "VEHICLE_TYPE_SUV_KEYWORDS",
    "DRIVE_TYPE_PATTERNS",
    "CATEGORY_SEQUENCE",
    "CATEGORY_ORDER",
    # Text utilities
    "normalize_text",
    "slugify",
    "htmlize_text",
    "unique_preserve_order",
    "parse_footnote_map",
    "split_main_notes_and_bullets",
    "sentence_chunks",
    "chunk_list",
    "compact_text",
]
