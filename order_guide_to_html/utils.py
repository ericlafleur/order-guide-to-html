from __future__ import annotations

from collections import OrderedDict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
import html  # noqa: F401 — re-exported so callers can use base.html.escape
import re
import unicodedata


FOOTNOTE_LINE_RE = re.compile(r"^\s*(\d+)\.\s*(.+?)\s*$")

URL_RE = re.compile(r"(https?://[^\s<]+)")

TRAILING_DIGITS_RE = re.compile(r"^(.*?)(\d+)\s*$")

CODE_IN_PARENS_RE = re.compile(r"\(([A-Z0-9]{2,6})\)")

NON_ALNUM_RE = re.compile(r"[^a-z0-9]+", re.I)

FILLER_TOKENS = {"Truck", "Trucks", "Cars", "Car", "SUV", "SUVs", "Camion", "Camions", "Voiture", "Voitures", "Auto", "Autos", "VUS"}

def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize('NFC', str(value))
    text = text.replace("\xa0", " ").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def slugify(text: str) -> str:
    text = normalize_text(text).replace("/", " ")
    text = re.sub(r"[^A-Za-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("_")

def htmlize_text(text: str) -> str:
    escaped = html.escape(normalize_text(text))
    return URL_RE.sub(r'<a href="\1" target="_blank">\1</a>', escaped)

def unique_preserve_order(items: Iterable[str]) -> List[str]:
    seen = OrderedDict()
    for item in items:
        item = normalize_text(item)
        if item:
            seen[item] = None
    return list(seen.keys())

def sentence_chunks(text: str, max_words: int = 110) -> List[str]:
    text = normalize_text(text)
    if not text:
        return []
    parts = [normalize_text(p) for p in re.split(r'(?<=[.!?])\s+|\n+', text) if normalize_text(p)]
    if not parts:
        return [text]
    chunks: List[str] = []
    current: List[str] = []
    current_words = 0
    for part in parts:
        words = len(part.split())
        if words > max_words and not current:
            raw_words = part.split()
            for i in range(0, len(raw_words), max_words):
                chunks.append(' '.join(raw_words[i:i + max_words]))
            continue
        if current and current_words + words > max_words:
            chunks.append(normalize_text(' '.join(current)))
            current = [part]
            current_words = words
        else:
            current.append(part)
            current_words += words
    if current:
        chunks.append(normalize_text(' '.join(current)))
    return [c for c in chunks if c]

def chunk_list(items: Sequence[str], max_words: int = 110, max_items: int = 8) -> List[List[str]]:
    chunks: List[List[str]] = []
    current: List[str] = []
    current_words = 0
    for item in items:
        item = normalize_text(item)
        if not item:
            continue
        words = len(item.split())
        if current and (current_words + words > max_words or len(current) >= max_items):
            chunks.append(current)
            current = [item]
            current_words = words
        else:
            current.append(item)
            current_words += words
    if current:
        chunks.append(current)
    return chunks

def render_article(title: str, fields: Sequence[Tuple[str, str]], bullet_groups: Sequence[Tuple[str, Sequence[str]]] = ()) -> str:
    parts = [f'<article class="guide-record"><h3>{htmlize_text(title)}</h3>']
    for label, value in fields:
        value = normalize_text(value)
        if value:
            parts.append(f'<p><strong>{html.escape(label)}:</strong> {htmlize_text(value)}</p>')
    for label, items in bullet_groups:
        clean_items = [normalize_text(item) for item in items if normalize_text(item)]
        if not clean_items:
            continue
        parts.append(f'<div class="record-list"><p><strong>{html.escape(label)}:</strong></p><ul>')
        for item in clean_items:
            parts.append(f'<li>{htmlize_text(item)}</li>')
        parts.append('</ul></div>')
    parts.append('</article>')
    return ''.join(parts)

def chunk_feature_items(items: Sequence[str]) -> List[List[str]]:
    return chunk_list(items, max_words=135, max_items=8)

def unique_output_path(output_dir: Path, filename: str, used_names: set[str]) -> Path:
    candidate = filename
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    index = 2
    while candidate in used_names:
        candidate = f'{stem}_{index}{suffix}'
        index += 1
    used_names.add(candidate)
    return output_dir / candidate

def short_slug(text: str, max_len: int = 72) -> str:
    slug = slugify(text)
    if len(slug) <= max_len:
        return slug or 'item'
    trimmed = slug[:max_len].rstrip('_')
    return trimmed or slug[:max_len] or 'item'

def bool_or_none(value: bool) -> Optional[bool]:
    return True if value else None

def numericish(text: str) -> bool:
    return bool(re.search(r'\d', normalize_text(text)))

# ---------------------------------------------------------------------------
# Status code constants (shared by parsing and classification)

from collections import OrderedDict as _OD  # already imported above; re-alias for clarity

STATUS_LABELS = _OD(
    [
        ("S", "Standard Equipment"),
        ("A", "Available"),
        ("D", "ADI Available"),
        ("■", "Included in Equipment Group"),
        ("□", "Included in Equipment Group but upgradeable"),
        ("*", "Indicates availability of feature on multiple models"),
        ("--", "Not Available"),
    ]
)

STATUS_PRIORITY: Dict[str, int] = {
    'Standard Equipment': 0,
    'Included in Equipment Group': 1,
    'Included in Equipment Group but upgradeable': 2,
    'ADI Available': 3,
    'Available': 4,
    'Indicates availability of feature on multiple models': 5,
    'Not Available': 6,
}

# ---------------------------------------------------------------------------
# Manifest helper constants and small utilities

MANIFEST_STANDARDISH_CODES: frozenset = frozenset({'S', '\u25a0', '\u25a1'})
MANIFEST_DRIVE_TOKENS = ('2WD', '4WD', 'AWD', 'FWD', 'RWD', '2RM', '4RM')
MANIFEST_BODY_STYLE_TOKENS = ('Crew Cab', 'Double Cab', 'Regular Cab', 'Cabine classique', 'Cabine double', 'Cabine multiplace', 'Coupe', 'Convertible', 'Coupé')


def first_unique(values: Iterable[str]) -> Optional[str]:
    cleaned = unique_preserve_order(normalize_text(v) for v in values if normalize_text(v))
    if len(cleaned) == 1:
        return cleaned[0]
    return None


def material_note_texts(note_texts: Sequence[str]) -> List[str]:
    notes: List[str] = []
    for note in unique_preserve_order(note_texts):
        if len(note.split()) >= 4:
            notes.append(note)
    return notes
