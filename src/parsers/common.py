"""Common parsing utilities."""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from ..models.base_models import TrimDef
from ..utils.constants import FILLER_TOKENS, STATUS_LABELS
from ..utils.text_utils import normalize_text


def parse_filename_metadata(path: Path) -> Tuple[str, str, str, str]:
    """Extract year, make, model from filename.
    
    Args:
        path: Path to workbook file
        
    Returns:
        Tuple of (year, make, model, vehicle_name)
        
    Raises:
        ValueError: If metadata cannot be extracted
    """
    stem = path.stem
    stem = re.sub(r"\s+Export$", "", stem, flags=re.I)
    stem = re.sub(r"\s+(Retail and Fleet|Retail|Fleet)$", "", stem, flags=re.I)
    parts = stem.split()
    
    if not parts or not re.fullmatch(r"\d{4}", parts[0]):
        raise ValueError(f"Cannot determine year from filename: {path.name}")
    year = parts[0]
    
    if len(parts) < 3:
        raise ValueError(f"Cannot determine make/model from filename: {path.name}")
    make = parts[1]
    
    model_tokens = [p for p in parts[2:] if p not in FILLER_TOKENS]
    if not model_tokens:
        raise ValueError(f"Cannot determine model from filename: {path.name}")
    model = " ".join(model_tokens)
    
    vehicle_name = f"{year} {make} {model}"
    return year, make, model, vehicle_name


def parse_trim_header(value: object) -> Optional[TrimDef]:
    """Parse trim definition from header cell.
    
    Args:
        value: Cell value containing trim header
        
    Returns:
        TrimDef or None if parsing fails
    """
    text = normalize_text(value)
    if not text:
        return None
    
    lines = [normalize_text(x) for x in text.split("\n") if normalize_text(x)]
    if not lines:
        return None
    
    if len(lines) == 1:
        return TrimDef(name=lines[0], code=lines[0], raw_header=text)
    
    return TrimDef(name=" ".join(lines[:-1]), code=lines[-1], raw_header=text)


def parse_status_value(
    raw: str, 
    row_notes: Dict[str, str], 
    sheet_notes: Dict[str, str]
) -> Tuple[str, str, List[str]]:
    """Parse status value and extract associated notes.
    
    Args:
        raw: Raw status value
        row_notes: Row-level footnotes
        sheet_notes: Sheet-level footnotes
        
    Returns:
        Tuple of (status_code, status_label, note_texts)
    """
    raw = normalize_text(raw)
    if not raw:
        return "", "", []
    
    m = re.match(r"^(--|[A-Z]+|[■□*]+)(.*)$", raw)
    if m:
        code = m.group(1)
        suffix = m.group(2)
    else:
        code = raw
        suffix = ""
    
    label = STATUS_LABELS.get(code, code)
    note_ids = re.findall(r"\d+", suffix)
    
    notes: List[str] = []
    for note_id in note_ids:
        if note_id in row_notes:
            notes.append(row_notes[note_id])
        elif note_id in sheet_notes:
            notes.append(sheet_notes[note_id])
    
    from ..utils.text_utils import unique_preserve_order
    return code, label, unique_preserve_order(notes)


def find_matrix_header_row(ws) -> Optional[int]:
    """Find the header row containing 'Description' column.
    
    Args:
        ws: Worksheet object
        
    Returns:
        Row number or None if not found
    """
    for r in range(1, min(ws.max_row, 10) + 1):
        values = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "Description" in values:
            return r
    return None


def parse_value_and_footnote_ids(raw: str) -> Tuple[str, List[str]]:
    """Parse a value and extract footnote IDs.
    
    Args:
        raw: Raw value string
        
    Returns:
        Tuple of (value, footnote_ids)
    """
    raw = normalize_text(raw)
    if not raw or raw == "--":
        return "", []
    
    # Try pattern: "text)digits"
    m = re.match(r"^(.*\))(\d+)$", raw)
    if m:
        return normalize_text(m.group(1)), re.findall(r"\d+", m.group(2))
    
    # Try pattern: "number.numberdigits"
    m = re.match(r"^([0-9]+\.[0-9]{2})(\d+)$", raw)
    if m:
        return normalize_text(m.group(1)), re.findall(r"\d+", m.group(2))
    
    # Try pattern: "text with letters digits"
    from ..utils.constants import TRAILING_DIGITS_RE
    m = TRAILING_DIGITS_RE.match(raw)
    if m and re.search(r"[A-Za-z]", m.group(1)):
        return normalize_text(m.group(1)), re.findall(r"\d+", m.group(2))
    
    return raw, []
