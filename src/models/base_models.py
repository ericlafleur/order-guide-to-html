"""Base data models for parsed workbook data."""

import re
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from ..utils.constants import NON_ALNUM_RE
from ..utils.text_utils import normalize_text


@dataclass
class TrimDef:
    """Definition of a vehicle trim level."""
    
    name: str
    code: str
    raw_header: str

    @property
    def key(self) -> str:
        """Unique key for this trim."""
        return self.code or self.name

    @property
    def label(self) -> str:
        """Display label for this trim."""
        if self.code and self.name:
            return f"{self.name} ({self.code})"
        return self.name or self.code


@dataclass
class MatrixRow:
    """A row from an equipment/features matrix sheet."""
    
    sheet_name: str
    row_group: Optional[str]
    option_code: Optional[str]
    ref_code: Optional[str]
    aux_meta: List[str]
    description_raw: str
    description_main: str
    inline_footnotes: Dict[str, str] = field(default_factory=dict)
    bullet_notes: List[str] = field(default_factory=list)
    status_by_trim: Dict[str, str] = field(default_factory=dict)

    @property
    def label(self) -> str:
        """Short display label for this feature."""
        text = self.description_main or self.description_raw
        text = normalize_text(text).split("\n")[0]
        text = re.sub(r"^NEW!\s+", "", text, flags=re.I)
        short = text
        if ", includes " in text.lower():
            short = text.split(",", 1)[0]
        elif ". " in text and len(text.split(". ", 1)[0]) >= 12:
            short = text.split(". ", 1)[0]
        if len(short) > 140:
            short = short[:137].rstrip() + "..."
        return normalize_text(short)

    @property
    def identity_key(self) -> str:
        """Unique identity key for deduplication."""
        basis = " | ".join(
            normalize_text(x)
            for x in [self.option_code or "", self.ref_code or "", self.description_main or self.description_raw]
            if normalize_text(x)
        )
        return NON_ALNUM_RE.sub(" ", basis).strip().lower()


@dataclass
class MatrixSheet:
    """Parsed equipment/features matrix sheet."""
    
    name: str
    legend_text: str
    trim_defs: List[TrimDef]
    footnotes: Dict[str, str]
    rows: List[MatrixRow]


@dataclass
class ColorInteriorRow:
    """Interior color and trim row."""
    
    decor_level: str
    seat_type: str
    seat_code: str
    seat_trim: str
    colors: Dict[str, str]


@dataclass
class ColorExteriorRow:
    """Exterior color row."""
    
    title: str
    color_code: str
    touch_up_paint_number: str
    colors: Dict[str, str]


@dataclass
class ColorSheet:
    """Parsed color and trim sheet."""
    
    name: str
    heading_text: str
    footnotes: Dict[str, str]
    bullet_notes: List[str]
    interior_rows: List[ColorInteriorRow]
    exterior_rows: List[ColorExteriorRow]


@dataclass
class SpecCell:
    """A specification cell value."""
    
    section: str
    label: str
    value: str


@dataclass
class SpecColumn:
    """A column from specifications sheet."""
    
    sheet_name: str
    top_label: str
    header: str
    header_lines: List[str]
    cells: List[SpecCell]


@dataclass
class EngineAxleItem:
    """An engine/axle availability item."""
    
    category: str
    name: str
    raw_status: str
    status_code: str
    status_label: str
    notes: List[str]


@dataclass
class EngineAxleEntry:
    """Engine and axle entry for a model code."""
    
    sheet_name: str
    top_label: str
    model_code: str
    engine: str
    items: List[EngineAxleItem]


@dataclass
class TraileringRecord:
    """Trailering specification record."""
    
    sheet_name: str
    rating_type: str
    note_text: str
    model_code: str
    engine: str
    axle_ratio: str
    max_trailer_weight: str
    footnotes: List[str]


@dataclass
class GCWRRecord:
    """Gross Combined Weight Rating record."""
    
    sheet_name: str
    table_title: str
    engine: str
    gcwr: str
    axle_ratio: str
    footnotes: List[str]


@dataclass
class WorkbookData:
    """Complete parsed workbook data."""
    
    path: Path
    year: str
    make: str
    model: str
    vehicle_name: str
    trim_defs: List[TrimDef]
    matrix_sheets: List[MatrixSheet]
    color_sheets: List[ColorSheet]
    spec_columns: List[SpecColumn]
    engine_axle_entries: List[EngineAxleEntry]
    trailering_records: List[TraileringRecord]
    gcwr_records: List[GCWRRecord]
    glossary: OrderedDict[str, str]
    sheet_names: List[str]
    # Inferred vehicle attributes
    propulsion: str = field(default='ICE')
    vehicle_type: str = field(default='suv')
    drive_types: List[str] = field(default_factory=list)
