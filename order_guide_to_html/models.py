from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import re

from .utils import NON_ALNUM_RE, normalize_text


@dataclass
class TrimDef:
    name: str
    code: str
    raw_header: str
    model_code: str = ""
    family_label: str = ""

    @property
    def key(self) -> str:
        code_part = self.code if normalize_text(self.code) and normalize_text(self.code) != normalize_text(self.name) else ''
        basis = ' | '.join(
            part for part in [self.family_label, self.name or self.raw_header, code_part] if normalize_text(part)
        )
        return NON_ALNUM_RE.sub(' ', basis).strip().lower()

    @property
    def label(self) -> str:
        if self.code and self.name and self.code.lower() not in self.name.lower().split():
            return f"{self.name} ({self.code})"
        return self.name or self.code

@dataclass
class MatrixRow:
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
        basis = " | ".join(
            normalize_text(x)
            for x in [self.option_code or "", self.ref_code or "", self.description_main or self.description_raw]
            if normalize_text(x)
        )
        return NON_ALNUM_RE.sub(" ", basis).strip().lower()

@dataclass
class MatrixSheet:
    name: str
    legend_text: str
    trim_defs: List[TrimDef]
    footnotes: Dict[str, str]
    rows: List[MatrixRow]

@dataclass
class ColorInteriorRow:
    decor_level: str
    seat_type: str
    seat_code: str
    seat_trim: str
    colors: Dict[str, str]

@dataclass
class ColorExteriorRow:
    title: str
    color_code: str
    touch_up_paint_number: str
    colors: Dict[str, str]

@dataclass
class ColorSheet:
    name: str
    heading_text: str
    footnotes: Dict[str, str]
    bullet_notes: List[str]
    interior_rows: List[ColorInteriorRow]
    exterior_rows: List[ColorExteriorRow]

@dataclass
class SpecCell:
    section: str
    label: str
    value: str

@dataclass
class SpecColumn:
    sheet_name: str
    top_label: str
    header: str
    header_lines: List[str]
    cells: List[SpecCell]

@dataclass
class EngineAxleItem:
    category: str
    name: str
    raw_status: str
    status_code: str
    status_label: str
    notes: List[str]

@dataclass
class EngineAxleEntry:
    sheet_name: str
    top_label: str
    model_code: str
    engine: str
    items: List[EngineAxleItem]

@dataclass
class TraileringRecord:
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
    sheet_name: str
    table_title: str
    engine: str
    gcwr: str
    axle_ratio: str
    footnotes: List[str]

@dataclass
class WorkbookData:
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
    language: str

@dataclass
class ModelFeatureAggregate:
    title: str
    description: str
    orderable_code: str
    reference_code: str
    source_contexts: List[str] = field(default_factory=list)
    availability_contexts: OrderedDict[Tuple[Tuple[Tuple[str, str, Tuple[str, ...]], ...]], List[str]] = field(default_factory=OrderedDict)
    notes: List[str] = field(default_factory=list)
    referenced_codes: List[Tuple[str, str]] = field(default_factory=list)

@dataclass
class TrimFeatureAggregate:
    title: str
    description: str
    orderable_code: str
    reference_code: str
    availability_contexts: OrderedDict[Tuple[str, str], List[str]] = field(default_factory=OrderedDict)
    source_contexts: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)
    referenced_codes: List[Tuple[str, str]] = field(default_factory=list)

@dataclass
class OutputFileRecord:
    objecttype: str
    type: str
    path: Path
    metadata: Dict[str, object]

@dataclass
class BoundRecord:
    record: OutputFileRecord
    collection: Optional[Path]
    parent: Optional[Path]
    parent_vehicle: Optional[Path]
    parent_trims: List[Path]

@dataclass
class SpecGroupDoc:
    top_label: str
    header: str
    header_lines: List[str]
    columns: List[SpecColumn]

@dataclass
class PowertrainTraileringGroup:
    model_code: str
    top_labels: List[str]
    engine_entries: List[EngineAxleEntry]
    trailering_records: List[TraileringRecord]
