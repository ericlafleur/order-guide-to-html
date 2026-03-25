"""Engine and axle sheet parser."""

from typing import Dict, List

from ..models.base_models import EngineAxleEntry, EngineAxleItem
from ..utils.constants import FOOTNOTE_LINE_RE
from ..utils.text_utils import normalize_text, parse_footnote_map
from .common import parse_status_value


def parse_engine_axles_sheet(ws) -> List[EngineAxleEntry]:
    """Parse engine and axle availability sheet.
    
    Args:
        ws: Worksheet object
        
    Returns:
        List of EngineAxleEntry objects
    """
    if not ws.title.startswith("Engine Axles"):
        return []
    
    top_label = normalize_text(ws.cell(1, 1).value)
    
    # Collect footnotes
    footnotes: Dict[str, str] = {}
    for r in range(1, ws.max_row + 1):
        first = normalize_text(ws.cell(r, 1).value)
        if first:
            footnotes.update(parse_footnote_map(first))

    # Parse section headers (row 3)
    section_headers: Dict[int, str] = {}
    current_section = ""
    for c in range(3, ws.max_column + 1):
        value = normalize_text(ws.cell(3, c).value)
        if value:
            current_section = value
        section_headers[c] = current_section

    # Parse engine/axle entries
    entries: List[EngineAxleEntry] = []
    current_model = ""
    
    for r in range(5, ws.max_row + 1):
        model = normalize_text(ws.cell(r, 1).value)
        engine = normalize_text(ws.cell(r, 2).value)
        row_values = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        
        if not any(row_values):
            continue
        if model and FOOTNOTE_LINE_RE.match(model):
            continue
        if model:
            current_model = model
        if not engine:
            continue
        
        items: List[EngineAxleItem] = []
        for c in range(3, ws.max_column + 1):
            raw_status = normalize_text(ws.cell(r, c).value)
            if not raw_status or raw_status == "--":
                continue
            
            code, label, notes = parse_status_value(raw_status, {}, footnotes)
            items.append(
                EngineAxleItem(
                    category=section_headers.get(c, ""),
                    name=normalize_text(ws.cell(4, c).value),
                    raw_status=raw_status,
                    status_code=code,
                    status_label=label,
                    notes=notes,
                )
            )
        
        if items:
            entries.append(
                EngineAxleEntry(
                    sheet_name=ws.title,
                    top_label=top_label,
                    model_code=current_model,
                    engine=engine,
                    items=items,
                )
            )
    
    return entries
