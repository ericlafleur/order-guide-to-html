"""Trailering specifications sheet parser."""

from typing import Dict, List, Tuple

from ..models.base_models import GCWRRecord, TraileringRecord
from ..utils.constants import FOOTNOTE_LINE_RE
from ..utils.text_utils import normalize_text, parse_footnote_map, unique_preserve_order
from .common import parse_value_and_footnote_ids


def parse_trailering_sheet(ws) -> Tuple[List[TraileringRecord], List[GCWRRecord]]:
    """Parse trailering specifications sheet.
    
    Args:
        ws: Worksheet object
        
    Returns:
        Tuple of (trailering_records, gcwr_records)
    """
    if not ws.title.startswith("Trailering Specs"):
        return [], []

    # Collect footnotes
    sheet_footnotes: Dict[str, str] = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            sheet_footnotes.update(parse_footnote_map(normalize_text(ws.cell(r, c).value)))

    rating_note = normalize_text(ws.cell(1, 1).value)
    rating_type = normalize_text(ws.cell(2, 1).value)

    # Parse engine column pairs (engine header spans 2 columns: axle ratio + max weight)
    engine_pairs: List[Tuple[str, int, int]] = []
    c = 2
    while c <= ws.max_column:
        engine = normalize_text(ws.cell(3, c).value)
        if engine:
            engine_pairs.append((engine, c, c + 1))
        c += 2

    # Parse trailering records
    trailering_records: List[TraileringRecord] = []
    current_model = ""
    gcwr_start_row = None
    
    for r in range(5, ws.max_row + 1):
        first = normalize_text(ws.cell(r, 1).value)
        if first.startswith("GCWR "):
            gcwr_start_row = r
            break
        if first and FOOTNOTE_LINE_RE.match(first):
            continue
        if first:
            current_model = first
        if not current_model:
            continue
        
        for engine, axle_col, weight_col in engine_pairs:
            raw_axle = normalize_text(ws.cell(r, axle_col).value)
            raw_weight = normalize_text(ws.cell(r, weight_col).value)
            
            if (not raw_axle or raw_axle == "--") and (not raw_weight or raw_weight == "--"):
                continue
            
            axle_ratio, axle_note_ids = parse_value_and_footnote_ids(raw_axle)
            max_weight, weight_note_ids = parse_value_and_footnote_ids(raw_weight)
            note_texts = [
                sheet_footnotes[nid] 
                for nid in axle_note_ids + weight_note_ids 
                if nid in sheet_footnotes
            ]
            
            trailering_records.append(
                TraileringRecord(
                    sheet_name=ws.title,
                    rating_type=rating_type,
                    note_text=rating_note,
                    model_code=current_model,
                    engine=engine,
                    axle_ratio=axle_ratio or raw_axle,
                    max_trailer_weight=max_weight or raw_weight,
                    footnotes=unique_preserve_order(note_texts),
                )
            )

    # Parse GCWR records
    gcwr_records: List[GCWRRecord] = []
    if gcwr_start_row is not None and gcwr_start_row + 3 <= ws.max_row:
        table_title = normalize_text(ws.cell(gcwr_start_row, 1).value)
        header_values = [
            normalize_text(ws.cell(gcwr_start_row + 2, c).value) 
            for c in range(2, ws.max_column + 1)
        ]
        
        for r in range(gcwr_start_row + 3, ws.max_row + 1):
            engine = normalize_text(ws.cell(r, 1).value)
            if not engine or FOOTNOTE_LINE_RE.match(engine):
                continue
            
            for idx, gcwr in enumerate(header_values, start=2):
                if not gcwr:
                    continue
                raw_axle = normalize_text(ws.cell(r, idx).value)
                if not raw_axle or raw_axle == "--":
                    continue
                
                axle_ratio, note_ids = parse_value_and_footnote_ids(raw_axle)
                gcwr_records.append(
                    GCWRRecord(
                        sheet_name=ws.title,
                        table_title=table_title,
                        engine=engine,
                        gcwr=gcwr,
                        axle_ratio=axle_ratio or raw_axle,
                        footnotes=unique_preserve_order(
                            sheet_footnotes[nid] for nid in note_ids if nid in sheet_footnotes
                        ),
                    )
                )

    return trailering_records, gcwr_records
