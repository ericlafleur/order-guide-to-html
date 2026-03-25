"""Matrix sheet parser."""

from typing import List, Optional

from ..models.base_models import MatrixRow, MatrixSheet, TrimDef
from ..utils.text_utils import (
    normalize_text,
    parse_footnote_map,
    split_main_notes_and_bullets,
    unique_preserve_order,
)
from .common import find_matrix_header_row, parse_trim_header


def parse_matrix_sheet(ws, trim_defs: Optional[List[TrimDef]] = None) -> Optional[MatrixSheet]:
    """Parse an equipment/features matrix sheet.
    
    Args:
        ws: Worksheet object
        trim_defs: Optional pre-parsed trim definitions
        
    Returns:
        MatrixSheet or None if not a matrix sheet
    """
    header_row = find_matrix_header_row(ws)
    if header_row is None:
        return None

    headers = [normalize_text(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
    description_col = headers.index("Description") + 1

    # Parse legend and footnotes from rows above header
    legend_text_parts: List[str] = []
    sheet_footnotes: dict = {}
    for r in range(1, header_row):
        row_texts = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        nonempty = [t for t in row_texts if t]
        if nonempty:
            legend_text_parts.extend(nonempty)
        for item in nonempty:
            sheet_footnotes.update(parse_footnote_map(item))

    # Parse trim definitions from header row
    inferred_trim_defs: List[TrimDef] = []
    c = description_col + 1
    while c <= ws.max_column:
        raw = normalize_text(ws.cell(header_row, c).value)
        if not raw:
            break
        parsed = parse_trim_header(raw)
        if parsed:
            inferred_trim_defs.append(parsed)
        c += 1
    
    active_trim_defs = trim_defs or inferred_trim_defs
    trim_cols = list(range(description_col + 1, description_col + 1 + len(active_trim_defs)))

    # Parse data rows
    rows: List[MatrixRow] = []
    current_group: Optional[str] = None
    
    for r in range(header_row + 1, ws.max_row + 1):
        meta = [normalize_text(ws.cell(r, c).value) for c in range(1, description_col + 1)]
        status_values = [normalize_text(ws.cell(r, c).value) for c in trim_cols]
        
        if not any(meta) and not any(status_values):
            continue

        # Check if this is a group header or footnote row
        if not any(status_values):
            nonempty = [m for m in meta if m]
            for item in nonempty:
                sheet_footnotes.update(parse_footnote_map(item))
            
            # Check if it's a group header (not a footnote)
            if nonempty and not any(
                any(line for line in item.split("\n") if parse_footnote_map(line))
                for item in nonempty
            ):
                current_group = nonempty[0]
            continue

        description_raw = meta[-1]
        if not description_raw:
            continue

        option_code = meta[0] or None
        ref_code = meta[1] or None if len(meta) > 1 else None
        aux_meta = [x for x in meta[2:-1] if x]
        main_text, inline_notes, bullet_notes = split_main_notes_and_bullets(description_raw)

        row = MatrixRow(
            sheet_name=ws.title,
            row_group=current_group,
            option_code=option_code,
            ref_code=ref_code,
            aux_meta=aux_meta,
            description_raw=description_raw,
            description_main=main_text or description_raw,
            inline_footnotes=inline_notes,
            bullet_notes=bullet_notes,
            status_by_trim={
                active_trim_defs[idx].key: status_values[idx]
                for idx in range(min(len(active_trim_defs), len(status_values)))
                if normalize_text(status_values[idx])
            },
        )
        rows.append(row)

    legend_text = "\n".join(unique_preserve_order(legend_text_parts))
    return MatrixSheet(
        name=ws.title,
        legend_text=legend_text,
        trim_defs=active_trim_defs,
        footnotes=sheet_footnotes,
        rows=rows,
    )
