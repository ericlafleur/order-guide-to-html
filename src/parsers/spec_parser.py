"""Specification sheet parser."""

from typing import List

from ..models.base_models import SpecCell, SpecColumn
from ..utils.text_utils import normalize_text


def parse_spec_sheet(ws) -> List[SpecColumn]:
    """Parse a specifications and dimensions sheet.
    
    Args:
        ws: Worksheet object
        
    Returns:
        List of SpecColumn objects
    """
    section_markers = {"Specifications", "Capacities"}
    first_section_row: int | None = None
    
    for r in range(1, min(ws.max_row, 10) + 1):
        if normalize_text(ws.cell(r, 1).value) in section_markers:
            first_section_row = r
            break
    
    if first_section_row is None:
        return []

    header_row = first_section_row
    nonempty_on_section_row = sum(
        1 for c in range(1, ws.max_column + 1) 
        if normalize_text(ws.cell(first_section_row, c).value)
    )
    if nonempty_on_section_row <= 1 and first_section_row > 1:
        header_row = first_section_row - 1

    # Extract top label
    top_label = ""
    for r in range(1, header_row + 1):
        cell = normalize_text(ws.cell(r, 1).value)
        if cell and cell not in section_markers and "all dimensions" not in cell.lower():
            top_label = cell
            break

    # Parse columns
    columns: List[SpecColumn] = []
    for c in range(2, ws.max_column + 1):
        header = normalize_text(ws.cell(header_row, c).value)
        if not header:
            continue
        
        header_lines = [normalize_text(x) for x in header.split("\n") if normalize_text(x)]
        current_section = (
            normalize_text(ws.cell(first_section_row, 1).value) 
            if header_row == first_section_row else ""
        )
        
        cells: List[SpecCell] = []
        for r in range(header_row + 1, ws.max_row + 1):
            label = normalize_text(ws.cell(r, 1).value)
            row_values = [normalize_text(ws.cell(r, cc).value) for cc in range(1, ws.max_column + 1)]
            
            if label in section_markers and sum(1 for x in row_values[1:] if x) == 0:
                current_section = label
                continue
            
            value = normalize_text(ws.cell(r, c).value)
            if label and value and value != "--":
                cells.append(SpecCell(section=current_section or "Data", label=label, value=value))
        
        if cells:
            columns.append(
                SpecColumn(
                    sheet_name=ws.title,
                    top_label=top_label,
                    header=header,
                    header_lines=header_lines,
                    cells=cells,
                )
            )
    
    return columns
