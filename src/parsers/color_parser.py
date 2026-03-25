"""Color and trim sheet parser."""

from typing import List

from ..models.base_models import ColorExteriorRow, ColorInteriorRow, ColorSheet
from ..utils.constants import FOOTNOTE_LINE_RE
from ..utils.text_utils import normalize_text, parse_footnote_map, unique_preserve_order


def parse_color_sheet(ws) -> ColorSheet:
    """Parse a color and trim sheet.
    
    Args:
        ws: Worksheet object
        
    Returns:
        ColorSheet with interior and exterior color data
    """
    footnotes: dict = {}
    bullet_notes: List[str] = []
    heading_lines: List[str] = []
    interior_rows: List[ColorInteriorRow] = []
    exterior_rows: List[ColorExteriorRow] = []

    # First pass: collect headings and footnotes
    for r in range(1, ws.max_row + 1):
        row_values = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        first = row_values[0] if row_values else ""
        
        if r <= 3 and any(row_values):
            heading_lines.extend([v for v in row_values if v])
        
        for value in row_values:
            if value:
                footnotes.update(parse_footnote_map(value))
                for line in normalize_text(value).split("\n"):
                    if line.startswith("•"):
                        bullet_notes.append(normalize_text(line.lstrip("•").strip()))

        # Parse interior colors section
        if first == "Decor Level":
            color_headers = [
                normalize_text(ws.cell(r, c).value).split("\n")[0] 
                for c in range(5, ws.max_column + 1)
            ]
            rr = r + 1
            while rr <= ws.max_row:
                values = [normalize_text(ws.cell(rr, c).value) for c in range(1, ws.max_column + 1)]
                if values[0] == "Exterior Solid Paint":
                    break
                if not any(values):
                    rr += 1
                    continue
                if values[0] and values[0] != "Decor Level":
                    colors = {
                        color_headers[i]: values[4 + i]
                        for i in range(len(color_headers))
                        if color_headers[i]
                    }
                    interior_rows.append(
                        ColorInteriorRow(
                            decor_level=values[0],
                            seat_type=values[1],
                            seat_code=values[2],
                            seat_trim=values[3],
                            colors=colors,
                        )
                    )
                rr += 1

        # Parse exterior colors section
        if first == "Exterior Solid Paint":
            color_headers = [
                normalize_text(ws.cell(r, c).value).split("\n")[0] 
                for c in range(5, ws.max_column + 1)
            ]
            rr = r + 1
            while rr <= ws.max_row:
                values = [normalize_text(ws.cell(rr, c).value) for c in range(1, ws.max_column + 1)]
                if not any(values):
                    rr += 1
                    continue
                if FOOTNOTE_LINE_RE.match(values[0]) or values[0].startswith("•"):
                    break
                if values[0] and values[0] != "Exterior Solid Paint":
                    colors = {
                        color_headers[i]: values[4 + i]
                        for i in range(len(color_headers))
                        if color_headers[i]
                    }
                    exterior_rows.append(
                        ColorExteriorRow(
                            title=values[0],
                            color_code=values[2],
                            touch_up_paint_number=values[3],
                            colors=colors,
                        )
                    )
                rr += 1

    return ColorSheet(
        name=ws.title,
        heading_text=" | ".join(unique_preserve_order(heading_lines)),
        footnotes=footnotes,
        bullet_notes=unique_preserve_order(bullet_notes),
        interior_rows=interior_rows,
        exterior_rows=exterior_rows,
    )
