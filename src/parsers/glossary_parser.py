"""Glossary sheet parser."""

from collections import OrderedDict

from ..utils.text_utils import normalize_text


def parse_glossary_sheet(ws) -> OrderedDict[str, str]:
    """Parse glossary/option codes sheet.
    
    Args:
        ws: Worksheet object
        
    Returns:
        OrderedDict mapping option codes to descriptions
    """
    glossary: OrderedDict[str, str] = OrderedDict()
    
    first = normalize_text(ws.cell(1, 1).value)
    second = normalize_text(ws.cell(1, 2).value)
    
    if first != "Option Code" or second != "Description":
        return glossary
    
    for r in range(2, ws.max_row + 1):
        code = normalize_text(ws.cell(r, 1).value)
        description = normalize_text(ws.cell(r, 2).value)
        if code and description:
            glossary[code] = description
    
    return glossary
