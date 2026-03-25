"""Code reference utilities."""

from typing import Dict, List, Tuple

from ..utils.constants import CODE_IN_PARENS_RE


def referenced_codes_for_text(text: str, glossary: Dict[str, str]) -> List[Tuple[str, str]]:
    """Extract option codes referenced in text that match glossary entries.
    
    Args:
        text: Text to search for codes
        glossary: Glossary mapping codes to descriptions
        
    Returns:
        List of (code, description) tuples
    """
    codes = []
    for code in CODE_IN_PARENS_RE.findall(text):
        if code in glossary:
            codes.append((code, glossary[code]))
    
    # Deduplicate while preserving order
    from collections import OrderedDict
    return list(OrderedDict(((code, desc), None) for code, desc in codes).keys())
