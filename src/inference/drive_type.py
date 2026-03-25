"""Drive type inference."""

from typing import TYPE_CHECKING, List, Optional

from ..utils.constants import DRIVE_TYPE_PATTERNS
from ..utils.text_utils import normalize_text

if TYPE_CHECKING:
    from ..models.base_models import EngineAxleEntry, MatrixSheet, SpecColumn, TrimDef


def parse_drive_type_from_text(text: str) -> List[str]:
    """Extract all drive type tokens found in text (e.g. a spec column header).
    
    Args:
        text: Text to search for drive type patterns
        
    Returns:
        List of drive type labels found
    """
    found: List[str] = []
    for pattern, label in DRIVE_TYPE_PATTERNS:
        if pattern.search(text):
            found.append(label)
    return found


def infer_drive_types(
    engine_axle_entries: 'List[EngineAxleEntry]',
    spec_columns: 'List[SpecColumn]',
    matrix_sheets: 'List[MatrixSheet]',
    propulsion: str,
) -> List[str]:
    """Collect all drive types from engine axle model codes, spec column headers,
    and equipment descriptions.
    
    Args:
        engine_axle_entries: Engine/axle entries
        spec_columns: Spec columns
        matrix_sheets: Matrix sheets
        propulsion: Vehicle propulsion type
        
    Returns:
        Sorted list of drive types
    """
    drives: set = set()

    # Engine axle model codes: CC prefix = 2WD, CK prefix = 4WD
    for entry in engine_axle_entries:
        code = entry.model_code
        if code.startswith('CK'):
            drives.add('4WD')
        elif code.startswith('CC'):
            drives.add('2WD')

    # Spec column headers (encodes drive info for trucks: "2WD Short Bed Crew Cab")
    for col in spec_columns:
        for text in [col.top_label, col.header] + col.header_lines:
            for dt in parse_drive_type_from_text(text):
                drives.add(dt)

    # Equipment / feature descriptions mentioning drive type
    for sheet in matrix_sheets:
        for row in sheet.rows:
            desc = normalize_text(row.description_main or row.description_raw).lower()
            if 'all-wheel drive' in desc or ' awd' in desc:
                drives.add('AWD')
            elif 'front-wheel drive' in desc or ' fwd' in desc:
                drives.add('FWD')
            elif 'rear-wheel drive' in desc or ' rwd' in desc:
                drives.add('RWD')

    # EV fallback: if still nothing, assume AWD
    if not drives and propulsion == 'EV':
        drives.add('AWD')

    return sorted(drives) if drives else ['FWD']


def infer_trim_drive_type(
    spec_columns: 'List[SpecColumn]',
    trim: 'TrimDef',
    vehicle_drive_types: List[str],
) -> Optional[str]:
    """Best-effort per-trim drive type using fallback chain.
    
    1. Look for a spec column whose header contains this trim's name/code and
       mentions a drive type token explicitly.
    2. If the vehicle as a whole has exactly one drive type, inherit it.
    3. Otherwise return None (ambiguous).
    
    Args:
        spec_columns: Spec columns
        trim: Trim definition
        vehicle_drive_types: Vehicle-level drive types
        
    Returns:
        Drive type or None if ambiguous
    """
    candidate: set = set()
    for col in spec_columns:
        blob = ' '.join([col.top_label, col.header] + col.header_lines)
        if trim.name.lower() in blob.lower() or trim.code.lower() in blob.lower():
            for dt in parse_drive_type_from_text(blob):
                candidate.add(dt)
    
    if len(candidate) == 1:
        return candidate.pop()
    if len(vehicle_drive_types) == 1:
        return vehicle_drive_types[0]
    
    return None
