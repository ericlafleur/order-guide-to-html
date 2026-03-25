"""Metadata extraction for manifest files."""

from typing import TYPE_CHECKING, Dict, List, Optional, Sequence

from ..aggregators.feature_aggregator import aggregate_model_features, aggregate_trim_features
from ..inference import infer_trim_drive_type
from ..parsers.common import parse_status_value
from ..utils.constants import MANIFEST_BODY_STYLE_TOKENS, MANIFEST_DRIVE_TOKENS, MANIFEST_STANDARDISH_CODES
from ..utils.text_utils import normalize_text, unique_preserve_order

if TYPE_CHECKING:
    from ..models.base_models import SpecColumn, TrimDef, WorkbookData


def first_unique(values: Sequence[str]) -> Optional[str]:
    """Get first value if unique.
    
    Args:
        values: Sequence of values
        
    Returns:
        Value if unique, else None
    """
    cleaned = unique_preserve_order(normalize_text(v) for v in values if normalize_text(v))
    if len(cleaned) == 1:
        return cleaned[0]
    return None


def standardish_trim_descriptions(data: 'WorkbookData', trim: 'TrimDef') -> List[str]:
    """Get standard/included descriptions for trim.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        List of descriptions
    """
    values: List[str] = []
    for sheet in data.matrix_sheets:
        for row in sheet.rows:
            raw = normalize_text(row.status_by_trim.get(trim.key))
            if not raw:
                continue
            status_code, _status_label, _notes = parse_status_value(raw, row.inline_footnotes, sheet.footnotes)
            if status_code not in MANIFEST_STANDARDISH_CODES:
                continue
            desc = normalize_text(row.description_main or row.description_raw)
            if desc:
                values.append(desc)
    return unique_preserve_order(values)


def model_descriptions_standard_for_all_trims(data: 'WorkbookData') -> List[str]:
    """Get descriptions standard across all trims.
    
    Args:
        data: Workbook data
        
    Returns:
        List of descriptions
    """
    values: List[str] = []
    trim_defs = list(data.trim_defs)
    if not trim_defs:
        return values
    
    for sheet in data.matrix_sheets:
        for row in sheet.rows:
            codes: List[str] = []
            for trim in trim_defs:
                raw = normalize_text(row.status_by_trim.get(trim.key))
                if not raw:
                    codes = []
                    break
                status_code, _status_label, _notes = parse_status_value(raw, row.inline_footnotes, sheet.footnotes)
                codes.append(status_code)
            if not codes or any(code not in MANIFEST_STANDARDISH_CODES for code in codes):
                continue
            desc = normalize_text(row.description_main or row.description_raw)
            if desc:
                values.append(desc)
    return unique_preserve_order(values)


def pick_engine_description(values: Sequence[str]) -> Optional[str]:
    """Pick best engine description.
    
    Args:
        values: Candidate values
        
    Returns:
        Best match or None
    """
    preferred = [
        v for v in values
        if normalize_text(v).lower().startswith('electric drive unit')
    ]
    choice = first_unique(preferred)
    if choice:
        return choice
    
    preferred = [
        v for v in values
        if normalize_text(v).lower().startswith('engine,') and normalize_text(v).lower() != 'engine, none'
    ]
    choice = first_unique(preferred)
    if choice:
        return choice
    
    return None


def pick_fuel_description(values: Sequence[str]) -> Optional[str]:
    """Pick fuel description.
    
    Args:
        values: Candidate values
        
    Returns:
        Best match or None
    """
    return first_unique(v for v in values if normalize_text(v).lower().startswith('fuel,'))


def pick_drivetrain_description(values: Sequence[str]) -> Optional[str]:
    """Pick drivetrain description.
    
    Args:
        values: Candidate values
        
    Returns:
        Best match or None
    """
    direct = [v for v in values if 'wheel drive' in normalize_text(v).lower()]
    choice = first_unique(direct)
    if choice:
        return choice
    
    propulsion = [
        v for v in values 
        if normalize_text(v).lower().startswith('propulsion,') 
        and 'fwd' not in normalize_text(v).lower() 
        and 'awd' not in normalize_text(v).lower() 
        and 'rwd' not in normalize_text(v).lower() 
        and '4wd' not in normalize_text(v).lower() 
        and '2wd' not in normalize_text(v).lower()
    ]
    choice = first_unique(propulsion)
    if choice:
        return choice
    
    tokenized = [v for v in values if normalize_text(v).lower().startswith('propulsion,')]
    return first_unique(tokenized)


def trim_direct_spec_columns(data: 'WorkbookData', trim: 'TrimDef') -> List['SpecColumn']:
    """Get spec columns directly related to trim.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        List of matching spec columns
    """
    from ..renderers.common import column_matches_trim
    return [column for column in data.spec_columns if column_matches_trim(column, trim)]


def extract_trim_seating(data: 'WorkbookData', trim: 'TrimDef') -> Optional[str]:
    """Extract seating capacity for trim.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        Seating capacity or None
    """
    cols = trim_direct_spec_columns(data, trim)
    values = [
        cell.value
        for column in cols
        for cell in column.cells
        if normalize_text(cell.label).lower().startswith('seating capacity')
    ]
    return first_unique(values)


def extract_model_seating(data: 'WorkbookData') -> Optional[str]:
    """Extract seating capacity for model.
    
    Args:
        data: Workbook data
        
    Returns:
        Seating capacity or None
    """
    values = [
        cell.value
        for column in data.spec_columns
        for cell in column.cells
        if normalize_text(cell.label).lower().startswith('seating capacity')
    ]
    return first_unique(values)


def extract_trim_body_style(data: 'WorkbookData', trim: 'TrimDef') -> Optional[str]:
    """Extract body style for trim.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        Body style or None
    """
    cols = trim_direct_spec_columns(data, trim)
    top_labels = [column.top_label for column in cols if normalize_text(column.top_label)]
    choice = first_unique(top_labels)
    if choice:
        return choice
    
    header_hits: List[str] = []
    for column in cols:
        for token in MANIFEST_BODY_STYLE_TOKENS:
            if token.lower() in ' '.join([column.top_label, column.header] + column.header_lines).lower():
                header_hits.append(token)
    return first_unique(header_hits)


def extract_model_body_style(data: 'WorkbookData') -> Optional[str]:
    """Extract body style for model.
    
    Args:
        data: Workbook data
        
    Returns:
        Body style or None
    """
    top_labels = [column.top_label for column in data.spec_columns if normalize_text(column.top_label)]
    return first_unique(top_labels)


def extract_trim_drive_token_from_headers(data: 'WorkbookData', trim: 'TrimDef') -> Optional[str]:
    """Extract drive type token from headers for trim.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        Drive type token or None
    """
    cols = trim_direct_spec_columns(data, trim)
    tokens: List[str] = []
    for column in cols:
        blob = '\n'.join([column.top_label, column.header] + column.header_lines)
        for token in MANIFEST_DRIVE_TOKENS:
            if token.lower() in blob.lower():
                tokens.append(token)
    return first_unique(tokens)


def extract_manifest_metadata_for_model(data: 'WorkbookData') -> Dict[str, object]:
    """Extract manifest metadata for model.
    
    Args:
        data: Workbook data
        
    Returns:
        Metadata dictionary
    """
    descs = model_descriptions_standard_for_all_trims(data)
    metadata: Dict[str, object] = {}
    
    seating = extract_model_seating(data)
    if seating:
        metadata['seating'] = seating
    
    body_style = extract_model_body_style(data)
    if body_style:
        metadata['body_style'] = body_style
    
    engine = pick_engine_description(descs)
    if engine:
        metadata['engine'] = engine
    
    fuel = pick_fuel_description(descs)
    if fuel:
        metadata['fuel_type'] = fuel
    
    drivetrain = pick_drivetrain_description(descs)
    if drivetrain:
        metadata['drivetrain'] = drivetrain
    
    if data.propulsion:
        metadata['propulsion'] = data.propulsion
    if data.vehicle_type:
        metadata['vehicle_type'] = data.vehicle_type
    if data.drive_types:
        metadata['drive_types'] = data.drive_types
    
    return metadata


def extract_manifest_metadata_for_trim(data: 'WorkbookData', trim: 'TrimDef') -> Dict[str, object]:
    """Extract manifest metadata for trim.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        Metadata dictionary
    """
    descs = standardish_trim_descriptions(data, trim)
    metadata: Dict[str, object] = {}
    
    trim_name = normalize_text(trim.name)
    if trim_name:
        metadata['name'] = trim_name
    
    title = normalize_text(trim.raw_header)
    if title:
        metadata['title'] = title
    
    seating = extract_trim_seating(data, trim)
    if seating:
        metadata['seating'] = seating
    
    body_style = extract_trim_body_style(data, trim)
    if body_style:
        metadata['body_style'] = body_style
    
    engine = pick_engine_description(descs)
    if engine:
        metadata['engine'] = engine
    
    fuel = pick_fuel_description(descs)
    if fuel:
        metadata['fuel_type'] = fuel
    
    drivetrain = pick_drivetrain_description(descs)
    if not drivetrain:
        drivetrain = extract_trim_drive_token_from_headers(data, trim)
    if drivetrain:
        metadata['drivetrain'] = drivetrain
    
    if data.propulsion:
        metadata['propulsion'] = data.propulsion
    if data.vehicle_type:
        metadata['vehicle_type'] = data.vehicle_type
    
    trim_drive = infer_trim_drive_type(data.spec_columns, trim, data.drive_types)
    if trim_drive:
        metadata['drive_type'] = trim_drive
    elif data.drive_types:
        metadata['drive_types'] = data.drive_types
    
    return metadata
