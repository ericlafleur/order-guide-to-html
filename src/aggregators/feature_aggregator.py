"""Feature aggregation logic."""

from collections import OrderedDict
from typing import TYPE_CHECKING, List, Sequence, Tuple

from ..models.aggregate_models import ModelFeatureAggregate, TrimFeatureAggregate
from ..parsers.common import parse_status_value
from ..utils.constants import STATUS_PRIORITY
from ..utils.text_utils import normalize_text, unique_preserve_order
from .code_reference import referenced_codes_for_text

if TYPE_CHECKING:
    from ..models.base_models import MatrixRow, MatrixSheet, TrimDef, WorkbookData


def feature_title(label: str, orderable_code: str = '', reference_code: str = '') -> str:
    """Generate feature title with codes.
    
    Args:
        label: Feature label
        orderable_code: Orderable option code
        reference_code: Reference code
        
    Returns:
        Formatted title string
    """
    prefix = orderable_code or reference_code
    label = normalize_text(label)
    if prefix:
        return f'{prefix} | {label}'
    return label


def source_context(sheet_name: str, row_group: str | None = None) -> str:
    """Generate source context string.
    
    Args:
        sheet_name: Sheet name
        row_group: Optional row group
        
    Returns:
        Context string
    """
    sheet_name = normalize_text(sheet_name)
    row_group = normalize_text(row_group)
    if row_group and row_group.lower() != sheet_name.lower():
        return f'{sheet_name} | {row_group}'
    return sheet_name


def collect_row_note_texts(row: 'MatrixRow', sheet: 'MatrixSheet') -> List[str]:
    """Collect all note texts from a matrix row.
    
    Args:
        row: Matrix row
        sheet: Matrix sheet containing the row
        
    Returns:
        List of note texts
    """
    notes: List[str] = []
    
    for note_text in row.inline_footnotes.values():
        notes.append(note_text)
    
    for raw in row.status_by_trim.values():
        raw = normalize_text(raw)
        if not raw:
            continue
        import re
        m = re.match(r'^(--|[A-Z]+|[■□*]+)(.*)$', raw)
        suffix = m.group(2) if m else ''
        for note_id in re.findall(r'\d+', suffix):
            if note_id in row.inline_footnotes:
                notes.append(row.inline_footnotes[note_id])
            elif note_id in sheet.footnotes:
                notes.append(sheet.footnotes[note_id])
    
    notes.extend(row.bullet_notes)
    return unique_preserve_order([normalize_text(n) for n in notes if normalize_text(n)])


def summarize_model_status_groups(
    row: 'MatrixRow', 
    trim_defs: Sequence['TrimDef'], 
    sheet: 'MatrixSheet'
) -> Tuple[Tuple[str, str, Tuple[str, ...]], ...]:
    """Summarize status groups across trims for model-level view.
    
    Args:
        row: Matrix row
        trim_defs: Trim definitions
        sheet: Matrix sheet
        
    Returns:
        Tuple of (raw_status, label, trim_names)
    """
    groups: 'OrderedDict[Tuple[str, str], List[str]]' = OrderedDict()
    for trim in trim_defs:
        raw = normalize_text(row.status_by_trim.get(trim.key))
        if not raw:
            continue
        _code, label, _notes = parse_status_value(raw, row.inline_footnotes, sheet.footnotes)
        groups.setdefault((raw, label), []).append(trim.name or trim.code)
    return tuple((raw, label, tuple(names)) for (raw, label), names in groups.items())


def aggregate_model_features(data: 'WorkbookData') -> List[ModelFeatureAggregate]:
    """Aggregate features across all trims for model-level view.
    
    Args:
        data: Workbook data
        
    Returns:
        List of aggregated model features
    """
    groups: 'OrderedDict[str, ModelFeatureAggregate]' = OrderedDict()
    
    for sheet in data.matrix_sheets:
        for row in sheet.rows:
            title = feature_title(row.label, row.option_code or '', row.ref_code or '')
            agg = groups.get(row.identity_key)
            
            if agg is None:
                agg = ModelFeatureAggregate(
                    title=title,
                    description=normalize_text(row.description_main or row.description_raw),
                    orderable_code=normalize_text(row.option_code),
                    reference_code=normalize_text(row.ref_code),
                )
                groups[row.identity_key] = agg
            else:
                candidate_desc = normalize_text(row.description_main or row.description_raw)
                if len(candidate_desc) > len(agg.description):
                    agg.description = candidate_desc
                if not agg.orderable_code and row.option_code:
                    agg.orderable_code = normalize_text(row.option_code)
                if not agg.reference_code and row.ref_code:
                    agg.reference_code = normalize_text(row.ref_code)
            
            ctx = source_context(sheet.name, row.row_group)
            agg.source_contexts = unique_preserve_order(agg.source_contexts + [ctx])
            
            signature = summarize_model_status_groups(row, sheet.trim_defs, sheet)
            agg.availability_contexts.setdefault(signature, [])
            agg.availability_contexts[signature] = unique_preserve_order(
                agg.availability_contexts[signature] + [ctx]
            )
            
            agg.notes = unique_preserve_order(agg.notes + collect_row_note_texts(row, sheet))
            agg.referenced_codes = list(
                OrderedDict(
                    ((code, desc), None) 
                    for code, desc in (agg.referenced_codes + referenced_codes_for_text(row.description_raw, data.glossary))
                ).keys()
            )
    
    return list(groups.values())


def aggregate_trim_features(data: 'WorkbookData', trim: 'TrimDef') -> List[TrimFeatureAggregate]:
    """Aggregate features for a specific trim.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        List of aggregated trim features
    """
    groups: 'OrderedDict[str, TrimFeatureAggregate]' = OrderedDict()
    
    for sheet in data.matrix_sheets:
        for row in sheet.rows:
            raw = normalize_text(row.status_by_trim.get(trim.key))
            if not raw:
                continue
            
            code, label, _notes = parse_status_value(raw, row.inline_footnotes, sheet.footnotes)
            # Skip features that are explicitly not available on this trim
            if code == '--':
                continue
            
            title = feature_title(row.label, row.option_code or '', row.ref_code or '')
            agg = groups.get(row.identity_key)
            
            if agg is None:
                agg = TrimFeatureAggregate(
                    title=title,
                    description=normalize_text(row.description_main or row.description_raw),
                    orderable_code=normalize_text(row.option_code),
                    reference_code=normalize_text(row.ref_code),
                )
                groups[row.identity_key] = agg
            else:
                candidate_desc = normalize_text(row.description_main or row.description_raw)
                if len(candidate_desc) > len(agg.description):
                    agg.description = candidate_desc
                if not agg.orderable_code and row.option_code:
                    agg.orderable_code = normalize_text(row.option_code)
                if not agg.reference_code and row.ref_code:
                    agg.reference_code = normalize_text(row.ref_code)
            
            ctx = source_context(sheet.name, row.row_group)
            agg.source_contexts = unique_preserve_order(agg.source_contexts + [ctx])
            
            agg.availability_contexts.setdefault((raw, label), [])
            agg.availability_contexts[(raw, label)] = unique_preserve_order(
                agg.availability_contexts[(raw, label)] + [ctx]
            )
            
            agg.notes = unique_preserve_order(agg.notes + collect_row_note_texts(row, sheet))
            agg.referenced_codes = list(
                OrderedDict(
                    ((code, desc), None) 
                    for code, desc in (agg.referenced_codes + referenced_codes_for_text(row.description_raw, data.glossary))
                ).keys()
            )
    
    return sorted(groups.values(), key=sort_trim_feature)


def sort_trim_feature(agg: TrimFeatureAggregate) -> Tuple[int, str]:
    """Sort key for trim features.
    
    Args:
        agg: Trim feature aggregate
        
    Returns:
        Sort key tuple
    """
    first_key = next(iter(agg.availability_contexts.keys()), ('', ''))
    raw, label = first_key
    return (STATUS_PRIORITY.get(label, 99), normalize_text(agg.title).lower(), raw)


def _infer_feature_category(*texts: str) -> str:
    """Infer feature category from text content.
    
    Args:
        *texts: Variable number of text strings to analyze
        
    Returns:
        Category string
    """
    blob = ' '.join(normalize_text(text).lower() for text in texts if normalize_text(text))
    if not blob:
        return 'Other guide content'

    if 'colour and trim' in blob or 'color and trim' in blob or ' paint' in blob or blob.startswith('paint ') or 'decor level' in blob:
        return 'Colour and trim'

    safety_keywords = [
        'airbag', 'airbags', 'blind zone', 'collision', 'cruise', 'driver assistance', 'following distance',
        'hd surround vision', 'lane ', 'pedestrian', 'parking assist', 'rear cross traffic', 'rear pedestrian',
        'safety', 'seat belt', 'stability control', 'traffic sign', 'traction control', 'warning', 'alert',
        'camera', 'restraint', 'teen driver', 'sensing system', 'automatic emergency braking', 'brake assist',
        'reverse automatic braking', 'tire pressure monitor', 'rear park assist'
    ]
    if any(keyword in blob for keyword in safety_keywords):
        return 'Safety and driver assistance'

    technology_keywords = [
        'android auto', 'apple carplay', 'audio system', 'bluetooth', 'display', 'google built-in',
        'head-up display', 'infotainment', 'mychevrolet', 'navigation', 'onstar', 'phone', 'radio',
        'remote start', 'screen', 'siriusxm', 'smartphone', 'speaker', 'usb', 'wi-fi', 'wifi',
        'wireless', 'charging pad', 'charging-only', 'device charging', 'driver information center'
    ]
    if any(keyword in blob for keyword in technology_keywords):
        return 'Technology and connectivity'

    wheels_keywords = ['wheel', 'wheels', 'tire', 'tires', 'spare tire', 'spare wheel', 'lug nut', 'wheel lock']
    if any(keyword in blob for keyword in wheels_keywords):
        return 'Wheels and tires'

    mechanical_keywords = [
        'all-wheel drive', 'axle', 'battery', 'brakes', 'charging', 'charger', 'drive unit', 'drivetrain',
        'electric drive', 'engine', 'evot? ', 'fuel', 'gvwr', 'horsepower', 'motor', 'payload', 'performance',
        'powertrain', 'propulsion', 'range', 'rear axle', 'suspension', 'torque', 'tow', 'trailer',
        'trailering', 'transmission'
    ]
    if any(keyword in blob for keyword in mechanical_keywords):
        return 'Mechanical and performance'

    exterior_keywords = [
        'bed', 'box', 'bumper', 'cargo', 'cross rails', 'door', 'emblem', 'fascia', 'glass', 'grille',
        'headlamp', 'hood', 'lamp', 'liftgate', 'license plate', 'mirror, outside', 'nameplate', 'roof',
        'running board', 'splash guard', 'tailgate', 'window', 'wiper', 'privacy glass', 'deep tint'
    ]
    if any(keyword in blob for keyword in exterior_keywords):
        return 'Exterior and utility'

    interior_keywords = [
        'air conditioning', 'ambient', 'armrest', 'carpet', 'climate', 'console', 'cup holder', 'driver seat',
        'floor mat', 'headrest', 'heated seat', 'inside rearview', 'instrument panel', 'lumbar', 'rear seat',
        'seat adjuster', 'seat trim', 'seating', 'seats', 'steering wheel', 'sun visor', 'visor', 'interior'
    ]
    if any(keyword in blob for keyword in interior_keywords):
        return 'Interior and comfort'

    package_keywords = ['package', 'equipment group', 'lpo', 'accessory', 'dealer-installed', 'option']
    if any(keyword in blob for keyword in package_keywords):
        return 'Packages and options'

    if 'interior' in blob:
        return 'Interior and comfort'
    if 'exterior' in blob:
        return 'Exterior and utility'
    if 'mechanical' in blob:
        return 'Mechanical and performance'
    if 'wheels' in blob:
        return 'Wheels and tires'
    if 'onstar' in blob or 'siriusxm' in blob:
        return 'Technology and connectivity'
    
    return 'Other guide content'


def category_for_trim_feature(agg: TrimFeatureAggregate) -> str:
    """Infer category for a trim feature.
    
    Args:
        agg: Trim feature aggregate
        
    Returns:
        Category string
    """
    return _infer_feature_category(' '.join(agg.source_contexts), agg.title, agg.description)


def category_for_model_feature(agg: ModelFeatureAggregate) -> str:
    """Infer category for a model feature.
    
    Args:
        agg: Model feature aggregate
        
    Returns:
        Category string
    """
    return _infer_feature_category(' '.join(agg.source_contexts), agg.title, agg.description)
