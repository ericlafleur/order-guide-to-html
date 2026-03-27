from __future__ import annotations

from collections import OrderedDict
from typing import Dict, List, Optional, Sequence, Tuple
import re

from .utils import STATUS_LABELS, STATUS_PRIORITY, normalize_text, unique_preserve_order
from .models import MatrixRow, MatrixSheet, ModelFeatureAggregate, TrimDef, TrimFeatureAggregate
from .parsing import parse_status_value

def feature_title(label: str, orderable_code: str = '', reference_code: str = '') -> str:
    prefix = orderable_code or reference_code
    label = normalize_text(label)
    if prefix:
        return f'{prefix} | {label}'
    return label

def source_context(sheet_name: str, row_group: Optional[str] = None) -> str:
    sheet_name = normalize_text(sheet_name)
    row_group = normalize_text(row_group)
    if row_group and row_group.lower() != sheet_name.lower():
        return f'{sheet_name} | {row_group}'
    return sheet_name

def collect_row_note_texts(row: MatrixRow, sheet: MatrixSheet) -> List[str]:
    notes: List[str] = []
    for note_text in row.inline_footnotes.values():
        notes.append(note_text)
    for raw in row.status_by_trim.values():
        raw = normalize_text(raw)
        if not raw:
            continue
        m = re.match(r'^(--|[A-Z]+|[■□*]+)(.*)$', raw)
        suffix = m.group(2) if m else ''
        for note_id in re.findall(r'\d+', suffix):
            if note_id in row.inline_footnotes:
                notes.append(row.inline_footnotes[note_id])
            elif note_id in sheet.footnotes:
                notes.append(sheet.footnotes[note_id])
    notes.extend(row.bullet_notes)
    return unique_preserve_order([normalize_text(n) for n in notes if normalize_text(n)])

def summarize_model_status_groups(row: MatrixRow, trim_defs: Sequence[TrimDef], sheet: MatrixSheet) -> Tuple[Tuple[str, str, Tuple[str, ...]], ...]:
    groups: 'OrderedDict[Tuple[str, str], List[str]]' = OrderedDict()
    for trim in trim_defs:
        raw = normalize_text(row.status_by_trim.get(trim.key))
        if not raw:
            continue
        _code, label, _notes = parse_status_value(raw, row.inline_footnotes, sheet.footnotes)
        groups.setdefault((raw, label), []).append(trim.name or trim.code)
    return tuple((raw, label, tuple(names)) for (raw, label), names in groups.items())

def model_status_summary_lines(signature: Tuple[Tuple[str, str, Tuple[str, ...]], ...]) -> List[str]:
    lines = []
    for raw, label, names in signature:
        lines.append(f'{label} [{raw}]: {", ".join(names)}')
    return lines

def sort_trim_feature(agg: TrimFeatureAggregate) -> Tuple[int, str]:
    first_key = next(iter(agg.availability_contexts.keys()), ('', ''))
    raw, label = first_key
    return (STATUS_PRIORITY.get(label, 99), normalize_text(agg.title).lower(), raw)

CATEGORY_SEQUENCE = [
    'Safety and driver assistance',
    'Technology and connectivity',
    'Interior and comfort',
    'Exterior and utility',
    'Wheels and tires',
    'Mechanical and performance',
    'Packages and options',
    'Colour and trim',
    'Specifications and dimensions',
    'Engine, axle and GVWR',
    'Trailering and GCWR',
    'Other guide content',
]

CATEGORY_ORDER = {name: idx for idx, name in enumerate(CATEGORY_SEQUENCE)}

def infer_feature_category(*texts: str) -> str:
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

def sort_category_key(category: str) -> Tuple[int, str]:
    return (CATEGORY_ORDER.get(category, 999), normalize_text(category).lower())

def category_for_trim_feature(agg: TrimFeatureAggregate) -> str:
    return infer_feature_category(' '.join(agg.source_contexts), agg.title, agg.description)

def category_for_model_feature(agg: ModelFeatureAggregate) -> str:
    return infer_feature_category(' '.join(agg.source_contexts), agg.title, agg.description)

CONFIG_OBJECTTYPE = 'Configuration'

CONFIG_TYPE = 'configuration-spec'

CONFIG_KIND_SPEC_COLUMN = 'spec-column'

CONFIG_KIND_ENGINE_AXLE = 'engine-axle'

CONFIG_KIND_TRAILERING = 'trailering'

CONFIG_KIND_GCWR = 'gcwr'

MODEL_OBJECTTYPE = 'Product'

TRIM_OBJECTTYPE = 'Variant'

FEATURE_OBJECTTYPE = 'Feature'

COMPARISON_OBJECTTYPE = 'Comparison'

NOTE_OBJECTTYPE = 'Note'

DOMAIN_COLOR = 'Colour and trim'

DOMAIN_OVERVIEW = 'Overview'

DOMAIN_COMPARE = 'Comparison'

DOC_ROLE_ROOT = 'root'

DOC_ROLE_PARENT = 'parent'

DOC_ROLE_CHILD = 'child'

def source_tab_list_from_contexts(contexts: Sequence[str]) -> List[str]:
    tabs: List[str] = []
    for context in contexts:
        context = normalize_text(context)
        if not context:
            continue
        tabs.append(normalize_text(context.split('|', 1)[0]))
    return unique_preserve_order(tabs)

def source_tab_list_from_strings(*values: object) -> List[str]:
    tabs: List[str] = []
    for value in values:
        if isinstance(value, (list, tuple)):
            for item in value:
                tabs.extend(source_tab_list_from_strings(item))
        else:
            text = normalize_text(value)
            if text:
                tabs.append(text)
    return unique_preserve_order(tabs)

def with_doc_metadata(metadata: Dict[str, object], **extra: object) -> Dict[str, object]:
    combined = dict(metadata)
    combined.update(extra)
    return {k: v for k, v in combined.items() if v not in ('', [], None)}

def availability_pairs_for_trim(agg: TrimFeatureAggregate) -> Tuple[List[str], List[str]]:
    raws: List[str] = []
    labels: List[str] = []
    for (raw, label), _contexts in agg.availability_contexts.items():
        if raw:
            raws.append(raw)
        if label:
            labels.append(label)
    return unique_preserve_order(raws), unique_preserve_order(labels)

def trim_feature_is_present(agg: TrimFeatureAggregate) -> bool:
    raw_values, _labels = availability_pairs_for_trim(agg)
    return any(raw != '--' for raw in raw_values)

def availability_pairs_for_model(agg: ModelFeatureAggregate) -> Tuple[List[str], List[str]]:
    raws: List[str] = []
    labels: List[str] = []
    for signature in agg.availability_contexts.keys():
        for raw, label, _names in signature:
            if raw:
                raws.append(raw)
            if label:
                labels.append(label)
    return unique_preserve_order(raws), unique_preserve_order(labels)

def comparison_varies_by_trim(agg: ModelFeatureAggregate) -> bool:
    raws, labels = availability_pairs_for_model(agg)
    if len(raws) > 1 or len(labels) > 1:
        return True
    trim_sets = set()
    for signature in agg.availability_contexts.keys():
        for _raw, _label, names in signature:
            trim_sets.add(tuple(names))
    return len(trim_sets) > 1

def normalize_domain_value(domain: str) -> str:
    return normalize_text(domain) or DOMAIN_OVERVIEW

SURFACE_BOTH = 'both'

SURFACE_PASSAGE_ONLY = 'passage_only'

DOC_ROLE_ENTITY = 'entity'

DOC_ROLE_PASSAGE = 'passage'

CONFIG_KIND_SPEC_GROUP = 'spec-group'

CONFIG_KIND_POWERTRAIN_TRAILERING_GROUP = 'powertrain-trailering-group'

CONFIG_KIND_GCWR_REFERENCE = 'gcwr-reference'

MODEL_CODE_RE = re.compile(r'\b[A-Z]{2}\d{5}\b')
