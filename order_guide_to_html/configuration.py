from __future__ import annotations

from collections import OrderedDict
from typing import Dict, List, Optional, Sequence, Tuple
import re

from .utils import normalize_text, unique_preserve_order
from .models import ColorExteriorRow, ColorInteriorRow, ColorSheet, PowertrainTraileringGroup, SpecColumn, SpecGroupDoc, TrimDef, WorkbookData
from .parsing import parse_value_and_footnote_ids
from .classification import MODEL_CODE_RE
from .utils import MANIFEST_BODY_STYLE_TOKENS, MANIFEST_DRIVE_TOKENS, MANIFEST_STANDARDISH_CODES, first_unique, material_note_texts


def trim_matches_decor(trim: TrimDef, decor_value: str) -> bool:
    decor = normalize_text(decor_value)
    if not decor:
        return False
    parts = re.split(r'\s*/\s*|\s*,\s*', decor)
    trim_values = {trim.name.lower(), trim.code.lower()}
    for part in parts:
        p = normalize_text(part).lower()
        if not p:
            continue
        if p in trim_values:
            return True
        if trim.name.lower() == p:
            return True
        if trim.code.lower().startswith(p) or p.startswith(trim.code.lower()):
            return True
        name_words = trim.name.lower().split()
        if len(name_words) == 1 and p == name_words[0]:
            return True
    return False

def column_matches_trim(column: SpecColumn, trim: TrimDef) -> bool:
    blob = ' '.join([column.top_label, column.header] + column.header_lines).lower()
    if trim.name.lower() in blob or trim.code.lower() in blob:
        return True
    return False

def trim_header_list(data: WorkbookData) -> List[str]:
    return [normalize_text(trim.raw_header) for trim in data.trim_defs if normalize_text(trim.raw_header)]

def trim_name_list(data: WorkbookData) -> List[str]:
    return [normalize_text(trim.name) for trim in data.trim_defs if normalize_text(trim.name)]

def trim_code_list(data: WorkbookData) -> List[str]:
    return [normalize_text(trim.code) for trim in data.trim_defs if normalize_text(trim.code)]

def workbook_tab_metadata(data: WorkbookData) -> Dict[str, List[str]]:
    return {
        'source_tabs': list(data.sheet_names),
        'matrix_sheet_names': [sheet.name for sheet in data.matrix_sheets],
        'colour_and_trim_tabs': [sheet.name for sheet in data.color_sheets],
        'spec_sheet_names': unique_preserve_order(column.sheet_name for column in data.spec_columns),
        'engine_axle_tabs': unique_preserve_order(entry.sheet_name for entry in data.engine_axle_entries),
        'trailering_tabs': unique_preserve_order([record.sheet_name for record in data.trailering_records] + [record.sheet_name for record in data.gcwr_records]),
    }

def phrase_occurs_in_text(phrase: str, text: str) -> bool:
    phrase = normalize_text(phrase)
    text = normalize_text(text)
    if not phrase or not text:
        return False
    if re.fullmatch(r'[A-Za-z0-9]+', phrase):
        return re.search(rf'(?<![A-Za-z0-9]){re.escape(phrase)}(?![A-Za-z0-9])', text, re.I) is not None
    parts = [re.escape(part) for part in phrase.split()]
    pattern = r'(?<![A-Za-z0-9])' + r'\s+'.join(parts) + r'(?![A-Za-z0-9])'
    return re.search(pattern, text, re.I) is not None

def best_trim_match(data: WorkbookData, *texts: str) -> Optional[TrimDef]:
    blob = ' \n '.join(normalize_text(text) for text in texts if normalize_text(text))
    if not blob:
        return None
    scored: List[Tuple[int, TrimDef]] = []
    for trim in data.trim_defs:
        score = 0
        candidates = unique_preserve_order([trim.raw_header, trim.name, trim.code])
        for candidate in candidates:
            if phrase_occurs_in_text(candidate, blob):
                score = max(score, len(normalize_text(candidate)))
        if score > 0:
            scored.append((score, trim))
    if not scored:
        return None
    scored.sort(key=lambda item: (-item[0], item[1].name.lower(), item[1].code.lower()))
    if len(scored) > 1 and scored[0][0] == scored[1][0]:
        return None
    return scored[0][1]

def best_trim_match_for_spec_column(data: WorkbookData, column: SpecColumn) -> Optional[TrimDef]:
    texts = [column.top_label, column.header] + list(column.header_lines)
    return best_trim_match(data, *texts)

def find_cell_values_by_label_prefix(column: SpecColumn, prefixes: Sequence[str]) -> List[str]:
    values: List[str] = []
    for cell in column.cells:
        label = normalize_text(cell.label).lower()
        if any(label.startswith(prefix.lower()) for prefix in prefixes):
            values.append(cell.value)
    return unique_preserve_order(values)

def find_cell_values_by_label_contains(column: SpecColumn, snippets: Sequence[str]) -> List[str]:
    values: List[str] = []
    for cell in column.cells:
        label = normalize_text(cell.label).lower()
        if any(snippet.lower() in label for snippet in snippets):
            values.append(cell.value)
    return unique_preserve_order(values)

def spec_column_context_values(column: SpecColumn) -> List[str]:
    return unique_preserve_order([column.top_label, column.header] + list(column.header_lines))

def spec_column_context_text(column: SpecColumn) -> str:
    return ' | '.join(spec_column_context_values(column))

def section_names_for_column(column: SpecColumn) -> List[str]:
    return unique_preserve_order(cell.section or 'Data' for cell in column.cells)

def spec_column_engine_value(column: SpecColumn) -> Optional[str]:
    values = (
        find_cell_values_by_label_prefix(column, ['Engine', 'Electric drive unit', 'Electric drive', 'Drive unit']) +
        find_cell_values_by_label_contains(column, ['engine', 'electric drive', 'drive unit'])
    )
    return first_unique(values)

def spec_column_fuel_value(column: SpecColumn) -> Optional[str]:
    values = find_cell_values_by_label_contains(column, ['fuel'])
    return first_unique(values)

def spec_column_drivetrain_value(column: SpecColumn) -> Optional[str]:
    values = find_cell_values_by_label_contains(column, ['drive', 'drivetrain', 'propulsion'])
    choice = first_unique(values)
    if choice:
        return choice
    context_blob = ' | '.join(spec_column_context_values(column))
    for token in MANIFEST_DRIVE_TOKENS:
        if phrase_occurs_in_text(token, context_blob):
            return token
    return None

def spec_column_seating_value(column: SpecColumn) -> Optional[str]:
    return first_unique(find_cell_values_by_label_contains(column, ['seating capacity']))

def spec_column_body_style_value(column: SpecColumn) -> Optional[str]:
    direct = first_unique(find_cell_values_by_label_contains(column, ['body style']))
    if direct:
        return direct
    if normalize_text(column.top_label):
        return normalize_text(column.top_label)
    return None

def trim_colour_context(data: WorkbookData, trim: TrimDef) -> Dict[str, object]:
    interior_items: List[Tuple[ColorSheet, ColorInteriorRow, List[str]]] = []
    exterior_items: List[Tuple[ColorSheet, ColorExteriorRow, List[str], List[str]]] = []
    domain_notes: List[Tuple[str, str]] = []
    for sheet in data.color_sheets:
        relevant_interior_columns: List[str] = []
        matched_any_interior = False
        for row in sheet.interior_rows:
            if not trim_matches_decor(trim, row.decor_level):
                continue
            matched_any_interior = True
            color_lines = [f'{color}: {code}' for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--']
            relevant_interior_columns.extend([color for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--'])
            interior_items.append((sheet, row, color_lines))
        relevant_interior_columns = unique_preserve_order(relevant_interior_columns)
        for row in sheet.exterior_rows:
            title_value, title_note_ids = parse_value_and_footnote_ids(row.title)
            title_note_texts = [sheet.footnotes[nid] for nid in title_note_ids if nid in sheet.footnotes]
            active_columns = relevant_interior_columns if matched_any_interior and relevant_interior_columns else list(row.colors.keys())
            availability_lines = [f'{color}: {status}' for color, status in row.colors.items() if normalize_text(status) and color in active_columns]
            if availability_lines:
                exterior_items.append((sheet, row, availability_lines, title_note_texts))
        for note in material_note_texts(list(sheet.footnotes.values()) + list(sheet.bullet_notes)):
            domain_notes.append((sheet.name, note))
    return {
        'interior_items': interior_items,
        'exterior_items': exterior_items,
        'domain_notes': unique_preserve_order([f'{sheet_name}: {note}' for sheet_name, note in domain_notes]),
    }

def extract_model_code_from_text(*values: object) -> Optional[str]:
    for value in values:
        if isinstance(value, (list, tuple)):
            found = extract_model_code_from_text(*value)
            if found:
                return found
            continue
        text = normalize_text(value)
        if not text:
            continue
        match = MODEL_CODE_RE.search(text)
        if match:
            return match.group(0)
    return None

def spec_group_key(column: SpecColumn) -> Tuple[str, Tuple[str, ...]]:
    header_lines = tuple(normalize_text(x) for x in column.header_lines if normalize_text(x))
    if not header_lines:
        header_lines = (normalize_text(column.header),)
    return normalize_text(column.top_label), header_lines

def group_spec_columns_for_cpr(data: WorkbookData) -> List[SpecGroupDoc]:
    grouped: 'OrderedDict[Tuple[str, Tuple[str, ...]], SpecGroupDoc]' = OrderedDict()
    for column in data.spec_columns:
        key = spec_group_key(column)
        if key not in grouped:
            grouped[key] = SpecGroupDoc(
                top_label=normalize_text(column.top_label),
                header=normalize_text(column.header),
                header_lines=[normalize_text(x) for x in column.header_lines if normalize_text(x)],
                columns=[],
            )
        grouped[key].columns.append(column)
    return list(grouped.values())

def spec_group_model_code(group: SpecGroupDoc) -> Optional[str]:
    return extract_model_code_from_text(group.header_lines, group.header)

def spec_group_context_values(group: SpecGroupDoc) -> List[str]:
    return unique_preserve_order([group.top_label, group.header] + list(group.header_lines))

def spec_group_context_text(group: SpecGroupDoc) -> str:
    return ' | '.join(spec_group_context_values(group))

def spec_group_section_names(group: SpecGroupDoc) -> List[str]:
    names: List[str] = []
    for column in group.columns:
        names.extend(section_names_for_column(column))
    return unique_preserve_order(names)

def spec_group_first_value(group: SpecGroupDoc, extractor) -> Optional[str]:
    values: List[str] = []
    for column in group.columns:
        value = extractor(column)
        if normalize_text(value):
            values.append(normalize_text(value))
    return first_unique(values)

def group_powertrain_trailering_for_cpr(data: WorkbookData) -> List[PowertrainTraileringGroup]:
    grouped: 'OrderedDict[str, PowertrainTraileringGroup]' = OrderedDict()

    def ensure_group(model_code: str) -> PowertrainTraileringGroup:
        key = normalize_text(model_code)
        if key not in grouped:
            grouped[key] = PowertrainTraileringGroup(
                model_code=key,
                top_labels=[],
                engine_entries=[],
                trailering_records=[],
            )
        return grouped[key]

    for group in group_spec_columns_for_cpr(data):
        model_code = spec_group_model_code(group)
        if not model_code:
            continue
        g = ensure_group(model_code)
        if group.top_label:
            g.top_labels.append(group.top_label)

    for entry in data.engine_axle_entries:
        g = ensure_group(entry.model_code)
        if entry.top_label:
            g.top_labels.append(normalize_text(entry.top_label))
        g.engine_entries.append(entry)

    for record in data.trailering_records:
        g = ensure_group(record.model_code)
        g.trailering_records.append(record)

    ordered_groups: List[PowertrainTraileringGroup] = []
    for key, group in grouped.items():
        group.top_labels = unique_preserve_order(group.top_labels)
        if group.engine_entries or group.trailering_records:
            ordered_groups.append(group)
    return ordered_groups

def powertrain_group_trim_match(data: WorkbookData, group: PowertrainTraileringGroup) -> Optional[TrimDef]:
    texts = [group.model_code] + list(group.top_labels)
    texts.extend(entry.engine for entry in group.engine_entries)
    texts.extend(record.engine for record in group.trailering_records)
    return best_trim_match(data, *texts)
