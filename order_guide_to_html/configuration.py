from __future__ import annotations

from collections import OrderedDict
from typing import Dict, List, Optional, Sequence, Tuple
import re

from .utils import normalize_text, unique_preserve_order
from .models import ColorExteriorRow, ColorInteriorRow, ColorSheet, PowertrainTraileringGroup, SpecColumn, SpecGroupDoc, TrimDef, WorkbookData
from .parsing import parse_value_and_footnote_ids
from .classification import MODEL_CODE_RE
from .utils import MANIFEST_BODY_STYLE_TOKENS, MANIFEST_DRIVE_TOKENS, MANIFEST_STANDARDISH_CODES, first_unique, material_note_texts

BED_TYPE_RE = re.compile(r'\b(Short Bed|Standard Bed|Long Bed|Caisse courte|Caisse standard|Caisse longue)\b', re.IGNORECASE)


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

def strip_drive_tokens(name: str) -> str:
    pattern = r'\s*\b(?:' + '|'.join(MANIFEST_DRIVE_TOKENS) + r')\b'
    stripped = re.sub(pattern, '', name, flags=re.IGNORECASE)
    # 'À prop.' ends with '.' so \b fails at the end; handle it separately
    stripped = re.sub(r'\s*\bÀ prop\.', '', stripped, flags=re.IGNORECASE)
    stripped = MODEL_CODE_RE.sub('', stripped)
    stripped = re.sub(r'\s*\b[0-9][A-Z0-9]{4,6}\b', '', stripped)
    stripped = re.sub(r'\s*/\s*', ' ', stripped)
    stripped = re.sub(r'\s+', ' ', stripped)
    # Remove dangling conjunctions left over after model-code/drive-token stripping
    stripped = re.sub(r'^\s*et\s+|\s+et\s*$', '', stripped, flags=re.IGNORECASE)
    return stripped.strip()

def trim_name_list(data: WorkbookData) -> List[str]:
    return [strip_drive_tokens(normalize_text(trim.name)) for trim in data.trim_defs if normalize_text(trim.name)]

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

def all_trim_matches(data: WorkbookData, *texts: str) -> List[TrimDef]:
    blob = ' \n '.join(normalize_text(text) for text in texts if normalize_text(text))
    if not blob:
        return []
    matched: List[TrimDef] = []
    for trim in data.trim_defs:
        candidates = unique_preserve_order([trim.raw_header, trim.name, trim.code])
        for candidate in candidates:
            if phrase_occurs_in_text(candidate, blob):
                matched.append(trim)
                break
    return matched

def all_trim_matches_for_spec_group(data: WorkbookData, group: 'SpecGroupDoc') -> List[TrimDef]:
    """Find trims for a spec group, using three progressively broader strategies.

    1. Match full trim names / codes against the column header (works when the
       header already embeds a trim name or code, e.g. Trax "1TU58 FWD LT, 2RS
       and ACTIV").
    2. Match stripped trim names (no drive tokens, no model codes) against the
       header (works for Trailblazer-style headers like "1TR56 LS FWD" where the
       trim's raw name "LS 1TR56 FWD / 1TV56 AWD" doesn't appear verbatim).
    3. Scan spec cell labels in the group for trim name mentions (works for
       Tahoe-style workbooks with generic headers like "CC10706 / 2WD" where
       trim-specific dimension rows such as "Overall height, LS" carry the
       association).
    """
    # Strategy 1: full-name / code match in header text
    matched = all_trim_matches(data, group.top_label, group.header, *group.header_lines)
    if matched:
        return matched

    # Strategy 2: stripped-name match in header text
    blob = ' \n '.join(normalize_text(t) for t in [group.top_label, group.header] + list(group.header_lines) if normalize_text(t))
    if blob:
        stripped_matches: List[TrimDef] = []
        for trim in data.trim_defs:
            stripped = strip_drive_tokens(normalize_text(trim.name))
            if stripped and phrase_occurs_in_text(stripped, blob):
                stripped_matches.append(trim)
        if stripped_matches:
            return stripped_matches

    # Strategy 3: scan spec cell labels for trim name mentions
    cell_labels = [cell.label for col in group.columns for cell in col.cells]
    return all_trim_matches(data, *cell_labels)

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
        find_cell_values_by_label_prefix(column, ['Engine', 'Electric drive unit', 'Electric drive', 'Drive unit', 'Moteur', "Unité d'entraînement électrique"]) +
        find_cell_values_by_label_contains(column, ['engine', 'electric drive', 'drive unit', 'moteur', 'entraînement électrique'])
    )
    return first_unique(values)

def spec_column_fuel_value(column: SpecColumn) -> Optional[str]:
    values = find_cell_values_by_label_contains(column, ['fuel', 'carburant'])
    return first_unique(values)

def spec_column_drivetrain_value(column: SpecColumn) -> Optional[str]:
    raw_values = find_cell_values_by_label_contains(column, ['drive', 'drivetrain', 'propulsion', 'motricité', 'transmission'])
    # Trailering rows contain 'propulsion' in their labels but their *values* are
    # weights like '1500 (680)', not drivetrain identifiers.  Only keep values
    # that actually look like a drivetrain: contain a recognised drive token or
    # the phrase "wheel drive".
    def _is_drivetrain_value(v: str) -> bool:
        v_lower = normalize_text(v).lower()
        if 'wheel drive' in v_lower:
            return True
        return any(phrase_occurs_in_text(token, v) for token in MANIFEST_DRIVE_TOKENS)
    values = [v for v in raw_values if _is_drivetrain_value(v)]
    choice = first_unique(values)
    if choice:
        return choice
    context_blob = ' | '.join(spec_column_context_values(column))
    for token in MANIFEST_DRIVE_TOKENS:
        if phrase_occurs_in_text(token, context_blob):
            return token
    return None

def spec_column_seating_value(column: SpecColumn) -> Optional[str]:
    return first_unique(find_cell_values_by_label_contains(column, ['seating capacity', 'nombre de places', 'places assises']))

def trim_drivetrains(data: WorkbookData, trim: TrimDef) -> List[str]:
    """Return all drivetrain values applicable to a trim.

    Three-tier lookup:

    1. Explicitly-matched groups: spec groups where all_trim_matches_for_spec_group
       returns a non-empty list and includes this trim.  Covers Trailblazer-style
       (header contains trim name) and Tahoe-style (cell labels mention trim name).

    2. Generic groups: spec groups where all_trim_matches_for_spec_group returns []
       (no trim specifically identified — i.e. the configuration applies to all
       trims not covered by an explicit group).  Covers Silverado-style workbooks
       where most spec columns carry a body-config code like "CC30743 / 2WD Crew
       Cab" with no trim marker.  A trim is considered "covered by an explicit
       group" only if it appears in at least one group's explicit match list, in
       which case we skip the generic groups for it (e.g. Silverado High Country
       has its own CK-only columns, so it should not also inherit the generic
       CC/CK drivetrains).

    3. Fallback: drive tokens found in the trim's own raw header / name (e.g.
       "RS 1TY56 AWD" → AWD).
    """
    groups = group_spec_columns_for_cpr(data)
    matches_by_group: List[List[TrimDef]] = [all_trim_matches_for_spec_group(data, g) for g in groups]

    # Trims that appear in at least one explicit (non-generic) match list
    explicitly_covered: List[TrimDef] = [t for m in matches_by_group for t in m]

    values: List[str] = []

    for group, matches in zip(groups, matches_by_group):
        include = False
        if matches and trim in matches:
            # Strategy 1: explicit specific match
            include = True
        elif not matches and trim not in explicitly_covered:
            # Strategy 2: generic group, trim has no specific group of its own
            include = True
        if include:
            for col in group.columns:
                dt = spec_column_drivetrain_value(col)
                if dt and dt not in values:
                    values.append(dt)

    # Strategy 3: drive token in trim name / header.
    # Runs when no values were found, but also overrides values that came from
    # a generic (multi-trim) spec-group when the trim's own header/name already
    # encodes a specific drivetrain — e.g. "Work Truck 4WD" must not inherit
    # "2WD" from a shared "2WD / 4WD WT Crew Cab" spec-group header.
    blob = normalize_text(trim.raw_header) + ' ' + normalize_text(trim.name)
    name_tokens = [token for token in MANIFEST_DRIVE_TOKENS if phrase_occurs_in_text(token, blob)]
    if name_tokens and (not values or trim not in explicitly_covered):
        return name_tokens

    # Strategy 4: propulsion rows in matrix sheets (EV workbooks).
    # Covers vehicles like Equinox EV and Bolt where drivetrain is not present in
    # spec columns or trim headers.  Scans every matrix sheet for rows whose
    # description starts with "Propulsion," and extracts drive tokens for any trim
    # that has a non-"--" availability status in that row.  Captures all drivetrains
    # the trim can have (both standard and available).
    if not values:
        matrix_tokens: List[str] = []
        for ms in data.matrix_sheets:
            for row in ms.rows:
                desc = normalize_text(row.description_main)
                if not desc.lower().startswith('propulsion,'):
                    continue
                status = row.status_by_trim.get(trim.key, '--')
                if not status or status.strip() == '--':
                    continue
                for token in MANIFEST_DRIVE_TOKENS:
                    if phrase_occurs_in_text(token, desc) and token not in matrix_tokens:
                        matrix_tokens.append(token)
        if matrix_tokens:
            return matrix_tokens

    return values


def spec_group_body_style_label(group: SpecGroupDoc) -> Optional[str]:
    """Return a composite body style label for a spec group.

    Combines the cab type from top_label (e.g. 'Crew Cab') with a bed type
    extracted from the header lines (e.g. 'Short Bed', 'Standard Bed',
    'Long Bed') to produce a label like 'Crew Cab, Short Bed'.  Falls back
    to the bare cab type when no bed type is present, and to None when the
    top_label has no recognised body style token.
    """
    cab_type = normalize_text(group.top_label)
    if not cab_type or not any(token.lower() in cab_type.lower() for token in MANIFEST_BODY_STYLE_TOKENS):
        return None
    bed_type: Optional[str] = None
    for line in group.header_lines:
        m = BED_TYPE_RE.search(normalize_text(line))
        if m:
            bed_type = m.group(1)
            break
    if bed_type:
        return f'{cab_type}, {bed_type}'
    return cab_type


def trim_body_styles(data: WorkbookData, trim: TrimDef) -> List[str]:
    """Return all body style values applicable to a trim.

    Uses the same three-tier logic as trim_drivetrains:
    1. Explicit match groups (trim name appears in header or cell labels).
    2. Generic groups (no trim matched at all) for trims not explicitly covered.
    3. Body style tokens found in the trim's own raw header / name as a fallback.

    Each group contributes a composite label combining the cab type with bed
    type when available (e.g. 'Crew Cab, Short Bed').
    """
    groups = group_spec_columns_for_cpr(data)
    matches_by_group: List[List[TrimDef]] = [all_trim_matches_for_spec_group(data, g) for g in groups]
    explicitly_covered: List[TrimDef] = [t for m in matches_by_group for t in m]

    values: List[str] = []
    for group, matches in zip(groups, matches_by_group):
        include = False
        if matches and trim in matches:
            include = True
        elif not matches and trim not in explicitly_covered:
            include = True
        if include:
            bs = spec_group_body_style_label(group)
            if bs and bs not in values:
                values.append(bs)

    if not values:
        blob = normalize_text(trim.raw_header) + ' ' + normalize_text(trim.name)
        for token in MANIFEST_BODY_STYLE_TOKENS:
            if token.lower() in blob.lower() and token not in values:
                values.append(token)

    return values


def trim_seating(data: WorkbookData, trim: TrimDef) -> List[str]:
    """Return all seating capacity values applicable to a trim.

    Uses the same three-tier logic as trim_drivetrains / trim_body_styles:
    1. Explicitly-matched groups (trim name in header or cell labels).
    2. Generic groups (no trim matched) for trims not explicitly covered.
    3. No fallback — seating does not appear in trim names or headers.
    """
    groups = group_spec_columns_for_cpr(data)
    matches_by_group: List[List[TrimDef]] = [all_trim_matches_for_spec_group(data, g) for g in groups]
    explicitly_covered: List[TrimDef] = [t for m in matches_by_group for t in m]

    values: List[str] = []
    for group, matches in zip(groups, matches_by_group):
        include = False
        if matches and trim in matches:
            include = True
        elif not matches and trim not in explicitly_covered:
            include = True
        if include:
            for col in group.columns:
                sv = spec_column_seating_value(col)
                if sv and sv not in values:
                    values.append(sv)
    return values


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
