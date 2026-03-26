#!/usr/bin/env python3
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

BASE_SCRIPT = Path('/mnt/data/order_guide_to_html_rag_cpr_coarse.py')

spec = importlib.util.spec_from_file_location('order_guide_base_cpr_coarse', str(BASE_SCRIPT))
if spec is None or spec.loader is None:
    raise RuntimeError(f'Unable to load base script: {BASE_SCRIPT}')
base = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = base
spec.loader.exec_module(base)

normalize_text = base.normalize_text
unique_preserve_order = base.unique_preserve_order
htmlize_text = base.htmlize_text

SHEET_SEGMENT_RE = re.compile(
    r'^(?:Standard Equipment|Equipment Groups|PEG Stairstep|Interior|Exterior|Mechanical|'
    r'Engine Axles(?:\s+\d+)?|Colour and Trim(?:\s+\d+)?|Color and Trim(?:\s+\d+)?|'
    r'SEO Ship Thru|OnStar SiriusXM Fleet Options|Dimensions(?:\s+\d+)?|Specs(?:\s+\d+)?|'
    r'Wheels|Trailering Specs(?:\s+\d+)?|All)$',
    re.I,
)
STATUS_BRACKET_RE = re.compile(r'\s*\[(?:--|[A-Z]+\d*|[■□*]+\d*)\]')
CODE_PARENS_RE = re.compile(r'\(([^()]*)\)')
SOURCE_PARENS_HINT_RE = re.compile(
    r'(Standard Equipment|Equipment Groups|PEG Stairstep|Interior|Exterior|Mechanical|Wheels|Dimensions|Specs|'
    r'OnStar|SiriusXM|Colour and Trim|Color and Trim|Trailering|Engine Axles|SEO Ship Thru|All)',
    re.I,
)
CODE_ONLY_RE = re.compile(r'^(?=.*[A-Z])[A-Z0-9-]{2,10}$')
FILE_CODE_RE = re.compile(r'\b[A-Z]{2}\d{5}\b')

FIELD_DROP_LABELS = {
    'Source tab', 'Source tabs', 'Source context', 'Guide category', 'Trim code', 'Trim header from guide',
    'Legend from guide', 'Column context', 'Configuration context', 'Top label', 'Guide categories',
    'Guide sections', 'Availability raw value', 'Model code', 'Table title', 'Table titles',
    'Trim headers from guide', 'Configuration top labels', 'Configuration top label',
    'Configuration header', 'Configuration header lines',
}
FIELD_RENAME = {
    'Guide text': 'Details',
    'Guide values': 'Specifications',
    'Guide notes': 'Notes',
    'Feature lines from guide': 'Feature highlights',
    'Colour and trim lines from guide': 'Colour and trim combinations',
    'Availability on this trim': 'Availability',
    'Availability by trim': 'Availability by trim',
    'Paint note': 'Notes',
    'Engines from guide': 'Engines',
    'Trailering rating types': 'Trailering ratings',
}
MANIFEST_DROP_KEYS = {
    'workbook', 'source_tabs', 'source_tab', 'matrix_sheet_names', 'colour_and_trim_tabs', 'spec_sheet_names',
    'engine_axle_tabs', 'trailering_tabs', 'trim_codes_from_guide', 'trim_code', 'trim_header_from_guide',
    'trim_headers_from_guide', 'model_code', 'configuration_top_label', 'configuration_header_lines',
    'guide_sections', 'table_title', 'table_titles',
}
MANIFEST_RENAME_KEYS = {
    'configuration_header': 'configuration_label',
    'configuration_top_labels': 'configuration_labels',
}
ACTIVE_GLOSSARY: Dict[str, str] = {}
WITH_CODES_BRACKETS_RE = re.compile(r'\[w/\s*([^\]]+)\]', re.I)
LOW_SIGNAL_COMPARISON_DOMAINS = {
    # Keep the comparison corpus focused on domains that are more likely to matter in buyer questions.
    'Colour and trim',
}


def is_code_only(text: str) -> bool:
    text = normalize_text(text)
    if not text or not CODE_ONLY_RE.fullmatch(text):
        return False
    return any(ch.isdigit() for ch in text) or len(text) > 4


def expand_code_sequence(seq: str) -> str:
    parts = [normalize_text(x) for x in re.split(r'\s+and\s+|\s*,\s*|\s*;\s*', normalize_text(seq)) if normalize_text(x)]
    descriptions: List[str] = []
    leftovers: List[str] = []
    for part in parts:
        if normalize_text(part) in ACTIVE_GLOSSARY:
            descriptions.append(clean_customer_text(ACTIVE_GLOSSARY[normalize_text(part)]))
        elif is_code_only(part):
            leftovers.append('')
        else:
            leftovers.append(clean_customer_text(part))
    descriptions = [d for d in descriptions if d]
    leftovers = [x for x in leftovers if x]
    joined = unique_preserve_order(descriptions + leftovers)
    return ' with ' + ' and '.join(joined) if joined else ''


def _expand_with_codes(match: re.Match) -> str:
    return expand_code_sequence(match.group(1))


def _strip_code_parenthetical(match: re.Match) -> str:
    content = normalize_text(match.group(1))
    if re.fullmatch(r'[A-Z0-9-]{2,10}', content) and re.search(r'[A-Z]', content):
        return ''
    return f'({content})'


def clean_customer_text(text: object) -> str:
    text = normalize_text(text)
    if not text:
        return ''
    text = WITH_CODES_BRACKETS_RE.sub(_expand_with_codes, text)
    text = CODE_PARENS_RE.sub(_strip_code_parenthetical, text)
    text = FILE_CODE_RE.sub('', text)
    text = STATUS_BRACKET_RE.sub('', text)
    text = re.sub(r'\b(?:Trim code|Model code|Feature code|Reference code|Orderable code)\s*:?\s*[A-Z0-9-]{2,10}\b', '', text, flags=re.I)
    if SOURCE_PARENS_HINT_RE.search(text):
        text = re.sub(r'\s+\(([^()]*)\)$', lambda m: '' if (';' in m.group(1) or SOURCE_PARENS_HINT_RE.search(m.group(1))) else m.group(0), text)
    text = text.replace('includes ,', 'includes ')
    text = re.sub(r'\s+,', ',', text)
    text = re.sub(r',\s*,', ', ', text)
    text = re.sub(r'\s{2,}', ' ', text)
    text = re.sub(r'\|\s*\|', '|', text)
    text = re.sub(r'\s*\|\s*$', '', text)
    text = re.sub(r'^\s*\|\s*', '', text)
    text = re.sub(r'\s+([.;:])', r'\1', text)
    return normalize_text(text).strip(' |,;')


def clean_heading_text(text: object) -> str:
    text = normalize_text(text)
    if not text:
        return ''
    replacements = {
        'Vehicle Order Guide trim overview': 'Trim overview',
        'Vehicle Order Guide model overview': 'Model overview',
        'grouped guide passage': 'Highlights',
        'trim comparison grouped passage': 'Trim comparison highlights',
        'comparison from guide': 'comparison',
        'guide values': 'Specifications',
        'from guide': '',
        'identity from guide': 'Overview',
        'values from guide': 'Values',
        'reference from guide': 'Reference',
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    segments: List[str] = []
    seen = set()
    for seg in text.split('|'):
        seg = clean_customer_text(seg)
        if not seg:
            continue
        if SHEET_SEGMENT_RE.fullmatch(seg):
            continue
        if is_code_only(seg) or FILE_CODE_RE.fullmatch(seg):
            continue
        key = seg.lower()
        if key in seen:
            continue
        seen.add(key)
        segments.append(seg)
    return ' | '.join(segments)


def dedupe_fields(fields: Sequence[Tuple[str, str]]) -> List[Tuple[str, str]]:
    cleaned: List[Tuple[str, str]] = []
    seen = set()
    for label, value in fields:
        label = normalize_text(label)
        value = normalize_text(value)
        if not label or not value:
            continue
        key = (label, value)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append((label, value))
    return cleaned


def filtered_identity_fields(
    data,
    trim=None,
    *,
    category: str = '',
    source_context: str = '',
    source_tabs: str = '',
    extra_fields: Sequence[Tuple[str, str]] = (),
):
    fields: List[Tuple[str, str]] = [('Vehicle', data.vehicle_name)]
    if trim is not None:
        fields.append(('Trim', trim.name))
    if category:
        fields.append(('Category', category))
    for label, value in extra_fields:
        label = normalize_text(label)
        if label in FIELD_DROP_LABELS:
            continue
        label = FIELD_RENAME.get(label, label)
        value = clean_customer_text(value)
        if not value or is_code_only(value):
            continue
        fields.append((label, value))
    return dedupe_fields(fields)


def cleaned_render_article(title: str, fields: Sequence[Tuple[str, str]], bullet_groups: Sequence[Tuple[str, Sequence[str]]] = ()) -> str:
    title = clean_heading_text(title)
    parts = [f'<article class="guide-record"><h3>{htmlize_text(title)}</h3>']
    clean_fields: List[Tuple[str, str]] = []
    seen_fields = set()
    for label, value in fields:
        label = normalize_text(label)
        if label in FIELD_DROP_LABELS:
            continue
        label = FIELD_RENAME.get(label, label)
        value = clean_customer_text(value)
        if not label or not value or is_code_only(value):
            continue
        item = (label, value)
        if item in seen_fields:
            continue
        seen_fields.add(item)
        clean_fields.append(item)
    for label, value in clean_fields:
        parts.append(f'<p><strong>{base.html.escape(label)}:</strong> {htmlize_text(value)}</p>')
    for label, items in bullet_groups:
        label = normalize_text(label)
        if label in FIELD_DROP_LABELS:
            continue
        label = FIELD_RENAME.get(label, label)
        clean_items: List[str] = []
        seen_items = set()
        for item in items:
            item = clean_customer_text(item)
            if not item or is_code_only(item):
                continue
            if item in seen_items:
                continue
            seen_items.add(item)
            clean_items.append(item)
        if not clean_items:
            continue
        parts.append(f'<div class="record-list"><p><strong>{base.html.escape(label)}:</strong></p><ul>')
        for item in clean_items:
            parts.append(f'<li>{htmlize_text(item)}</li>')
        parts.append('</ul></div>')
    parts.append('</article>')
    return ''.join(parts)


def clean_trim_heading(data, trim) -> str:
    return normalize_text(f'{data.vehicle_name} {trim.name}')


def clean_feature_title(label: str, orderable_code: str = '', reference_code: str = '') -> str:
    return clean_customer_text(label)


def clean_compact_text(text: str, max_words: int = 20) -> str:
    return clean_customer_text(text)


def clean_model_status_summary_lines(signature):
    lines = []
    for _raw, label, names in signature:
        if names:
            lines.append(f'{clean_customer_text(label)}: {", ".join(clean_customer_text(name) for name in names if clean_customer_text(name))}')
    return [line for line in lines if normalize_text(line)]


def clean_availability_summary_for_trim(agg):
    labels = [clean_customer_text(label) for (_raw, label), _contexts in agg.availability_contexts.items()]
    labels = unique_preserve_order(labels)
    return '; '.join(labels)


def clean_availability_lines_for_trim(agg):
    labels = [clean_customer_text(label) for (_raw, label), _contexts in agg.availability_contexts.items()]
    return unique_preserve_order(labels)


def clean_availability_summary_for_model(agg):
    parts: List[str] = []
    for signature, _contexts in agg.availability_contexts.items():
        parts.extend(clean_model_status_summary_lines(signature))
    return ' / '.join(unique_preserve_order(parts))


def clean_availability_lines_for_model(agg):
    lines: List[str] = []
    for signature, _contexts in agg.availability_contexts.items():
        lines.extend(clean_model_status_summary_lines(signature))
    return unique_preserve_order(lines)


def clean_trim_group_line(agg):
    feature_text = clean_customer_text(agg.description or agg.title)
    availability = clean_availability_summary_for_trim(agg)
    if availability:
        return f'{feature_text} — {availability}'
    return feature_text


def clean_model_group_line(agg):
    feature_text = clean_customer_text(agg.description or agg.title)
    availability = clean_availability_summary_for_model(agg)
    if availability:
        return f'{feature_text} — {availability}'
    return feature_text


def trim_feature_is_unavailable(agg) -> bool:
    statuses = []
    for (raw, label), _contexts in agg.availability_contexts.items():
        raw_value = normalize_text(raw)
        label_value = normalize_text(clean_customer_text(label)).lower()
        statuses.append((raw_value, label_value))
    if not statuses:
        return False
    return all(label == 'not available' or raw == '--' for raw, label in statuses)


def render_feature_section(data, entity: str, category_groups: Dict[str, List[object]], *, trim=None, model_mode: bool, section_class: str, section_heading: str, bullet_label: str, title_suffix: str) -> str:
    if not category_groups:
        return ''
    parts = [f'<section class="{section_class}"><h2>{base.html.escape(entity)} | {base.html.escape(section_heading)}</h2>']
    for category in sorted(category_groups.keys(), key=base.sort_category_key):
        items = category_groups[category]
        lines = [clean_model_group_line(item) if model_mode else clean_trim_group_line(item) for item in items]
        for idx, line_chunk in enumerate(base.chunk_feature_items(lines), start=1):
            title = f'{category} | {title_suffix}'
            if idx > 1:
                title += f' | part {idx}'
            parts.append(
                cleaned_render_article(
                    article_heading_clean(entity, title),
                    filtered_identity_fields(data, trim, category=category),
                    [(bullet_label, line_chunk)],
                )
            )
    parts.append('</section>')
    return ''.join(parts)


def render_minimal_page_identity_section(data, trim=None) -> str:
    entity = clean_trim_heading(data, trim) if trim is not None else data.vehicle_name
    if trim is None:
        fields = [('Vehicle', data.vehicle_name)]
        trim_names = [normalize_text(t.name) for t in data.trim_defs if normalize_text(t.name)]
        bullets = [('Available trims', unique_preserve_order(trim_names))] if trim_names else []
        title = f'{entity} | Model overview'
    else:
        fields = [('Vehicle', data.vehicle_name), ('Trim', trim.name)]
        bullets = []
        title = f'{entity} | Trim overview'
    return '<section class="vehicle-identity">' + cleaned_render_article(title, fields, bullets) + '</section>'


def render_no_matrix_legend_section(data, page_title):
    return ''


def render_note_articles_clean(title_prefix: str, note_texts: Sequence[str], context_text: str, label: str = 'Note') -> str:
    parts = []
    for note_index, note in enumerate(unique_preserve_order(note_texts), start=1):
        chunks = base.sentence_chunks(note, max_words=105)
        for chunk_index, chunk in enumerate(chunks, start=1):
            title = f'{title_prefix} | {label.lower()} {note_index}'
            if len(chunks) > 1:
                title += f' | part {chunk_index}'
            parts.append(cleaned_render_article(title, [(label, chunk)]))
    return ''.join(parts)


def clean_title_document(title: str, *body_parts: str) -> str:
    title = clean_heading_text(title)
    parts = ['<html><head><meta charset="utf-8"></head><body>', f'<h1>{base.html.escape(normalize_text(title))}</h1>']
    parts.extend(part for part in body_parts if part)
    parts.append('</body></html>')
    return ''.join(parts)


def matrixrow_label(self) -> str:
    text = self.description_main or self.description_raw
    text = normalize_text(text).split('\n')[0]
    text = re.sub(r'^NEW!\s+', '', text, flags=re.I)
    if ', includes ' in text.lower():
        text = text.split(',', 1)[0]
    elif '. ' in text and len(text.split('. ', 1)[0]) >= 12:
        text = text.split('. ', 1)[0]
    return clean_customer_text(text)


def clean_manifest_value(value):
    if isinstance(value, str):
        return clean_heading_text(value) if ('|' in value or value.lower().endswith('overview') or 'comparison' in value.lower()) else clean_customer_text(value)
    if isinstance(value, list):
        cleaned = []
        for item in value:
            item = clean_manifest_value(item)
            if item in ('', [], None):
                continue
            cleaned.append(item)
        # preserve order
        out = []
        seen = set()
        for item in cleaned:
            key = json.dumps(item, sort_keys=True, ensure_ascii=False) if isinstance(item, (dict, list)) else str(item)
            if key in seen:
                continue
            seen.add(key)
            out.append(item)
        return out
    if isinstance(value, dict):
        return clean_manifest_entry(value)
    return value


def clean_manifest_entry(entry: Dict[str, object]) -> Dict[str, object]:
    cleaned: Dict[str, object] = {}
    for key, value in entry.items():
        if key in MANIFEST_DROP_KEYS:
            continue
        key = MANIFEST_RENAME_KEYS.get(key, key)
        value = clean_manifest_value(value)
        if value in ('', [], None):
            continue
        cleaned[key] = value
    if cleaned.get('type') == 'model' and 'trim_names_from_guide' in cleaned:
        cleaned['available_trims'] = cleaned.pop('trim_names_from_guide')
    return cleaned


original_build_manifest = base.build_manifest_from_bindings
original_build_model_and_comparison_records = base.build_model_and_comparison_records
original_write_outputs = base.write_outputs


def write_outputs_with_glossary(data, output_dir):
    global ACTIVE_GLOSSARY
    ACTIVE_GLOSSARY = {normalize_text(k): normalize_text(v) for k, v in getattr(data, 'glossary', {}).items() if normalize_text(k) and normalize_text(v)}
    return original_write_outputs(data, output_dir)


def build_manifest_from_bindings_clean(data, bindings, manifest_path: Path):
    manifest = original_build_manifest(data, bindings, manifest_path)
    cleaned = {
        'vehicle_name': clean_customer_text(manifest.get('vehicle_name', '')),
        'vehicle_key': manifest.get('vehicle_key', ''),
        'files': [clean_manifest_entry(entry) for entry in manifest.get('files', [])],
    }
    cleaned = {k: v for k, v in cleaned.items() if v not in ('', [], None)}
    manifest_path.write_text(json.dumps(cleaned, indent=2, ensure_ascii=False), encoding='utf-8')
    return cleaned


def build_model_and_comparison_records_clean(data, output_dir, used_names, bindings):
    model_path = original_build_model_and_comparison_records(data, output_dir, used_names, bindings)
    # Remove low-signal comparison docs if present.
    filtered = []
    for binding in bindings:
        if binding.record.type == 'comparison':
            domain = normalize_text(binding.record.metadata.get('domain', ''))
            if domain in LOW_SIGNAL_COMPARISON_DOMAINS:
                try:
                    binding.record.path.unlink(missing_ok=True)
                except Exception:
                    pass
                continue
        filtered.append(binding)
    bindings[:] = filtered
    return model_path



def article_heading_clean(entity: str, base_title: str) -> str:
    base_title = normalize_text(base_title)
    if not base_title:
        return clean_heading_text(entity)
    return clean_heading_text(f'{entity} | {base_title}')


def clean_model_colour_group_lines(data):
    lines: List[str] = []
    for sheet in data.color_sheets:
        for row in sheet.interior_rows:
            colors = [clean_customer_text(color) for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--']
            parts = ['Interior trim', row.decor_level, row.seat_type, row.seat_trim]
            summary = ' | '.join(clean_customer_text(x) for x in parts if clean_customer_text(x))
            if colors:
                summary += ' — Available colours: ' + '; '.join(unique_preserve_order(colors))
            lines.append(summary)
        for row in sheet.exterior_rows:
            title_value, _title_note_ids = base.parse_value_and_footnote_ids(row.title)
            availability_lines = []
            for color, status in row.colors.items():
                status_code, status_label, _ = base.parse_status_value(status, {}, sheet.footnotes)
                color = clean_customer_text(color)
                if not color or not status_label:
                    continue
                if status_code == '--':
                    availability_lines.append(f'{color}: Not available')
                else:
                    availability_lines.append(f'{color}: {clean_customer_text(status_label)}')
            summary = ' | '.join(clean_customer_text(x) for x in ['Exterior paint', title_value or row.title] if clean_customer_text(x))
            if availability_lines:
                summary += ' — ' + '; '.join(unique_preserve_order(availability_lines))
            lines.append(summary)
    return unique_preserve_order(lines)


def clean_trim_colour_group_lines(data, trim):
    ctx = base.trim_colour_context(data, trim)
    lines: List[str] = []
    for _sheet, row, color_lines in ctx['interior_items']:
        colors = [clean_customer_text(line.split(':', 1)[0]) for line in color_lines if clean_customer_text(line.split(':', 1)[0])]
        summary = ' | '.join(clean_customer_text(x) for x in ['Interior trim', row.decor_level, row.seat_type, row.seat_trim] if clean_customer_text(x))
        if colors:
            summary += ' — Available colours: ' + '; '.join(unique_preserve_order(colors))
        lines.append(summary)
    for sheet, row, availability_lines, _note_texts in ctx['exterior_items']:
        title_value, _title_note_ids = base.parse_value_and_footnote_ids(row.title)
        clean_lines = []
        for line in availability_lines:
            color, _, raw_status = line.partition(':')
            status_code, status_label, _ = base.parse_status_value(raw_status, {}, sheet.footnotes)
            color = clean_customer_text(color)
            if not color or not status_label:
                continue
            if status_code == '--':
                clean_lines.append(f'{color}: Not available')
            else:
                clean_lines.append(f'{color}: {clean_customer_text(status_label)}')
        summary = ' | '.join(clean_customer_text(x) for x in ['Exterior paint', title_value or row.title] if clean_customer_text(x))
        if clean_lines:
            summary += ' — ' + '; '.join(unique_preserve_order(clean_lines))
        lines.append(summary)
    return unique_preserve_order(lines)


def clean_render_grouped_colour_summary(data, trim=None):
    if not data.color_sheets:
        return ''
    entity = clean_trim_heading(data, trim) if trim is not None else data.vehicle_name
    lines = clean_trim_colour_group_lines(data, trim) if trim is not None else clean_model_colour_group_lines(data)
    if not lines:
        return ''
    parts = [f'<section class="grouped-colour-passages"><h2>{base.html.escape(entity)} | Colour and trim</h2>']
    for idx, chunk in enumerate(base.chunk_feature_items(lines), start=1):
        chunk_title = article_heading_clean(entity, 'Colour and trim') if idx == 1 else article_heading_clean(entity, f'Colour and trim | part {idx}')
        parts.append(
            cleaned_render_article(
                chunk_title,
                filtered_identity_fields(data, trim, category=base.DOMAIN_COLOR),
                [('Colour and trim combinations', chunk)],
            )
        )
    parts.append('</section>')
    return ''.join(parts)


def clean_grouped_feature_sections(data, features, *, trim=None, model_mode=False):
    if not features:
        return ''
    entity = clean_trim_heading(data, trim) if trim is not None else data.vehicle_name
    if trim is not None and not model_mode:
        positive_groups: Dict[str, List[object]] = base.OrderedDict()
        unavailable_groups: Dict[str, List[object]] = base.OrderedDict()
        for feature in features:
            category = base.category_for_trim_feature(feature)
            target = unavailable_groups if trim_feature_is_unavailable(feature) else positive_groups
            target.setdefault(category, []).append(feature)
        return ''.join([
            render_feature_section(
                data,
                entity,
                positive_groups,
                trim=trim,
                model_mode=False,
                section_class='grouped-feature-passages',
                section_heading='Feature highlights',
                bullet_label='Feature highlights',
                title_suffix='Highlights',
            ),
            render_feature_section(
                data,
                entity,
                unavailable_groups,
                trim=trim,
                model_mode=False,
                section_class='unavailable-feature-passages',
                section_heading='Features explicitly not offered on this trim',
                bullet_label='Unavailable features',
                title_suffix='Unavailable features',
            ),
        ])

    category_groups: Dict[str, List[object]] = base.OrderedDict()
    for feature in features:
        category = base.category_for_model_feature(feature) if model_mode else base.category_for_trim_feature(feature)
        category_groups.setdefault(category, []).append(feature)
    return render_feature_section(
        data,
        entity,
        category_groups,
        trim=trim,
        model_mode=model_mode,
        section_class='grouped-feature-passages',
        section_heading='Feature highlights',
        bullet_label='Feature highlights',
        title_suffix='Highlights',
    )


# Monkey patches.
base.full_trim_heading = clean_trim_heading
base.feature_title = clean_feature_title
base.compact_text = clean_compact_text
base.model_status_summary_lines = clean_model_status_summary_lines
base.availability_summary_for_trim = clean_availability_summary_for_trim
base.availability_lines_for_trim = clean_availability_lines_for_trim
base.availability_summary_for_model = clean_availability_summary_for_model
base.availability_lines_for_model = clean_availability_lines_for_model
base.trim_group_line = clean_trim_group_line
base.model_group_line = clean_model_group_line
base.identity_fields = filtered_identity_fields
base.render_article = cleaned_render_article
base.render_page_identity_section = render_minimal_page_identity_section
base.render_matrix_legend_section = render_no_matrix_legend_section
base.render_note_articles = render_note_articles_clean
base.grouped_feature_sections = clean_grouped_feature_sections
base.render_grouped_colour_summary = clean_render_grouped_colour_summary
base.model_colour_group_lines = clean_model_colour_group_lines
base.trim_colour_group_lines = clean_trim_colour_group_lines
base.article_heading = article_heading_clean
base.html_document = clean_title_document
base.build_manifest_from_bindings = build_manifest_from_bindings_clean
base.build_model_and_comparison_records = build_model_and_comparison_records_clean
base.write_outputs = write_outputs_with_glossary
base.MatrixRow.label = property(matrixrow_label)


if __name__ == '__main__':
    raise SystemExit(base.main())
