from __future__ import annotations

from collections import OrderedDict, defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
import hashlib
import html
import json
import re
import sys

from .utils import (
    MANIFEST_BODY_STYLE_TOKENS, MANIFEST_DRIVE_TOKENS, MANIFEST_STANDARDISH_CODES,
    bool_or_none, chunk_feature_items, chunk_list, first_unique, material_note_texts,
    normalize_text, numericish, render_article, sentence_chunks, short_slug, slugify,
    unique_output_path, unique_preserve_order,
)
from .models import BoundRecord, EngineAxleEntry, EngineAxleItem, GCWRRecord, ModelFeatureAggregate, OutputFileRecord, PowertrainTraileringGroup, SpecCell, SpecColumn, SpecGroupDoc, TraileringRecord, TrimDef, TrimFeatureAggregate, WorkbookData
from .parsing import parse_status_value, parse_value_and_footnote_ids, parse_workbook, referenced_codes_for_text
from .classification import COMPARISON_OBJECTTYPE, CONFIG_KIND_ENGINE_AXLE, CONFIG_KIND_GCWR, CONFIG_KIND_GCWR_REFERENCE, CONFIG_KIND_POWERTRAIN_TRAILERING_GROUP, CONFIG_KIND_SPEC_COLUMN, CONFIG_KIND_SPEC_GROUP, CONFIG_KIND_TRAILERING, CONFIG_OBJECTTYPE, CONFIG_TYPE, DOC_ROLE_CHILD, DOC_ROLE_ENTITY, DOC_ROLE_PARENT, DOC_ROLE_PASSAGE, DOMAIN_COLOR, DOMAIN_OVERVIEW, MODEL_OBJECTTYPE, NOTE_OBJECTTYPE, SURFACE_BOTH, SURFACE_PASSAGE_ONLY, TRIM_OBJECTTYPE, availability_pairs_for_model, availability_pairs_for_trim, category_for_model_feature, category_for_trim_feature, collect_row_note_texts, comparison_varies_by_trim, feature_title, model_status_summary_lines, normalize_domain_value, sort_category_key, sort_trim_feature, source_context, source_tab_list_from_contexts, source_tab_list_from_strings, summarize_model_status_groups, with_doc_metadata
from .configuration import all_trim_matches, all_trim_matches_for_spec_group, best_trim_match, best_trim_match_for_spec_column, column_matches_trim, group_powertrain_trailering_for_cpr, group_spec_columns_for_cpr, powertrain_group_trim_match, section_names_for_column, spec_column_body_style_value, spec_column_context_text, spec_column_drivetrain_value, spec_column_engine_value, spec_column_fuel_value, spec_column_seating_value, spec_group_context_text, spec_group_first_value, spec_group_model_code, spec_group_section_names, strip_drive_tokens, trim_body_styles, trim_code_list, trim_colour_context, trim_drivetrains, trim_header_list, trim_matches_decor, trim_name_list, trim_seating, workbook_tab_metadata


def aggregate_model_features(data: WorkbookData) -> List[ModelFeatureAggregate]:
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
            agg.availability_contexts[signature] = unique_preserve_order(agg.availability_contexts[signature] + [ctx])
            agg.notes = unique_preserve_order(agg.notes + collect_row_note_texts(row, sheet))
            agg.referenced_codes = list(OrderedDict(((code, desc), None) for code, desc in (agg.referenced_codes + referenced_codes_for_text(row.description_raw, data.glossary))).keys())
    return list(groups.values())

def aggregate_trim_features(data: WorkbookData, trim: TrimDef) -> List[TrimFeatureAggregate]:
    groups: 'OrderedDict[str, TrimFeatureAggregate]' = OrderedDict()
    for sheet in data.matrix_sheets:
        for row in sheet.rows:
            raw = normalize_text(row.status_by_trim.get(trim.key))
            if not raw:
                continue
            _code, label, _notes = parse_status_value(raw, row.inline_footnotes, sheet.footnotes)
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
            agg.availability_contexts[(raw, label)] = unique_preserve_order(agg.availability_contexts[(raw, label)] + [ctx])
            agg.notes = unique_preserve_order(agg.notes + collect_row_note_texts(row, sheet))
            agg.referenced_codes = list(OrderedDict(((code, desc), None) for code, desc in (agg.referenced_codes + referenced_codes_for_text(row.description_raw, data.glossary))).keys())
    return sorted(groups.values(), key=sort_trim_feature)

def collect_referenced_codes_for_model(data: WorkbookData) -> List[str]:
    codes: List[str] = []
    for sheet in data.matrix_sheets:
        for row in sheet.rows:
            for code, _desc in referenced_codes_for_text(row.description_raw, data.glossary):
                codes.append(code)
            for code in [row.option_code, row.ref_code]:
                if normalize_text(code) in data.glossary:
                    codes.append(normalize_text(code))
    for sheet in data.color_sheets:
        for row in sheet.interior_rows:
            if normalize_text(row.seat_code) in data.glossary:
                codes.append(normalize_text(row.seat_code))
            for value in row.colors.values():
                if normalize_text(value) in data.glossary:
                    codes.append(normalize_text(value))
        for row in sheet.exterior_rows:
            if normalize_text(row.color_code) in data.glossary:
                codes.append(normalize_text(row.color_code))
    return unique_preserve_order(codes)

def referenced_glossary_codes_for_trim(data: WorkbookData, trim: TrimDef) -> List[str]:
    codes: List[str] = []
    for agg in aggregate_trim_features(data, trim):
        for code in [agg.orderable_code, agg.reference_code]:
            if code and code in data.glossary:
                codes.append(code)
        for code, _desc in agg.referenced_codes:
            if code in data.glossary:
                codes.append(code)
    for sheet in data.color_sheets:
        for row in sheet.interior_rows:
            if not trim_matches_decor(trim, row.decor_level):
                continue
            if row.seat_code and row.seat_code in data.glossary:
                codes.append(row.seat_code)
            for value in row.colors.values():
                if value and value in data.glossary:
                    codes.append(value)
        for row in sheet.exterior_rows:
            if row.color_code and row.color_code in data.glossary:
                codes.append(row.color_code)
    return unique_preserve_order(codes)

def render_note_articles(title_prefix: str, note_texts: Sequence[str], context_text: str, label: str = 'Guide note') -> str:
    parts: List[str] = []
    for note_index, note in enumerate(unique_preserve_order(note_texts), start=1):
        chunks = sentence_chunks(note, max_words=105)
        for chunk_index, chunk in enumerate(chunks, start=1):
            title = f'{title_prefix} | {label.lower()} {note_index}'
            if len(chunks) > 1:
                title += f' | part {chunk_index}'
            parts.append(
                render_article(
                    title,
                    [(label, chunk), ('Applies to source context', context_text)],
                )
            )
    return ''.join(parts)

def render_guide_context_section(data: WorkbookData, page_title: str, trim: Optional[TrimDef] = None) -> str:
    parts = [f'<section class="guide-context"><h2>{html.escape(page_title)} | Guide Context</h2>']
    trim_headers = [trim_def.raw_header for trim_def in data.trim_defs if normalize_text(trim_def.raw_header)]
    fields = [
        ('Source tabs', '; '.join(data.sheet_names)),
        ('Trim headers from guide', ' ; '.join(trim_headers)),
    ]
    if trim is not None:
        fields = [
            ('Trim name', trim.name),
            ('Trim code', trim.code),
            ('Trim header from guide', trim.raw_header),
            ('Source tabs', '; '.join(data.sheet_names)),
        ]
    parts.append(render_article(f'{page_title} | Guide structure', fields))
    parts.append('</section>')
    return ''.join(parts)

def render_spec_sections(data: WorkbookData, page_title: str, columns: List[SpecColumn]) -> str:
    if not columns:
        return ''
    parts = [f'<section class="spec-sections"><h2>{html.escape(page_title)} | Specifications and dimensions</h2>']
    for column in columns:
        grouped: 'OrderedDict[str, List[str]]' = OrderedDict()
        for cell in column.cells:
            grouped.setdefault(cell.section or 'Data', []).append(f'{cell.label}: {cell.value}')
        header_context = unique_preserve_order([x for x in [column.top_label, column.header] + column.header_lines if normalize_text(x)])
        header_text = ' | '.join(header_context)
        for section_name, values in grouped.items():
            value_chunks = chunk_list(values, max_words=95, max_items=7)
            for idx, value_chunk in enumerate(value_chunks, start=1):
                title = ' | '.join(x for x in [column.header or column.top_label, section_name] if normalize_text(x))
                if idx > 1:
                    title += f' | part {idx}'
                parts.append(
                    render_article(
                        title,
                        [('Source sheet', column.sheet_name), ('Column context', header_text)],
                        [('Guide values', value_chunk)],
                    )
                )
    parts.append('</section>')
    return ''.join(parts)

def render_glossary_section(page_title: str, glossary: Dict[str, str], limit_codes: Optional[Sequence[str]] = None) -> str:
    if not glossary:
        return ''
    if limit_codes is None:
        ordered_codes = list(glossary.keys())
    else:
        ordered_codes = [code for code in OrderedDict((normalize_text(code), None) for code in limit_codes) if code in glossary]
    if not ordered_codes:
        return ''
    parts = [f'<section class="glossary"><h2>{html.escape(page_title)} | Option code glossary</h2>']
    for code in ordered_codes:
        parts.append(render_article(f'Option code | {code}', [('Option code', code), ('Description from All tab', glossary[code])]))
    parts.append('</section>')
    return ''.join(parts)

def page_entity(data: WorkbookData, trim: Optional[TrimDef] = None) -> str:
    if trim is None:
        return data.vehicle_name
    return normalize_text(f'{data.vehicle_name} {strip_drive_tokens(trim.name)}')

def full_trim_heading(data: WorkbookData, trim: TrimDef) -> str:
    base = page_entity(data, trim)
    if trim.code and trim.code.lower() != trim.name.lower():
        return f'{base} ({trim.code})'
    return base

def article_heading(entity: str, base_title: str) -> str:
    base_title = normalize_text(base_title)
    if not base_title:
        return entity
    return f'{entity} | {base_title}'

def source_tabs_from_contexts(contexts: Sequence[str]) -> str:
    tabs: List[str] = []
    for context in contexts:
        context = normalize_text(context)
        if not context:
            continue
        first = normalize_text(context.split('|', 1)[0])
        if first:
            tabs.append(first)
    return '; '.join(unique_preserve_order(tabs))

def compact_text(text: str, max_words: int = 20) -> str:
    text = normalize_text(text)
    if not text:
        return ''
    words = text.split()
    if len(words) <= max_words:
        return text
    return ' '.join(words[:max_words]).rstrip(',;:') + '...'

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

def identity_fields(
    data: WorkbookData,
    trim: Optional[TrimDef] = None,
    *,
    category: str = '',
    source_context: str = '',
    source_tabs: str = '',
    extra_fields: Sequence[Tuple[str, str]] = (),
) -> List[Tuple[str, str]]:
    fields: List[Tuple[str, str]] = [('Vehicle', data.vehicle_name)]
    if trim is not None:
        fields.append(('Trim', trim.name))
        if trim.code:
            fields.append(('Trim code', trim.code))
    if category:
        fields.append(('Guide category', category))
    if source_context:
        fields.append(('Source context', source_context))
    if source_tabs:
        fields.append(('Source tabs', source_tabs))
    fields.extend(extra_fields)
    return dedupe_fields(fields)

def availability_lines_for_trim(agg: TrimFeatureAggregate) -> List[str]:
    lines: List[str] = []
    for (raw, label), contexts in agg.availability_contexts.items():
        context_text = '; '.join(contexts)
        if context_text:
            lines.append(f'{label} [{raw}] — {context_text}')
        else:
            lines.append(f'{label} [{raw}]')
    return lines

def availability_summary_for_trim(agg: TrimFeatureAggregate) -> str:
    parts: List[str] = []
    for (raw, label), contexts in agg.availability_contexts.items():
        tabs = source_tabs_from_contexts(contexts)
        bit = f'{label} [{raw}]'
        if tabs:
            bit += f' ({tabs})'
        parts.append(bit)
    return '; '.join(parts)

def availability_lines_for_model(agg: ModelFeatureAggregate) -> List[str]:
    lines: List[str] = []
    for signature, contexts in agg.availability_contexts.items():
        summary = ' ; '.join(model_status_summary_lines(signature))
        context_text = '; '.join(contexts)
        if context_text:
            lines.append(f'{summary} — {context_text}')
        else:
            lines.append(summary)
    return lines

def availability_summary_for_model(agg: ModelFeatureAggregate) -> str:
    parts: List[str] = []
    for signature, _contexts in agg.availability_contexts.items():
        parts.append(' ; '.join(model_status_summary_lines(signature)))
    return ' / '.join(parts)

def render_page_identity_section(data: WorkbookData, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    parts = [f'<section class="guide-context"><h2>{html.escape(entity)} | Vehicle identity and guide structure</h2>']
    trim_headers = [trim_def.raw_header for trim_def in data.trim_defs if normalize_text(trim_def.raw_header)]
    if trim is None:
        fields = [
            ('Vehicle', data.vehicle_name),
            ('Source tabs', '; '.join(data.sheet_names)),
            ('Trim headers from guide', ' ; '.join(trim_headers)),
        ]
        parts.append(render_article(article_heading(entity, 'Vehicle identity and guide structure'), dedupe_fields(fields)))
    else:
        fields = [
            ('Vehicle', data.vehicle_name),
            ('Trim', trim.name),
            ('Trim code', trim.code),
            ('Trim header from guide', trim.raw_header),
            ('Source tabs', '; '.join(data.sheet_names)),
        ]
        parts.append(render_article(article_heading(entity, 'Vehicle identity and guide structure'), dedupe_fields(fields)))
    parts.append('</section>')
    return ''.join(parts)

def render_matrix_legend_section(data: WorkbookData, page_title: str) -> str:
    if not data.matrix_sheets:
        return ''
    legend_text = normalize_text(data.matrix_sheets[0].legend_text)
    if not legend_text:
        return ''
    return '<section class="matrix-legend">' + render_article(
        article_heading(page_title, 'Matrix availability legend'),
        identity_fields(data, category='Other guide content', extra_fields=[('Legend from guide', legend_text)]),
    ) + '</section>'

def trim_group_line(agg: TrimFeatureAggregate) -> str:
    feature_text = compact_text(agg.description or agg.title, max_words=18)
    return f'{feature_text} — {availability_summary_for_trim(agg)}'

def model_group_line(agg: ModelFeatureAggregate) -> str:
    feature_text = compact_text(agg.description or agg.title, max_words=18)
    return f'{feature_text} — {availability_summary_for_model(agg)}'

def grouped_feature_sections(
    data: WorkbookData,
    features: Sequence[object],
    *,
    trim: Optional[TrimDef] = None,
    model_mode: bool = False,
) -> str:
    if not features:
        return ''
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    category_groups: Dict[str, List[object]] = OrderedDict()
    for feature in features:
        category = category_for_model_feature(feature) if model_mode else category_for_trim_feature(feature)
        category_groups.setdefault(category, []).append(feature)

    parts = [f'<section class="grouped-feature-passages"><h2>{html.escape(entity)} | Grouped feature passages from guide</h2>']
    for category in sorted(category_groups.keys(), key=sort_category_key):
        items = category_groups[category]
        lines = [model_group_line(item) if model_mode else trim_group_line(item) for item in items]
        source_tabs = source_tabs_from_contexts([ctx for item in items for ctx in getattr(item, 'source_contexts', [])])
        for idx, line_chunk in enumerate(chunk_feature_items(lines), start=1):
            title = f'{category} | grouped guide passage'
            if idx > 1:
                title += f' | part {idx}'
            parts.append(
                render_article(
                    article_heading(entity, title),
                    identity_fields(data, trim, category=category, source_tabs=source_tabs),
                    [('Feature lines from guide', line_chunk)],
                )
            )
    parts.append('</section>')
    return ''.join(parts)

def exact_model_feature_section(data: WorkbookData, features: Sequence[ModelFeatureAggregate]) -> str:
    if not features:
        return ''
    entity = data.vehicle_name
    parts = [f'<section class="exact-feature-records"><h2>{html.escape(entity)} | Exact feature records from guide</h2>']
    ordered = sorted(features, key=lambda agg: (sort_category_key(category_for_model_feature(agg)), normalize_text(agg.title).lower()))
    for agg in ordered:
        category = category_for_model_feature(agg)
        fields = identity_fields(
            data,
            category=category,
            source_context='; '.join(agg.source_contexts),
        )
        fields.append(('Guide text', agg.description))
        bullet_groups: List[Tuple[str, Sequence[str]]] = []
        availability_lines = availability_lines_for_model(agg)
        if availability_lines:
            bullet_groups.append(('Availability by trim', availability_lines))
        if agg.notes:
            bullet_groups.append(('Guide notes', agg.notes))
        parts.append(render_article(article_heading(entity, agg.title), fields, bullet_groups))
    parts.append('</section>')
    return ''.join(parts)

def exact_trim_feature_section(data: WorkbookData, trim: TrimDef, features: Sequence[TrimFeatureAggregate]) -> str:
    if not features:
        return ''
    entity = full_trim_heading(data, trim)
    parts = [f'<section class="exact-feature-records"><h2>{html.escape(entity)} | Exact feature records from guide</h2>']
    ordered = sorted(features, key=lambda agg: (sort_category_key(category_for_trim_feature(agg)), sort_trim_feature(agg)))
    for agg in ordered:
        category = category_for_trim_feature(agg)
        fields = identity_fields(
            data,
            trim,
            category=category,
            source_context='; '.join(agg.source_contexts),
        )
        fields.append(('Guide text', agg.description))
        bullet_groups: List[Tuple[str, Sequence[str]]] = []
        availability_lines = availability_lines_for_trim(agg)
        if availability_lines:
            bullet_groups.append(('Availability on this trim', availability_lines))
        if agg.notes:
            bullet_groups.append(('Guide notes', agg.notes))
        parts.append(render_article(article_heading(entity, agg.title), fields, bullet_groups))
    parts.append('</section>')
    return ''.join(parts)

def render_model_feature_sections(data: WorkbookData) -> str:
    features = aggregate_model_features(data)
    return ''.join([
        grouped_feature_sections(data, features, model_mode=True),
        exact_model_feature_section(data, features),
    ])

def render_trim_feature_sections(data: WorkbookData, trim: TrimDef) -> str:
    features = aggregate_trim_features(data, trim)
    return ''.join([
        grouped_feature_sections(data, features, trim=trim, model_mode=False),
        exact_trim_feature_section(data, trim, features),
    ])

def render_model_color_sections(data: WorkbookData) -> str:
    if not data.color_sheets:
        return ''
    entity = data.vehicle_name
    parts = [f'<section class="colour-trim-sections"><h2>{html.escape(entity)} | Colour and trim from guide</h2>']
    for sheet in data.color_sheets:
        interior_lines: List[str] = []
        for row in sheet.interior_rows:
            color_lines = [f'{color}: {code}' for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--']
            summary = ' | '.join(x for x in [row.decor_level, row.seat_type, row.seat_trim] if normalize_text(x))
            if color_lines:
                summary += ' — ' + '; '.join(color_lines)
            interior_lines.append(summary)
        if interior_lines:
            for idx, chunk in enumerate(chunk_feature_items(interior_lines), start=1):
                title = 'Colour and trim | interior grouped passage'
                if idx > 1:
                    title += f' | part {idx}'
                parts.append(
                    render_article(
                        article_heading(entity, title),
                        identity_fields(data, category='Colour and trim', source_tabs=sheet.name),
                        [('Interior colour and trim lines from guide', chunk)],
                    )
                )
        exterior_lines: List[str] = []
        for row in sheet.exterior_rows:
            title_value, _title_note_ids = parse_value_and_footnote_ids(row.title)
            availability = [f'{color}: {status}' for color, status in row.colors.items() if normalize_text(status)]
            line = ' | '.join(x for x in [title_value or row.title, row.color_code] if normalize_text(x))
            if availability:
                line += ' — ' + '; '.join(availability)
            exterior_lines.append(line)
        if exterior_lines:
            for idx, chunk in enumerate(chunk_feature_items(exterior_lines), start=1):
                title = 'Colour and trim | exterior paint grouped passage'
                if idx > 1:
                    title += f' | part {idx}'
                parts.append(
                    render_article(
                        article_heading(entity, title),
                        identity_fields(data, category='Colour and trim', source_tabs=sheet.name),
                        [('Exterior paint lines from guide', chunk)],
                    )
                )
        for row in sheet.interior_rows:
            color_lines = [f'{color}: {code}' for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--']
            parts.append(
                render_article(
                    article_heading(entity, feature_title(f'Interior trim | {row.decor_level} | {row.seat_trim}', row.seat_code)),
                    identity_fields(
                        data,
                        category='Colour and trim',
                        source_tabs=sheet.name,
                        extra_fields=[
                            ('Decor level', row.decor_level),
                            ('Seat type', row.seat_type),
                            ('Seat trim', row.seat_trim),
                        ],
                    ),
                    [('Interior colours and guide values', color_lines)] if color_lines else [],
                )
            )
        for row in sheet.exterior_rows:
            availability_lines = [f'{color}: {status}' for color, status in row.colors.items() if normalize_text(status)]
            title_value, title_note_ids = parse_value_and_footnote_ids(row.title)
            note_texts = [sheet.footnotes[nid] for nid in title_note_ids if nid in sheet.footnotes]
            bullet_groups: List[Tuple[str, Sequence[str]]] = []
            if availability_lines:
                bullet_groups.append(('Availability by interior colour column', availability_lines))
            if note_texts:
                bullet_groups.append(('Guide notes', note_texts))
            parts.append(
                render_article(
                    article_heading(entity, feature_title(f'Exterior paint | {title_value or row.title}', row.color_code)),
                    identity_fields(
                        data,
                        category='Colour and trim',
                        source_tabs=sheet.name,
                        extra_fields=[('Touch-Up Paint Number', row.touch_up_paint_number)],
                    ),
                    bullet_groups,
                )
            )
        general_notes = unique_preserve_order(list(sheet.footnotes.values()) + list(sheet.bullet_notes))
        if general_notes:
            parts.append(
                render_article(
                    article_heading(entity, f'{sheet.name} | colour and trim notes'),
                    identity_fields(data, category='Colour and trim', source_tabs=sheet.name),
                    [('Guide notes', general_notes)],
                )
            )
    parts.append('</section>')
    return ''.join(parts)

def render_trim_color_sections(data: WorkbookData, trim: TrimDef) -> str:
    if not data.color_sheets:
        return ''
    entity = full_trim_heading(data, trim)
    parts = [f'<section class="trim-colours"><h2>{html.escape(entity)} | Colour and trim from guide</h2>']
    for sheet in data.color_sheets:
        grouped_interior_lines: List[str] = []
        relevant_interior_columns: List[str] = []
        for row in sheet.interior_rows:
            if not trim_matches_decor(trim, row.decor_level):
                continue
            color_lines = [f'{color}: {code}' for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--']
            relevant_interior_columns.extend([color for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--'])
            summary = ' | '.join(x for x in [row.decor_level, row.seat_type, row.seat_trim] if normalize_text(x))
            if color_lines:
                summary += ' — ' + '; '.join(color_lines)
            grouped_interior_lines.append(summary)
            parts.append(
                render_article(
                    article_heading(entity, feature_title(f'Interior trim | {row.decor_level} | {row.seat_trim}', row.seat_code)),
                    identity_fields(
                        data,
                        trim,
                        category='Colour and trim',
                        source_tabs=sheet.name,
                        extra_fields=[
                            ('Decor level', row.decor_level),
                            ('Seat type', row.seat_type),
                            ('Seat trim', row.seat_trim),
                        ],
                    ),
                    [('Interior colours and guide values', color_lines)] if color_lines else [],
                )
            )
        relevant_interior_columns = unique_preserve_order(relevant_interior_columns)
        if grouped_interior_lines:
            for idx, chunk in enumerate(chunk_feature_items(grouped_interior_lines), start=1):
                title = 'Colour and trim | interior grouped passage'
                if idx > 1:
                    title += f' | part {idx}'
                parts.append(
                    render_article(
                        article_heading(entity, title),
                        identity_fields(data, trim, category='Colour and trim', source_tabs=sheet.name),
                        [('Interior colour and trim lines from guide', chunk)],
                    )
                )
        grouped_exterior_lines: List[str] = []
        for row in sheet.exterior_rows:
            availability_lines: List[str] = []
            for color_name in relevant_interior_columns:
                raw = normalize_text(row.colors.get(color_name))
                if not raw:
                    continue
                _code, label, _notes = parse_status_value(raw, {}, sheet.footnotes)
                availability_lines.append(f'{color_name}: {label} [{raw}]')
            if not availability_lines:
                continue
            title_value, title_note_ids = parse_value_and_footnote_ids(row.title)
            note_texts = [sheet.footnotes[nid] for nid in title_note_ids if nid in sheet.footnotes]
            grouped_line = ' | '.join(x for x in [title_value or row.title, row.color_code] if normalize_text(x))
            grouped_line += ' — ' + '; '.join(availability_lines)
            grouped_exterior_lines.append(grouped_line)
            bullet_groups: List[Tuple[str, Sequence[str]]] = [('Availability by interior colour', availability_lines)]
            if note_texts:
                bullet_groups.append(('Guide notes', note_texts))
            parts.append(
                render_article(
                    article_heading(entity, feature_title(f'Exterior paint | {title_value or row.title}', row.color_code)),
                    identity_fields(
                        data,
                        trim,
                        category='Colour and trim',
                        source_tabs=sheet.name,
                        extra_fields=[('Touch-Up Paint Number', row.touch_up_paint_number)],
                    ),
                    bullet_groups,
                )
            )
        if grouped_exterior_lines:
            for idx, chunk in enumerate(chunk_feature_items(grouped_exterior_lines), start=1):
                title = 'Colour and trim | exterior paint grouped passage'
                if idx > 1:
                    title += f' | part {idx}'
                parts.append(
                    render_article(
                        article_heading(entity, title),
                        identity_fields(data, trim, category='Colour and trim', source_tabs=sheet.name),
                        [('Exterior paint lines from guide', chunk)],
                    )
                )
        general_notes = unique_preserve_order(list(sheet.footnotes.values()) + list(sheet.bullet_notes))
        if general_notes:
            parts.append(
                render_article(
                    article_heading(entity, f'{sheet.name} | colour and trim notes'),
                    identity_fields(data, trim, category='Colour and trim', source_tabs=sheet.name),
                    [('Guide notes', general_notes)],
                )
            )
    parts.append('</section>')
    return ''.join(parts)

def render_spec_records(data: WorkbookData, columns: List[SpecColumn], *, trim: Optional[TrimDef] = None) -> str:
    if not columns:
        return ''
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    parts = [f'<section class="spec-sections"><h2>{html.escape(entity)} | Specifications and dimensions from guide</h2>']
    for column in columns:
        grouped: 'OrderedDict[str, List[str]]' = OrderedDict()
        for cell in column.cells:
            grouped.setdefault(cell.section or 'Data', []).append(f'{cell.label}: {cell.value}')
        header_context = unique_preserve_order([x for x in [column.top_label, column.header] + column.header_lines if normalize_text(x)])
        header_text = ' | '.join(header_context)
        for section_name, values in grouped.items():
            for idx, value_chunk in enumerate(chunk_list(values, max_words=110, max_items=8), start=1):
                title = f'Specifications and dimensions | {section_name}'
                if header_text:
                    title += f' | {column.header or column.top_label}'
                if idx > 1:
                    title += f' | part {idx}'
                parts.append(
                    render_article(
                        article_heading(entity, title),
                        identity_fields(
                            data,
                            trim,
                            category='Specifications and dimensions',
                            source_tabs=column.sheet_name,
                            extra_fields=[('Column context', header_text)],
                        ),
                        [('Guide values', value_chunk)],
                    )
                )
    parts.append('</section>')
    return ''.join(parts)

def render_trim_spec_sections(data: WorkbookData, trim: TrimDef) -> str:
    direct_columns = [column for column in data.spec_columns if column_matches_trim(column, trim)]
    if direct_columns:
        return render_spec_records(data, direct_columns, trim=trim)
    return render_spec_records(data, data.spec_columns, trim=trim)

def render_engine_axles_section(data: WorkbookData) -> str:
    if not data.engine_axle_entries:
        return ''
    entity = data.vehicle_name
    parts = [f'<section class="engine-axles"><h2>{html.escape(entity)} | Engine, axle and GVWR from guide</h2>']
    for entry in data.engine_axle_entries:
        grouped: 'OrderedDict[str, List[str]]' = OrderedDict()
        for item in entry.items:
            line = f'{item.name}: {item.status_label} [{item.raw_status}]'
            if item.notes:
                line += ' — ' + '; '.join(item.notes)
            grouped.setdefault(item.category or 'Guide values', []).append(line)
        for category, items in grouped.items():
            for idx, chunk in enumerate(chunk_list(items, max_words=110, max_items=8), start=1):
                title = f'Engine, axle and GVWR | {entry.model_code} | {entry.engine} | {category}'
                if idx > 1:
                    title += f' | part {idx}'
                parts.append(
                    render_article(
                        article_heading(entity, title),
                        identity_fields(
                            data,
                            category='Engine, axle and GVWR',
                            source_tabs=entry.sheet_name,
                            extra_fields=[('Top label', entry.top_label)],
                        ),
                        [('Guide values', chunk)],
                    )
                )
    parts.append('</section>')
    return ''.join(parts)

def render_trailering_section(data: WorkbookData) -> str:
    if not data.trailering_records and not data.gcwr_records:
        return ''
    entity = data.vehicle_name
    parts = [f'<section class="trailering"><h2>{html.escape(entity)} | Trailering and GCWR from guide</h2>']
    for record in data.trailering_records:
        bullet_groups: List[Tuple[str, Sequence[str]]] = []
        if record.note_text:
            bullet_groups.append(('Guide text', sentence_chunks(record.note_text, max_words=90)))
        if record.footnotes:
            bullet_groups.append(('Guide notes', record.footnotes))
        parts.append(
            render_article(
                article_heading(entity, f'Trailering and GCWR | {record.model_code} | {record.engine} | {record.axle_ratio}'),
                identity_fields(
                    data,
                    category='Trailering and GCWR',
                    source_tabs=record.sheet_name,
                    extra_fields=[
                        ('Rating type', record.rating_type),
                        ('Maximum trailer weight', record.max_trailer_weight),
                    ],
                ),
                bullet_groups,
            )
        )
    for record in data.gcwr_records:
        bullet_groups: List[Tuple[str, Sequence[str]]] = []
        if record.footnotes:
            bullet_groups.append(('Guide notes', record.footnotes))
        else:
            bullet_groups.append(('Guide values', [f'GCWR: {record.gcwr}']))
        parts.append(
            render_article(
                article_heading(entity, f'Trailering and GCWR | GCWR | {record.engine} | {record.gcwr}'),
                identity_fields(
                    data,
                    category='Trailering and GCWR',
                    source_tabs=record.sheet_name,
                    extra_fields=[('Table title', record.table_title), ('Axle ratio', record.axle_ratio)],
                ),
                bullet_groups,
            )
        )
    parts.append('</section>')
    return ''.join(parts)

def manifest_relpath(path: Path, base_dir: Path) -> str:
    import os
    try:
        return os.path.relpath(str(path), str(base_dir))
    except Exception:
        return str(path)

def standardish_trim_descriptions(data: WorkbookData, trim: TrimDef) -> List[str]:
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

def model_descriptions_standard_for_all_trims(data: WorkbookData) -> List[str]:
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
    return first_unique(v for v in values if normalize_text(v).lower().startswith('fuel,'))


def _is_engine_fuel_emission_desc(lower: str) -> bool:
    """Return True if the lowercased description is an engine, fuel, or ZEV emission row."""
    if lower.startswith(('engine,', 'moteur,', 'fuel,', 'essence,', 'carburant,')):
        return True
    if ('emission' in lower or 'émission' in lower) and (
        'zero' in lower or 'zéro' in lower or 'zev' in lower or 'vze' in lower
    ):
        return True
    return False


def _all_engine_fuel_descriptions(data: WorkbookData) -> List[str]:
    """Collect all engine/fuel/emission descriptions from matrix rows (any trim status)."""
    values: List[str] = []
    for sheet in data.matrix_sheets:
        for row in sheet.rows:
            desc = normalize_text(row.description_main or row.description_raw)
            if desc and _is_engine_fuel_emission_desc(desc.lower()):
                values.append(desc)
    return unique_preserve_order(values)


def _trim_engine_fuel_descriptions(data: WorkbookData, trim: TrimDef) -> List[str]:
    """Collect engine/fuel/emission descriptions available on a specific trim (status != '--')."""
    values: List[str] = []
    for sheet in data.matrix_sheets:
        for row in sheet.rows:
            raw = normalize_text(row.status_by_trim.get(trim.key))
            if not raw:
                continue
            status_code, _, _ = parse_status_value(raw, row.inline_footnotes, sheet.footnotes)
            if status_code == '--':
                continue
            desc = normalize_text(row.description_main or row.description_raw)
            if desc and _is_engine_fuel_emission_desc(desc.lower()):
                values.append(desc)
    return unique_preserve_order(values)


def infer_fuel_types(descriptions: Sequence[str], vehicle_name: str = '') -> object:
    """Infer fuel type(s) from engine/fuel/emission descriptions.

    Returns a single string (e.g. 'Gasoline') or a sorted list
    (e.g. ['Diesel', 'Gasoline']) when multiple fuel types are detected.
    Returns None when nothing can be determined.
    """
    fuel_types: set = set()
    for desc in descriptions:
        lower = normalize_text(desc).lower()
        # Explicit fuel rows  (EN: Fuel,  FR: Essence, / Carburant,)
        if lower.startswith(('fuel,', 'essence,', 'carburant,')):
            if 'none' in lower or 'aucun' in lower:
                fuel_types.add('Electric')
            elif 'diesel' in lower:
                fuel_types.add('Diesel')
            elif 'gasoline' in lower or lower.startswith('essence,'):
                fuel_types.add('Gasoline')
            # else: skip non-informative rows (e.g. "additional fuel")
        # Engine / Moteur rows
        elif lower.startswith(('engine,', 'moteur,')):
            if 'none' in lower or 'aucun' in lower:
                fuel_types.add('Electric')
            elif 'diesel' in lower:
                fuel_types.add('Diesel')
            else:
                fuel_types.add('Gasoline')
        # ZEV / VZE emission rows
        elif ('emission' in lower or 'émission' in lower) and (
            'zero' in lower or 'zéro' in lower or 'zev' in lower or 'vze' in lower
        ):
            fuel_types.add('Electric')
    # Fallback: infer from vehicle name
    if not fuel_types:
        name_lower = vehicle_name.lower()
        if re.search(r'\bev\b', name_lower) or 'bolt' in name_lower:
            fuel_types.add('Electric')
        else:
            fuel_types.add('Gasoline')
    return sorted(fuel_types)

def pick_drivetrain_description(values: Sequence[str]) -> Optional[str]:
    direct = [v for v in values if 'wheel drive' in normalize_text(v).lower()]
    choice = first_unique(direct)
    if choice:
        return choice
    propulsion = [v for v in values if normalize_text(v).lower().startswith('propulsion,') and 'fwd' not in normalize_text(v).lower() and 'awd' not in normalize_text(v).lower() and 'rwd' not in normalize_text(v).lower() and '4wd' not in normalize_text(v).lower() and '2wd' not in normalize_text(v).lower()]
    choice = first_unique(propulsion)
    if choice:
        return choice
    tokenized = [v for v in values if normalize_text(v).lower().startswith('propulsion,')]
    return first_unique(tokenized)

def trim_direct_spec_columns(data: WorkbookData, trim: TrimDef) -> List[SpecColumn]:
    return [column for column in data.spec_columns if column_matches_trim(column, trim)]

def extract_trim_seating(data: WorkbookData, trim: TrimDef) -> Optional[str]:
    cols = trim_direct_spec_columns(data, trim)
    values = [
        cell.value
        for column in cols
        for cell in column.cells
        if normalize_text(cell.label).lower().startswith('seating capacity')
    ]
    return first_unique(values)

def extract_model_seating(data: WorkbookData) -> Optional[str]:
    values = [
        cell.value
        for column in data.spec_columns
        for cell in column.cells
        if normalize_text(cell.label).lower().startswith('seating capacity')
    ]
    return first_unique(values)

def extract_trim_body_style(data: WorkbookData, trim: TrimDef) -> Optional[str]:
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

def extract_model_body_style(data: WorkbookData) -> Optional[str]:
    top_labels = [column.top_label for column in data.spec_columns if normalize_text(column.top_label)]
    return first_unique(top_labels)

def extract_trim_drive_token_from_headers(data: WorkbookData, trim: TrimDef) -> Optional[str]:
    cols = trim_direct_spec_columns(data, trim)
    tokens: List[str] = []
    for column in cols:
        blob = '\n'.join([column.top_label, column.header] + column.header_lines)
        for token in MANIFEST_DRIVE_TOKENS:
            if token.lower() in blob.lower():
                tokens.append(token)
    return first_unique(tokens)

def extract_manifest_metadata_for_model(data: WorkbookData) -> Dict[str, object]:
    metadata: Dict[str, object] = {}
    seating_values = list(dict.fromkeys(
        sv
        for col in data.spec_columns
        if (sv := spec_column_seating_value(col))
    ))
    if seating_values:
        metadata['seating'] = seating_values
    fuel = infer_fuel_types(_all_engine_fuel_descriptions(data), data.vehicle_name)
    if fuel:
        metadata['fuel_type'] = fuel
    return metadata

def extract_manifest_metadata_for_trim(data: WorkbookData, trim: TrimDef) -> Dict[str, object]:
    metadata: Dict[str, object] = {}
    trim_name = normalize_text(trim.name)
    if trim_name:
        metadata['name'] = trim_name
    title = normalize_text(trim.raw_header)
    if title:
        metadata['title'] = title
    seating_values = trim_seating(data, trim)
    if seating_values:
        metadata['seating'] = seating_values
    fuel = infer_fuel_types(_trim_engine_fuel_descriptions(data, trim), data.vehicle_name)
    if fuel:
        metadata['fuel_type'] = fuel
    drivetrains = trim_drivetrains(data, trim)
    if drivetrains:
        metadata['drivetrains'] = drivetrains
    body_styles = trim_body_styles(data, trim)
    if body_styles:
        metadata['body_styles'] = body_styles
    return metadata

def render_trim_lineup_section(data: WorkbookData) -> str:
    headers = trim_header_list(data)
    if not headers:
        return ''
    entity = data.vehicle_name
    fields = identity_fields(
        data,
        category='Other guide content',
        source_tabs='; '.join(unique_preserve_order(sheet.name for sheet in data.matrix_sheets)),
        extra_fields=[('Trim headers from guide', ' ; '.join(headers))],
    )
    bullets: List[Tuple[str, Sequence[str]]] = []
    names = trim_name_list(data)
    codes = trim_code_list(data)
    if names:
        bullets.append(('Trim names from guide', names))
    if codes:
        bullets.append(('Trim codes from guide', codes))
    return '<section class="trim-lineup">' + render_article(
        article_heading(entity, 'Trim lineup from guide'),
        fields,
        bullets,
    ) + '</section>'

def render_model_page(data: WorkbookData) -> str:
    entity = data.vehicle_name
    parts = [
        '<html><head><meta charset="utf-8"></head><body>',
        f'<h1>{html.escape(entity)} | Vehicle Order Guide model page</h1>',
        render_page_identity_section(data),
        render_trim_lineup_section(data),
        render_matrix_legend_section(data, entity),
        render_model_feature_sections(data),
        render_model_color_sections(data),
        '</body></html>',
    ]
    return ''.join(part for part in parts if part)

def render_trim_page(data: WorkbookData, trim: TrimDef) -> str:
    entity = full_trim_heading(data, trim)
    parts = [
        '<html><head><meta charset="utf-8"></head><body>',
        f'<h1>{html.escape(entity)} | Vehicle Order Guide trim page</h1>',
        render_page_identity_section(data, trim=trim),
        render_matrix_legend_section(data, entity),
        render_trim_feature_sections(data, trim),
        render_trim_color_sections(data, trim),
        '</body></html>',
    ]
    return ''.join(part for part in parts if part)

def render_spec_column_page(data: WorkbookData, column: SpecColumn, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    header_context = spec_column_context_text(column)
    title = article_heading(entity, f'Configuration specifications | {column.sheet_name} | {column.header or column.top_label}')
    parts = [
        '<html><head><meta charset="utf-8"></head><body>',
        f'<h1>{html.escape(title)}</h1>',
    ]
    identity_extra = [
        ('Source tab', column.sheet_name),
        ('Configuration top label', column.top_label),
        ('Configuration header', column.header),
    ]
    if column.header_lines:
        identity_extra.append(('Configuration header lines', ' ; '.join(column.header_lines)))
    if header_context:
        identity_extra.append(('Configuration context', header_context))
    parts.append(
        '<section class="configuration-identity">' + render_article(
            article_heading(entity, 'Configuration identity from guide'),
            identity_fields(
                data,
                trim,
                category='Specifications and dimensions',
                source_tabs=column.sheet_name,
                extra_fields=identity_extra,
            ),
            [('Guide sections', section_names_for_column(column))] if section_names_for_column(column) else [],
        ) + '</section>'
    )

    grouped: 'OrderedDict[str, List[str]]' = OrderedDict()
    for cell in column.cells:
        grouped.setdefault(cell.section or 'Data', []).append(f'{cell.label}: {cell.value}')
    parts.append(f'<section class="configuration-spec-values"><h2>{html.escape(entity)} | Configuration specifications and dimensions from guide</h2>')
    for section_name, values in grouped.items():
        for idx, value_chunk in enumerate(chunk_list(values, max_words=110, max_items=8), start=1):
            section_title = f'{section_name} | guide values'
            if idx > 1:
                section_title += f' | part {idx}'
            parts.append(
                render_article(
                    article_heading(entity, section_title),
                    identity_fields(
                        data,
                        trim,
                        category='Specifications and dimensions',
                        source_tabs=column.sheet_name,
                        extra_fields=[('Configuration context', header_context)],
                    ),
                    [('Guide values', value_chunk)],
                )
            )
    parts.append('</section></body></html>')
    return ''.join(parts)

def render_engine_axle_page(data: WorkbookData, entry: EngineAxleEntry, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    title = article_heading(entity, f'Configuration engine, axle and GVWR | {entry.model_code} | {entry.engine}')
    parts = [
        '<html><head><meta charset="utf-8"></head><body>',
        f'<h1>{html.escape(title)}</h1>',
    ]
    identity_extra = [
        ('Source tab', entry.sheet_name),
        ('Top label', entry.top_label),
        ('Model code', entry.model_code),
        ('Engine', entry.engine),
    ]
    categories = unique_preserve_order(item.category for item in entry.items if normalize_text(item.category))
    parts.append(
        '<section class="configuration-identity">' + render_article(
            article_heading(entity, 'Configuration identity from guide'),
            identity_fields(
                data,
                trim,
                category='Engine, axle and GVWR',
                source_tabs=entry.sheet_name,
                extra_fields=identity_extra,
            ),
            [('Guide categories', categories)] if categories else [],
        ) + '</section>'
    )
    grouped: 'OrderedDict[str, List[str]]' = OrderedDict()
    note_groups: 'OrderedDict[str, List[str]]' = OrderedDict()
    for item in entry.items:
        grouped.setdefault(item.category or 'Guide values', []).append(f'{item.name}: {item.status_label} [{item.raw_status}]')
        if item.notes:
            note_groups.setdefault(item.category or 'Guide values', []).extend(item.notes)
    parts.append(f'<section class="configuration-engine-axle-values"><h2>{html.escape(entity)} | Engine, axle and GVWR from guide</h2>')
    for category, values in grouped.items():
        bullet_groups: List[Tuple[str, Sequence[str]]] = [('Guide values', values)]
        notes = unique_preserve_order(note_groups.get(category, []))
        if notes:
            bullet_groups.append(('Guide notes', notes))
        parts.append(
            render_article(
                article_heading(entity, f'{category} | guide values'),
                identity_fields(
                    data,
                    trim,
                    category='Engine, axle and GVWR',
                    source_tabs=entry.sheet_name,
                    extra_fields=[('Model code', entry.model_code), ('Engine', entry.engine)],
                ),
                bullet_groups,
            )
        )
    parts.append('</section></body></html>')
    return ''.join(parts)

def render_trailering_page(data: WorkbookData, record: TraileringRecord, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    title = article_heading(entity, f'Configuration trailering | {record.model_code} | {record.engine} | {record.axle_ratio}')
    parts = [
        '<html><head><meta charset="utf-8"></head><body>',
        f'<h1>{html.escape(title)}</h1>',
    ]
    identity_extra = [
        ('Source tab', record.sheet_name),
        ('Rating type', record.rating_type),
        ('Model code', record.model_code),
        ('Engine', record.engine),
        ('Axle ratio', record.axle_ratio),
        ('Maximum trailer weight', record.max_trailer_weight),
    ]
    parts.append(
        '<section class="configuration-identity">' + render_article(
            article_heading(entity, 'Configuration identity from guide'),
            identity_fields(
                data,
                trim,
                category='Trailering and GCWR',
                source_tabs=record.sheet_name,
                extra_fields=identity_extra,
            ),
            [('Guide text', sentence_chunks(record.note_text, max_words=90))] if record.note_text else [],
        ) + '</section>'
    )
    bullet_groups: List[Tuple[str, Sequence[str]]] = [('Guide values', [f'Axle ratio: {record.axle_ratio}', f'Maximum trailer weight: {record.max_trailer_weight}'])]
    if record.footnotes:
        bullet_groups.append(('Guide notes', record.footnotes))
    parts.append(
        '<section class="trailering-values">' + render_article(
            article_heading(entity, 'Trailering values from guide'),
            identity_fields(
                data,
                trim,
                category='Trailering and GCWR',
                source_tabs=record.sheet_name,
                extra_fields=[('Rating type', record.rating_type), ('Model code', record.model_code), ('Engine', record.engine)],
            ),
            bullet_groups,
        ) + '</section>'
    )
    parts.append('</body></html>')
    return ''.join(parts)

def render_gcwr_page(data: WorkbookData, records: Sequence[GCWRRecord], trim: Optional[TrimDef] = None) -> str:
    records = list(records)
    if not records:
        return ''
    engine = records[0].engine
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    title = article_heading(entity, f'Configuration GCWR | {engine}')
    parts = [
        '<html><head><meta charset="utf-8"></head><body>',
        f'<h1>{html.escape(title)}</h1>',
    ]
    source_tabs = '; '.join(unique_preserve_order(record.sheet_name for record in records))
    identity_extra = [
        ('Table title', records[0].table_title),
        ('Engine', engine),
    ]
    parts.append(
        '<section class="configuration-identity">' + render_article(
            article_heading(entity, 'Configuration identity from guide'),
            identity_fields(
                data,
                trim,
                category='Trailering and GCWR',
                source_tabs=source_tabs,
                extra_fields=identity_extra,
            ),
        ) + '</section>'
    )
    value_lines: List[str] = []
    note_lines: List[str] = []
    for record in records:
        value_lines.append(f'GCWR {record.gcwr}: axle ratio {record.axle_ratio}')
        note_lines.extend(record.footnotes)
    bullet_groups: List[Tuple[str, Sequence[str]]] = [('Guide values', value_lines)]
    deduped_notes = unique_preserve_order(note_lines)
    if deduped_notes:
        bullet_groups.append(('Guide notes', deduped_notes))
    parts.append(
        '<section class="gcwr-values">' + render_article(
            article_heading(entity, 'GCWR values from guide'),
            identity_fields(
                data,
                trim,
                category='Trailering and GCWR',
                source_tabs=source_tabs,
                extra_fields=[('Table title', records[0].table_title), ('Engine', engine)],
            ),
            bullet_groups,
        ) + '</section>'
    )
    parts.append('</body></html>')
    return ''.join(parts)

def model_manifest_metadata(data: WorkbookData) -> Dict[str, object]:
    metadata: Dict[str, object] = {}
    metadata.update(workbook_tab_metadata(data))
    metadata.update(extract_manifest_metadata_for_model(data))
    metadata.update({
        'name': data.vehicle_name,
        'title': data.vehicle_name,
        'year': normalize_text(data.year),
        'make': normalize_text(data.make),
        'model': normalize_text(data.model),
        'trim_headers_from_guide': trim_header_list(data),
        'trim_names_from_guide': trim_name_list(data),
        'trim_codes_from_guide': trim_code_list(data),
    })
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def trim_manifest_metadata(data: WorkbookData, trim: TrimDef) -> Dict[str, object]:
    metadata: Dict[str, object] = {}
    metadata.update(workbook_tab_metadata(data))
    metadata.update(extract_manifest_metadata_for_trim(data, trim))
    metadata.update({
        'name': strip_drive_tokens(normalize_text(trim.name)),
        'title': full_trim_heading(data, trim),
        'year': normalize_text(data.year),
        'make': normalize_text(data.make),
        'model': normalize_text(data.model),
        'trim': strip_drive_tokens(normalize_text(trim.name)),
        'trim_header_from_guide': normalize_text(trim.raw_header),
        'trim_code': normalize_text(trim.code),
    })
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def engine_axle_manifest_metadata(data: WorkbookData, entry: EngineAxleEntry, trim: Optional[TrimDef] = None) -> Dict[str, object]:
    metadata: Dict[str, object] = {
        'name': ' | '.join(unique_preserve_order([entry.model_code, entry.engine])),
        'title': article_heading(data.vehicle_name, f'Configuration engine, axle and GVWR | {entry.model_code} | {entry.engine}'),
        'source_tab': normalize_text(entry.sheet_name),
        'configuration_kind': CONFIG_KIND_ENGINE_AXLE,
        'top_label': normalize_text(entry.top_label),
        'model_code': normalize_text(entry.model_code),
        'engine': normalize_text(entry.engine),
        'guide_categories': unique_preserve_order(item.category for item in entry.items if normalize_text(item.category)),
    }
    if normalize_text(entry.top_label):
        metadata['body_style'] = normalize_text(entry.top_label)
    if trim is not None:
        metadata['trim_code'] = normalize_text(trim.code)
        metadata['trim_header_from_guide'] = normalize_text(trim.raw_header)
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def trailering_manifest_metadata(data: WorkbookData, record: TraileringRecord, trim: Optional[TrimDef] = None) -> Dict[str, object]:
    metadata: Dict[str, object] = {
        'name': ' | '.join(unique_preserve_order([record.model_code, record.engine, record.axle_ratio])),
        'title': article_heading(data.vehicle_name, f'Configuration trailering | {record.model_code} | {record.engine} | {record.axle_ratio}'),
        'source_tab': normalize_text(record.sheet_name),
        'configuration_kind': CONFIG_KIND_TRAILERING,
        'rating_type': normalize_text(record.rating_type),
        'model_code': normalize_text(record.model_code),
        'engine': normalize_text(record.engine),
        'axle_ratio': normalize_text(record.axle_ratio),
        'max_trailer_weight': normalize_text(record.max_trailer_weight),
    }
    if trim is not None:
        metadata['trim_code'] = normalize_text(trim.code)
        metadata['trim_header_from_guide'] = normalize_text(trim.raw_header)
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def gcwr_manifest_metadata(data: WorkbookData, records: Sequence[GCWRRecord], trim: Optional[TrimDef] = None) -> Dict[str, object]:
    records = list(records)
    metadata: Dict[str, object] = {
        'name': normalize_text(records[0].engine) if records else '',
        'title': article_heading(data.vehicle_name, f'Configuration GCWR | {records[0].engine}') if records else '',
        'source_tab': normalize_text(records[0].sheet_name) if records else '',
        'configuration_kind': CONFIG_KIND_GCWR,
        'table_title': normalize_text(records[0].table_title) if records else '',
        'engine': normalize_text(records[0].engine) if records else '',
        'gcwr_values': unique_preserve_order(record.gcwr for record in records),
        'axle_ratios': unique_preserve_order(record.axle_ratio for record in records),
    }
    if trim is not None:
        metadata['trim_code'] = normalize_text(trim.code)
        metadata['trim_header_from_guide'] = normalize_text(trim.raw_header)
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def vehicle_manifest_filename(data: WorkbookData) -> str:
    return f'manifest_{slugify(data.year)}_{slugify(data.make)}_{slugify(data.model)}.json'

def manifest_path_list(paths: Sequence[Path], base_dir: Path) -> List[str]:
    return [manifest_relpath(path, base_dir) for path in paths if path is not None]

def manifest_file_entry(
    record: OutputFileRecord,
    *,
    manifest_base: Path,
    collection: Optional[Path] = None,
    parent: Optional[Path] = None,
    child_paths: Optional[Sequence[Path]] = None,
    parent_vehicle: Optional[Path] = None,
    parent_trim: Optional[Path] = None,
) -> Dict[str, object]:
    entry: Dict[str, object] = {
        'objecttype': record.objecttype,
        'type': record.type,
    }
    entry.update(record.metadata)
    if collection is not None:
        entry['collection'] = manifest_relpath(collection, manifest_base)
    if parent is not None:
        entry['parent'] = manifest_relpath(parent, manifest_base)
    if child_paths:
        entry['child'] = manifest_path_list(child_paths, manifest_base)
    if parent_vehicle is not None:
        entry['parent_vehicle'] = manifest_relpath(parent_vehicle, manifest_base)
    if parent_trim is not None:
        entry['parent_trim'] = manifest_relpath(parent_trim, manifest_base)
    entry['path'] = manifest_relpath(record.path, manifest_base)
    return {k: v for k, v in entry.items() if v not in ('', [], None)}

def build_configuration_files(data: WorkbookData, output_dir: Path, used_names: set[str], trim_paths: Dict[str, Path], model_path: Path) -> List[OutputFileRecord]:
    records: List[OutputFileRecord] = []

    for column in data.spec_columns:
        matched_trim = best_trim_match_for_spec_column(data, column)
        trim_slug = slugify(matched_trim.name) if matched_trim is not None else ''
        base_name = f'spec_{data.year}_{slugify(data.make)}_{slugify(data.model)}_{slugify(column.sheet_name)}'
        if trim_slug:
            base_name += f'_{trim_slug}'
        if slugify(column.header or column.top_label):
            base_name += f'_{slugify(column.header or column.top_label)}'
        path = unique_output_path(output_dir, base_name + '.html', used_names)
        path.write_text(render_spec_column_page(data, column, trim=matched_trim), encoding='utf-8')
        metadata = spec_column_manifest_metadata(data, column, trim=matched_trim)
        records.append(OutputFileRecord(objecttype=CONFIG_OBJECTTYPE, type=CONFIG_TYPE, path=path, metadata=metadata))

    for entry in data.engine_axle_entries:
        matched_trim = best_trim_match(data, entry.top_label, entry.model_code, entry.engine)
        trim_slug = slugify(matched_trim.name) if matched_trim is not None else ''
        base_name = f'config_{data.year}_{slugify(data.make)}_{slugify(data.model)}_engine_axles_{slugify(entry.model_code)}_{slugify(entry.engine)}'
        if trim_slug:
            base_name += f'_{trim_slug}'
        path = unique_output_path(output_dir, base_name + '.html', used_names)
        path.write_text(render_engine_axle_page(data, entry, trim=matched_trim), encoding='utf-8')
        metadata = engine_axle_manifest_metadata(data, entry, trim=matched_trim)
        records.append(OutputFileRecord(objecttype=CONFIG_OBJECTTYPE, type=CONFIG_TYPE, path=path, metadata=metadata))

    for record in data.trailering_records:
        matched_trim = best_trim_match(data, record.model_code, record.engine, record.rating_type)
        trim_slug = slugify(matched_trim.name) if matched_trim is not None else ''
        base_name = f'config_{data.year}_{slugify(data.make)}_{slugify(data.model)}_trailering_{slugify(record.model_code)}_{slugify(record.engine)}_{slugify(record.axle_ratio)}'
        if trim_slug:
            base_name += f'_{trim_slug}'
        path = unique_output_path(output_dir, base_name + '.html', used_names)
        path.write_text(render_trailering_page(data, record, trim=matched_trim), encoding='utf-8')
        metadata = trailering_manifest_metadata(data, record, trim=matched_trim)
        records.append(OutputFileRecord(objecttype=CONFIG_OBJECTTYPE, type=CONFIG_TYPE, path=path, metadata=metadata))

    gcwr_groups: 'OrderedDict[Tuple[str, str, str], List[GCWRRecord]]' = OrderedDict()
    for record in data.gcwr_records:
        gcwr_groups.setdefault((record.sheet_name, record.table_title, record.engine), []).append(record)
    for (_sheet_name, _table_title, engine), group in gcwr_groups.items():
        matched_trim = best_trim_match(data, group[0].table_title, engine)
        trim_slug = slugify(matched_trim.name) if matched_trim is not None else ''
        base_name = f'config_{data.year}_{slugify(data.make)}_{slugify(data.model)}_gcwr_{slugify(engine)}'
        if trim_slug:
            base_name += f'_{trim_slug}'
        path = unique_output_path(output_dir, base_name + '.html', used_names)
        path.write_text(render_gcwr_page(data, group, trim=matched_trim), encoding='utf-8')
        metadata = gcwr_manifest_metadata(data, group, trim=matched_trim)
        records.append(OutputFileRecord(objecttype=CONFIG_OBJECTTYPE, type=CONFIG_TYPE, path=path, metadata=metadata))

    return records

def html_document(title: str, *body_parts: str) -> str:
    parts = ['<html><head><meta charset="utf-8"></head><body>', f'<h1>{html.escape(normalize_text(title))}</h1>']
    parts.extend(part for part in body_parts if part)
    parts.append('</body></html>')
    return ''.join(parts)

def config_parent_domain_for_kind(kind: str, source_tab: str = '') -> str:
    source_tab = normalize_text(source_tab)
    if kind == CONFIG_KIND_ENGINE_AXLE:
        return 'Powertrain'
    if kind == CONFIG_KIND_TRAILERING:
        return 'Trailering'
    if kind == CONFIG_KIND_GCWR:
        return 'Trailering'
    if kind == CONFIG_KIND_SPEC_COLUMN:
        if source_tab.lower().startswith('dimensions'):
            return 'Dimensions'
        if source_tab.lower().startswith('specs'):
            return 'Specifications'
    return 'Configuration'

def spec_cell_domain(column: SpecColumn, cell: SpecCell) -> str:
    blob = ' '.join([column.sheet_name, column.top_label, column.header, cell.section, cell.label]).lower()
    if 'trail' in blob or 'gcwr' in blob:
        return 'Trailering'
    if any(token in blob for token in ['engine', 'transmission', 'drivetrain', 'drive unit', 'propulsion', 'fuel', 'axle']):
        return 'Powertrain'
    if cell.section.lower().startswith('capacities') or any(token in blob for token in ['payload', 'gvwr', 'seating', 'tank', 'capacity', 'cargo volume']):
        return 'Capacities'
    if column.sheet_name.lower().startswith('dimensions') or any(token in blob for token in ['wheelbase', 'height', 'width', 'length', 'ground clearance', 'track', 'overhang', 'head room', 'leg room', 'shoulder room']):
        return 'Dimensions'
    return config_parent_domain_for_kind(CONFIG_KIND_SPEC_COLUMN, column.sheet_name)

def trim_feature_groups_by_category(data: WorkbookData, trim: TrimDef) -> 'OrderedDict[str, List[TrimFeatureAggregate]]':
    groups: 'OrderedDict[str, List[TrimFeatureAggregate]]' = OrderedDict()
    for feature in aggregate_trim_features(data, trim):
        groups.setdefault(category_for_trim_feature(feature), []).append(feature)
    ordered = OrderedDict()
    for category in sorted(groups.keys(), key=sort_category_key):
        ordered[category] = groups[category]
    return ordered

def model_feature_groups_by_category(data: WorkbookData) -> 'OrderedDict[str, List[ModelFeatureAggregate]]':
    groups: 'OrderedDict[str, List[ModelFeatureAggregate]]' = OrderedDict()
    for feature in aggregate_model_features(data):
        groups.setdefault(category_for_model_feature(feature), []).append(feature)
    ordered = OrderedDict()
    for category in sorted(groups.keys(), key=sort_category_key):
        ordered[category] = groups[category]
    return ordered

def model_colour_group_lines(data: WorkbookData) -> List[str]:
    lines: List[str] = []
    for sheet in data.color_sheets:
        for row in sheet.interior_rows:
            color_lines = [f'{color}: {code}' for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--']
            summary = ' | '.join(x for x in [sheet.name, row.decor_level, row.seat_type, row.seat_trim] if normalize_text(x))
            if color_lines:
                summary += ' — ' + '; '.join(color_lines)
            lines.append(summary)
        for row in sheet.exterior_rows:
            title_value, _title_note_ids = parse_value_and_footnote_ids(row.title)
            availability_lines = [f'{color}: {status}' for color, status in row.colors.items() if normalize_text(status)]
            summary = ' | '.join(x for x in [sheet.name, title_value or row.title, row.color_code, row.touch_up_paint_number] if normalize_text(x))
            if availability_lines:
                summary += ' — ' + '; '.join(availability_lines)
            lines.append(summary)
    return unique_preserve_order(lines)

def trim_colour_group_lines(data: WorkbookData, trim: TrimDef) -> List[str]:
    ctx = trim_colour_context(data, trim)
    lines: List[str] = []
    for sheet, row, color_lines in ctx['interior_items']:
        summary = ' | '.join(x for x in [sheet.name, row.decor_level, row.seat_type, row.seat_trim] if normalize_text(x))
        if color_lines:
            summary += ' — ' + '; '.join(color_lines)
        lines.append(summary)
    for sheet, row, availability_lines, _note_texts in ctx['exterior_items']:
        title_value, _title_note_ids = parse_value_and_footnote_ids(row.title)
        summary = ' | '.join(x for x in [sheet.name, title_value or row.title, row.color_code, row.touch_up_paint_number] if normalize_text(x))
        if availability_lines:
            summary += ' — ' + '; '.join(availability_lines)
        lines.append(summary)
    return unique_preserve_order(lines)

def render_grouped_colour_summary(data: WorkbookData, trim: Optional[TrimDef] = None) -> str:
    if not data.color_sheets:
        return ''
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    lines = trim_colour_group_lines(data, trim) if trim is not None else model_colour_group_lines(data)
    if not lines:
        return ''
    title = article_heading(entity, 'Colour and trim grouped passages from guide')
    parts = [f'<section class="grouped-colour-passages"><h2>{html.escape(entity)} | Colour and trim grouped passages from guide</h2>']
    for idx, chunk in enumerate(chunk_feature_items(lines), start=1):
        chunk_title = title if idx == 1 else article_heading(entity, f'Colour and trim grouped passage | part {idx}')
        parts.append(
            render_article(
                chunk_title,
                identity_fields(data, trim, category=DOMAIN_COLOR, source_tabs='; '.join(sheet.name for sheet in data.color_sheets)),
                [('Colour and trim lines from guide', chunk)],
            )
        )
    parts.append('</section>')
    return ''.join(parts)

def render_model_overview_page(data: WorkbookData) -> str:
    entity = data.vehicle_name
    return html_document(
        f'{entity} | Vehicle Order Guide model overview',
        render_page_identity_section(data),
        render_trim_lineup_section(data),
        render_matrix_legend_section(data, entity),
        grouped_feature_sections(data, aggregate_model_features(data), model_mode=True),
        render_grouped_colour_summary(data),
    )

def render_trim_overview_page(data: WorkbookData, trim: TrimDef) -> str:
    entity = full_trim_heading(data, trim)
    return html_document(
        f'{entity} | Vehicle Order Guide trim overview',
        render_page_identity_section(data, trim=trim),
        render_matrix_legend_section(data, entity),
        grouped_feature_sections(data, aggregate_trim_features(data, trim), trim=trim, model_mode=False),
        render_grouped_colour_summary(data, trim=trim),
    )

def render_model_trim_lineup_page(data: WorkbookData) -> str:
    entity = data.vehicle_name
    trim_headers = [trim.raw_header for trim in data.trim_defs if normalize_text(trim.raw_header)]
    bullet_groups: List[Tuple[str, Sequence[str]]] = []
    if trim_headers:
        bullet_groups.append(('Trim headers from guide', trim_headers))
    if trim_name_list(data):
        bullet_groups.append(('Trim names from guide', trim_name_list(data)))
    if trim_code_list(data):
        bullet_groups.append(('Trim codes from guide', trim_code_list(data)))
    return html_document(
        f'{entity} | Trim lineup from guide',
        '<section class="model-lineup">' + render_article(
            article_heading(entity, 'Trim lineup from guide'),
            identity_fields(data, category=DOMAIN_OVERVIEW, source_tabs='; '.join(data.sheet_names)),
            bullet_groups,
        ) + '</section>',
    )

def render_trim_domain_page(
    data: WorkbookData,
    trim: TrimDef,
    category: str,
    features: Sequence[TrimFeatureAggregate],
    colour_lines: Sequence[str] = (),
    note_lines: Sequence[str] = (),
) -> str:
    entity = full_trim_heading(data, trim)
    tabs = source_tab_list_from_contexts([ctx for feature in features for ctx in feature.source_contexts])
    if colour_lines:
        tabs = unique_preserve_order(tabs + [sheet.name for sheet in data.color_sheets])
    parts = [f'<section class="trim-domain"><h2>{html.escape(entity)} | {html.escape(category)} from guide</h2>']
    grouped_lines = [trim_group_line(feature) for feature in features]
    grouped_lines.extend(list(colour_lines))
    for idx, chunk in enumerate(chunk_feature_items(grouped_lines), start=1):
        title = article_heading(entity, f'{category} | grouped guide passage')
        if idx > 1:
            title = article_heading(entity, f'{category} | grouped guide passage | part {idx}')
        parts.append(
            render_article(
                title,
                identity_fields(data, trim, category=category, source_tabs='; '.join(tabs)),
                [('Feature lines from guide', chunk)],
            )
        )
    if note_lines:
        parts.append(
            render_article(
                article_heading(entity, f'{category} | guide notes'),
                identity_fields(data, trim, category=category, source_tabs='; '.join(tabs)),
                [('Guide notes', note_lines)],
            )
        )
    parts.append('</section>')
    return html_document(f'{entity} | {category} from guide', ''.join(parts))

def render_trim_feature_page(data: WorkbookData, trim: TrimDef, agg: TrimFeatureAggregate, category: str) -> str:
    entity = full_trim_heading(data, trim)
    tabs = source_tab_list_from_contexts(agg.source_contexts)
    fields = identity_fields(
        data,
        trim,
        category=category,
        source_context='; '.join(agg.source_contexts),
        source_tabs='; '.join(tabs),
        extra_fields=[
            ('Guide text', agg.description),
            ('Orderable code', agg.orderable_code),
            ('Reference code', agg.reference_code),
        ],
    )
    bullet_groups: List[Tuple[str, Sequence[str]]] = []
    availability_lines = availability_lines_for_trim(agg)
    if availability_lines:
        bullet_groups.append(('Availability on this trim', availability_lines))
    if agg.notes:
        bullet_groups.append(('Guide notes', agg.notes))
    return html_document(
        article_heading(entity, agg.title),
        '<section class="trim-feature">' + render_article(article_heading(entity, agg.title), fields, bullet_groups) + '</section>',
    )

def render_trim_colour_record_page(data: WorkbookData, trim: TrimDef, payload: Dict[str, object]) -> str:
    entity = full_trim_heading(data, trim)
    kind = payload['kind']
    if kind == 'interior':
        sheet = payload['sheet']
        row = payload['row']
        color_lines = payload['color_lines']
        title = feature_title(f'Interior trim | {row.decor_level} | {row.seat_trim}', row.seat_code)
        fields = identity_fields(
            data,
            trim,
            category=DOMAIN_COLOR,
            source_tabs=sheet.name,
            extra_fields=[
                ('Decor level', row.decor_level),
                ('Seat type', row.seat_type),
                ('Seat code', row.seat_code),
                ('Seat trim', row.seat_trim),
            ],
        )
        bullets = [('Interior colours and guide values', color_lines)] if color_lines else []
    else:
        sheet = payload['sheet']
        row = payload['row']
        availability_lines = payload['availability_lines']
        note_texts = payload['note_texts']
        title_value, _ids = parse_value_and_footnote_ids(row.title)
        title = feature_title(f'Exterior paint | {title_value or row.title}', row.color_code)
        fields = identity_fields(
            data,
            trim,
            category=DOMAIN_COLOR,
            source_tabs=sheet.name,
            extra_fields=[
                ('Colour code', row.color_code),
                ('Touch-Up Paint Number', row.touch_up_paint_number),
            ],
        )
        bullets = []
        if availability_lines:
            bullets.append(('Availability by interior colour', availability_lines))
        if note_texts:
            bullets.append(('Guide notes', note_texts))
    return html_document(
        article_heading(entity, title),
        '<section class="trim-colour-record">' + render_article(article_heading(entity, title), fields, bullets) + '</section>',
    )

def render_note_page(
    data: WorkbookData,
    title: str,
    note_text: str,
    *,
    trim: Optional[TrimDef] = None,
    category: str = 'Guide note',
    source_tabs: Sequence[str] = (),
    source_context: str = '',
) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    fields = identity_fields(
        data,
        trim,
        category=category,
        source_context=source_context,
        source_tabs='; '.join(source_tabs),
        extra_fields=[('Guide note', note_text)],
    )
    return html_document(
        article_heading(entity, title),
        '<section class="guide-note">' + render_article(article_heading(entity, title), fields) + '</section>',
    )

def render_comparison_domain_page(data: WorkbookData, category: str, features: Sequence[ModelFeatureAggregate]) -> str:
    entity = data.vehicle_name
    lines = [model_group_line(feature) for feature in features]
    tabs = source_tab_list_from_contexts([ctx for feature in features for ctx in feature.source_contexts])
    parts = [f'<section class="comparison-domain"><h2>{html.escape(entity)} | {html.escape(category)} trim comparison from guide</h2>']
    for idx, chunk in enumerate(chunk_feature_items(lines), start=1):
        title = article_heading(entity, f'{category} | trim comparison grouped passage')
        if idx > 1:
            title = article_heading(entity, f'{category} | trim comparison grouped passage | part {idx}')
        parts.append(
            render_article(
                title,
                identity_fields(
                    data,
                    category=category,
                    source_tabs='; '.join(tabs),
                    extra_fields=[('Comparison axis', 'Trim')],
                ),
                [('Comparison lines from guide', chunk)],
            )
        )
    parts.append('</section>')
    return html_document(f'{entity} | {category} trim comparison', ''.join(parts))

def render_comparison_feature_page(data: WorkbookData, agg: ModelFeatureAggregate, category: str) -> str:
    entity = data.vehicle_name
    tabs = source_tab_list_from_contexts(agg.source_contexts)
    fields = identity_fields(
        data,
        category=category,
        source_context='; '.join(agg.source_contexts),
        source_tabs='; '.join(tabs),
        extra_fields=[
            ('Comparison axis', 'Trim'),
            ('Guide text', agg.description),
            ('Orderable code', agg.orderable_code),
            ('Reference code', agg.reference_code),
        ],
    )
    bullet_groups: List[Tuple[str, Sequence[str]]] = []
    availability_lines = availability_lines_for_model(agg)
    if availability_lines:
        bullet_groups.append(('Availability by trim', availability_lines))
    if agg.notes:
        bullet_groups.append(('Guide notes', agg.notes))
    return html_document(
        article_heading(entity, f'Trim comparison | {agg.title}'),
        '<section class="comparison-feature">' + render_article(article_heading(entity, f'Trim comparison | {agg.title}'), fields, bullet_groups) + '</section>',
    )

def render_spec_cell_detail_page(data: WorkbookData, column: SpecColumn, cell: SpecCell, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    title = article_heading(entity, f'Configuration detail | {cell.section} | {cell.label}')
    fields = identity_fields(
        data,
        trim,
        category=spec_cell_domain(column, cell),
        source_tabs=column.sheet_name,
        extra_fields=[
            ('Configuration top label', column.top_label),
            ('Configuration header', column.header),
            ('Configuration section', cell.section),
            ('Specification label', cell.label),
            ('Guide value', cell.value),
        ],
    )
    if column.header_lines:
        fields.append(('Configuration header lines', ' ; '.join(column.header_lines)))
    return html_document(title, '<section class="configuration-detail">' + render_article(title, fields) + '</section>')

def render_engine_axle_detail_page(data: WorkbookData, entry: EngineAxleEntry, item: EngineAxleItem, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    title = article_heading(entity, f'Configuration engine, axle and GVWR detail | {entry.model_code} | {entry.engine} | {item.name}')
    fields = identity_fields(
        data,
        trim,
        category='Powertrain',
        source_tabs=entry.sheet_name,
        extra_fields=[
            ('Top label', entry.top_label),
            ('Model code', entry.model_code),
            ('Engine', entry.engine),
            ('Guide category', item.category),
            ('Guide item', item.name),
            ('Availability', item.status_label),
            ('Availability raw value', item.raw_status),
        ],
    )
    bullet_groups: List[Tuple[str, Sequence[str]]] = []
    if item.notes:
        bullet_groups.append(('Guide notes', item.notes))
    return html_document(title, '<section class="configuration-detail">' + render_article(title, fields, bullet_groups) + '</section>')

def render_trailering_detail_page(data: WorkbookData, record: TraileringRecord, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    title = article_heading(entity, f'Configuration trailering detail | {record.model_code} | {record.engine} | {record.axle_ratio}')
    fields = identity_fields(
        data,
        trim,
        category='Trailering',
        source_tabs=record.sheet_name,
        extra_fields=[
            ('Rating type', record.rating_type),
            ('Model code', record.model_code),
            ('Engine', record.engine),
            ('Axle ratio', record.axle_ratio),
            ('Maximum trailer weight', record.max_trailer_weight),
            ('Guide note context', record.note_text),
        ],
    )
    bullet_groups: List[Tuple[str, Sequence[str]]] = []
    if record.footnotes:
        bullet_groups.append(('Guide notes', record.footnotes))
    return html_document(title, '<section class="configuration-detail">' + render_article(title, fields, bullet_groups) + '</section>')

def render_gcwr_detail_page(data: WorkbookData, record: GCWRRecord, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    title = article_heading(entity, f'Configuration GCWR detail | {record.engine} | {record.gcwr}')
    fields = identity_fields(
        data,
        trim,
        category='Trailering',
        source_tabs=record.sheet_name,
        extra_fields=[
            ('Table title', record.table_title),
            ('Engine', record.engine),
            ('GCWR', record.gcwr),
            ('Axle ratio', record.axle_ratio),
        ],
    )
    bullet_groups: List[Tuple[str, Sequence[str]]] = []
    if record.footnotes:
        bullet_groups.append(('Guide notes', record.footnotes))
    return html_document(title, '<section class="configuration-detail">' + render_article(title, fields, bullet_groups) + '</section>')

_MODEL_METADATA_CACHE: Dict[int, Dict[str, object]] = {}

_TRIM_METADATA_CACHE: Dict[Tuple[int, str], Dict[str, object]] = {}

def clear_metadata_caches() -> None:
    _MODEL_METADATA_CACHE.clear()
    _TRIM_METADATA_CACHE.clear()

def cached_model_metadata(data: WorkbookData) -> Dict[str, object]:
    key = id(data)
    cached = _MODEL_METADATA_CACHE.get(key)
    if cached is None:
        cached = model_manifest_metadata(data)
        _MODEL_METADATA_CACHE[key] = cached
    return dict(cached)

def cached_trim_metadata(data: WorkbookData, trim: TrimDef) -> Dict[str, object]:
    key = (id(data), trim.key)
    cached = _TRIM_METADATA_CACHE.get(key)
    if cached is None:
        cached = trim_manifest_metadata(data, trim)
        _TRIM_METADATA_CACHE[key] = cached
    return dict(cached)

def model_lineup_manifest_metadata(data: WorkbookData) -> Dict[str, object]:
    return with_doc_metadata(
        cached_model_metadata(data),
        name=f'{data.vehicle_name} | Trim lineup',
        title=f'{data.vehicle_name} | Trim lineup from guide',
        doc_type='model-trim-lineup',
        doc_role=DOC_ROLE_CHILD,
        entity_level='model',
        domain='Overview',
    )

def trim_domain_manifest_metadata(data: WorkbookData, trim: TrimDef, category: str, features: Sequence[TrimFeatureAggregate], colour_lines: Sequence[str], note_lines: Sequence[str]) -> Dict[str, object]:
    tabs = source_tab_list_from_contexts([ctx for feature in features for ctx in feature.source_contexts])
    if colour_lines:
        tabs = unique_preserve_order(tabs + [sheet.name for sheet in data.color_sheets])
    return with_doc_metadata(
        cached_trim_metadata(data, trim),
        name=f'{trim.name} | {category}',
        title=f'{full_trim_heading(data, trim)} | {category} from guide',
        doc_type='trim-domain',
        doc_role=DOC_ROLE_PARENT,
        entity_level='trim',
        domain=category,
        source_tabs=tabs,
        feature_count=len(features) + len(colour_lines),
        has_notes=bool_or_none(bool(note_lines) or any(feature.notes for feature in features)),
    )

def trim_feature_manifest_metadata(data: WorkbookData, trim: TrimDef, agg: TrimFeatureAggregate, category: str) -> Dict[str, object]:
    raw_values, labels = availability_pairs_for_trim(agg)
    return with_doc_metadata(
        cached_trim_metadata(data, trim),
        name=normalize_text(agg.title),
        title=article_heading(full_trim_heading(data, trim), agg.title),
        doc_type='trim-feature',
        doc_role=DOC_ROLE_CHILD,
        entity_level='trim',
        domain=category,
        source_tabs=source_tab_list_from_contexts(agg.source_contexts),
        source_contexts=list(agg.source_contexts),
        feature_text_from_guide=normalize_text(agg.description),
        orderable_code=normalize_text(agg.orderable_code),
        reference_code=normalize_text(agg.reference_code),
        guide_status_values=raw_values,
        guide_status_labels=labels,
        has_notes=bool_or_none(bool(agg.notes)),
    )

def trim_colour_manifest_metadata(data: WorkbookData, trim: TrimDef, payload: Dict[str, object]) -> Dict[str, object]:
    if payload['kind'] == 'interior':
        row = payload['row']
        return with_doc_metadata(
            cached_trim_metadata(data, trim),
            name=feature_title(f'Interior trim | {row.decor_level} | {row.seat_trim}', row.seat_code),
            title=article_heading(full_trim_heading(data, trim), feature_title(f'Interior trim | {row.decor_level} | {row.seat_trim}', row.seat_code)),
            doc_type='trim-feature',
            doc_role=DOC_ROLE_CHILD,
            entity_level='trim',
            domain=DOMAIN_COLOR,
            source_tabs=[payload['sheet'].name],
            decor_level=normalize_text(row.decor_level),
            seat_type=normalize_text(row.seat_type),
            seat_code=normalize_text(row.seat_code),
            seat_trim=normalize_text(row.seat_trim),
            colour_values=list(payload['color_lines']),
        )
    row = payload['row']
    title_value, _ids = parse_value_and_footnote_ids(row.title)
    return with_doc_metadata(
        cached_trim_metadata(data, trim),
        name=feature_title(f'Exterior paint | {title_value or row.title}', row.color_code),
        title=article_heading(full_trim_heading(data, trim), feature_title(f'Exterior paint | {title_value or row.title}', row.color_code)),
        doc_type='trim-feature',
        doc_role=DOC_ROLE_CHILD,
        entity_level='trim',
        domain=DOMAIN_COLOR,
        source_tabs=[payload['sheet'].name],
        colour_code=normalize_text(row.color_code),
        touch_up_paint_number=normalize_text(row.touch_up_paint_number),
        availability_lines=list(payload['availability_lines']),
        has_notes=bool_or_none(bool(payload['note_texts'])),
    )

def comparison_feature_manifest_metadata(data: WorkbookData, agg: ModelFeatureAggregate, category: str) -> Dict[str, object]:
    raw_values, labels = availability_pairs_for_model(agg)
    return with_doc_metadata(
        cached_model_metadata(data),
        name=normalize_text(agg.title),
        title=article_heading(data.vehicle_name, f'Trim comparison | {agg.title}'),
        type='comparison',
        doc_type='comparison-feature',
        doc_role=DOC_ROLE_CHILD,
        entity_level='comparison',
        domain=category,
        source_tabs=source_tab_list_from_contexts(agg.source_contexts),
        source_contexts=list(agg.source_contexts),
        comparison_axis='trim',
        feature_text_from_guide=normalize_text(agg.description),
        orderable_code=normalize_text(agg.orderable_code),
        reference_code=normalize_text(agg.reference_code),
        guide_status_values=raw_values,
        guide_status_labels=labels,
        availability_varies_by_trim=comparison_varies_by_trim(agg),
        has_notes=bool_or_none(bool(agg.notes)),
    )

def note_manifest_metadata(
    title: str,
    *,
    base_metadata: Dict[str, object],
    entity_level: str,
    domain: str,
    source_tabs: Sequence[str],
    parent_doc_type: str,
) -> Dict[str, object]:
    metadata = {
        'name': normalize_text(title),
        'title': normalize_text(title),
        'type': 'note',
        'doc_type': 'note',
        'doc_role': DOC_ROLE_CHILD,
        'entity_level': entity_level,
        'domain': domain,
        'source_tabs': list(source_tabs),
        'parent_doc_type': parent_doc_type,
    }
    merged = dict(base_metadata)
    merged.update(metadata)
    return {k: v for k, v in merged.items() if v not in ('', [], None)}

def config_detail_manifest_metadata(
    data: WorkbookData,
    *,
    base_metadata: Dict[str, object],
    title: str,
    domain: str,
    source_tabs: Sequence[str],
    trim: Optional[TrimDef],
    extra: Dict[str, object],
) -> Dict[str, object]:
    metadata: Dict[str, object] = {}
    if trim is not None:
        metadata.update(cached_trim_metadata(data, trim))
    else:
        metadata.update(cached_model_metadata(data))
    metadata.update(base_metadata)
    metadata.update(extra)
    metadata.update({
        'name': normalize_text(title),
        'title': normalize_text(title),
        'doc_type': 'configuration-detail',
        'doc_role': DOC_ROLE_CHILD,
        'entity_level': 'configuration',
        'domain': domain,
        'source_tabs': list(source_tabs),
        'has_numeric_specs': bool_or_none(any(numericish(str(v)) for v in extra.values() if isinstance(v, str))),
    })
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def add_bound_record(
    bindings: List[BoundRecord],
    record: OutputFileRecord,
    *,
    collection: Optional[Path],
    parent: Optional[Path],
    parent_vehicle: Optional[Path],
    parent_trims: List[Path],
) -> None:
    bindings.append(
        BoundRecord(
            record=record,
            collection=collection,
            parent=parent,
            parent_vehicle=parent_vehicle,
            parent_trims=parent_trims,
        )
    )

def write_note_records(
    data: WorkbookData,
    output_dir: Path,
    used_names: set[str],
    bindings: List[BoundRecord],
    *,
    model_path: Path,
    parent_path: Path,
    parent_vehicle: Path,
    parent_trims: List[Path],
    title_prefix: str,
    notes: Sequence[str],
    base_metadata: Dict[str, object],
    entity_level: str,
    domain: str,
    source_tabs: Sequence[str],
    trim: Optional[TrimDef],
    parent_doc_type: str,
    stem_bits: Sequence[str],
    source_context: str = '',
) -> None:
    for note_index, note in enumerate(material_note_texts(notes), start=1):
        for chunk_index, chunk in enumerate(sentence_chunks(note, max_words=105), start=1):
            title = f'{title_prefix} | guide note {note_index}'
            if len(sentence_chunks(note, max_words=105)) > 1:
                title += f' | part {chunk_index}'
            filename = 'note_' + '_'.join(short_slug(bit, 28) for bit in stem_bits if normalize_text(bit)) + f'_{note_index}'
            if chunk_index > 1:
                filename += f'_{chunk_index}'
            note_path = unique_output_path(output_dir, filename + '.html', used_names)
            note_path.write_text(
                render_note_page(
                    data,
                    title,
                    chunk,
                    trim=trim,
                    category=domain,
                    source_tabs=source_tabs,
                    source_context=source_context,
                ),
                encoding='utf-8',
            )
            note_record = OutputFileRecord(
                objecttype=NOTE_OBJECTTYPE,
                type='note',
                path=note_path,
                metadata=note_manifest_metadata(
                    title,
                    base_metadata=base_metadata,
                    entity_level=entity_level,
                    domain=domain,
                    source_tabs=source_tabs,
                    parent_doc_type=parent_doc_type,
                ),
            )
            add_bound_record(
                bindings,
                note_record,
                collection=model_path,
                parent=parent_path,
                parent_vehicle=parent_vehicle,
                parent_trims=parent_trims,
            )

def vehicle_key(data: WorkbookData) -> str:
    return slugify(f'{data.year} {data.make} {data.model}').lower()

def trim_entity_key(data: WorkbookData, trim: TrimDef) -> str:
    return slugify(f'{data.year} {data.make} {data.model} {trim.name}').lower()

def domains_for_model_doc(data: WorkbookData) -> List[str]:
    domains = list(model_feature_groups_by_category(data).keys())
    if data.color_sheets:
        domains.append(DOMAIN_COLOR)
    domains = unique_preserve_order(normalize_domain_value(domain) for domain in domains)
    return [domain for domain in domains if normalize_text(domain)]

def domains_for_trim_doc(data: WorkbookData, trim: TrimDef) -> List[str]:
    domains = list(trim_feature_groups_by_category(data, trim).keys())
    ctx = trim_colour_context(data, trim)
    if ctx['interior_items'] or ctx['exterior_items'] or ctx['domain_notes']:
        domains.append(DOMAIN_COLOR)
    domains = unique_preserve_order(normalize_domain_value(domain) for domain in domains)
    return [domain for domain in domains if normalize_text(domain)]

def with_flat_doc_metadata(metadata: Dict[str, object], **extra: object) -> Dict[str, object]:
    combined = dict(metadata)
    combined.update(extra)
    return {k: v for k, v in combined.items() if v not in ('', [], None)}

def model_overview_manifest_metadata(data: WorkbookData) -> Dict[str, object]:
    return with_flat_doc_metadata(
        cached_model_metadata(data),
        doc_type='model',
        doc_role=DOC_ROLE_ENTITY,
        entity_level='model',
        domain='Overview',
        guide_domains=domains_for_model_doc(data),
        surface=SURFACE_BOTH,
        is_vehicle_entity=True,
        entity_name=data.vehicle_name,
        entity_key=vehicle_key(data),
        vehicle_key=vehicle_key(data),
    )

def trim_overview_manifest_metadata(data: WorkbookData, trim: TrimDef) -> Dict[str, object]:
    return with_flat_doc_metadata(
        cached_trim_metadata(data, trim),
        doc_type='trim',
        doc_role=DOC_ROLE_ENTITY,
        entity_level='trim',
        domain='Overview',
        guide_domains=domains_for_trim_doc(data, trim),
        surface=SURFACE_BOTH,
        is_vehicle_entity=True,
        entity_name=full_trim_heading(data, trim),
        entity_key=trim_entity_key(data, trim),
        vehicle_key=vehicle_key(data),
        trim_key=trim_entity_key(data, trim),
    )

def comparison_domain_manifest_metadata(data: WorkbookData, category: str, features: Sequence[ModelFeatureAggregate]) -> Dict[str, object]:
    return with_flat_doc_metadata(
        cached_model_metadata(data),
        name=f'{data.vehicle_name} | {category} trim comparison',
        title=f'{data.vehicle_name} | {category} trim comparison',
        type='comparison',
        doc_type='comparison',
        doc_role=DOC_ROLE_PASSAGE,
        entity_level='comparison',
        domain=normalize_domain_value(category),
        guide_domains=[normalize_domain_value(category)],
        source_tabs=source_tab_list_from_contexts([ctx for feature in features for ctx in feature.source_contexts]),
        comparison_axis='trim',
        feature_count=len(features),
        has_notes=bool_or_none(any(feature.notes for feature in features)),
        surface=SURFACE_PASSAGE_ONLY,
        is_vehicle_entity=None,
        entity_name=data.vehicle_name,
        entity_key=vehicle_key(data),
        vehicle_key=vehicle_key(data),
    )

def config_parent_manifest_metadata(base: Dict[str, object], *, domain: str, doc_type: str) -> Dict[str, object]:
    metadata = with_flat_doc_metadata(
        base,
        doc_type=doc_type,
        doc_role=DOC_ROLE_PASSAGE,
        entity_level='configuration',
        domain=normalize_domain_value(domain),
        guide_domains=[normalize_domain_value(domain)],
        source_tabs=source_tab_list_from_strings(base.get('source_tab', ''), base.get('source_tabs', [])),
        has_numeric_specs=bool_or_none(any(numericish(str(v)) for v in base.values() if isinstance(v, str))),
        surface=SURFACE_PASSAGE_ONLY,
        is_vehicle_entity=None,
    )
    return metadata

def build_trim_records(
    data: WorkbookData,
    output_dir: Path,
    used_names: set[str],
    bindings: List[BoundRecord],
    model_path: Path,
) -> Dict[str, Path]:
    trim_paths: Dict[str, Path] = {}
    # Keyed by canonical filename (before unique_output_path) to avoid _2.html
    # duplicates when the same trim name appears with different model codes across sheets.
    name_slug_to_path: Dict[str, Path] = {}
    for trim in data.trim_defs:
        overview_filename = f'trim_{data.year}_{slugify(data.make)}_{slugify(data.model)}_{slugify(trim.name)}.html'
        if overview_filename in name_slug_to_path:
            # Same trim name already has a page; reuse it so specs can still link here.
            trim_paths[trim.key] = name_slug_to_path[overview_filename]
            continue
        overview_path = unique_output_path(output_dir, overview_filename, used_names)
        overview_path.write_text(render_trim_overview_page(data, trim), encoding='utf-8')
        overview_record = OutputFileRecord(
            objecttype=TRIM_OBJECTTYPE,
            type='trim',
            path=overview_path,
            metadata=trim_overview_manifest_metadata(data, trim),
        )
        add_bound_record(bindings, overview_record, collection=model_path, parent=model_path, parent_vehicle=model_path, parent_trims=[])
        trim_paths[trim.key] = overview_path
        name_slug_to_path[overview_filename] = overview_path
    return trim_paths

def fold_safe_id(record: OutputFileRecord) -> str:
    doc_type = normalize_text(str(record.metadata.get('doc_type', record.type)))
    prefix_map = {
        'model': 'M',
        'trim': 'T',
        'comparison': 'C',
        'configuration-spec': 'S',
    }
    prefix = prefix_map.get(doc_type, 'D')
    basis = '|'.join([
        record.type,
        doc_type,
        normalize_text(str(record.metadata.get('entity_key', ''))),
        normalize_text(str(record.metadata.get('domain', ''))),
        record.path.name,
    ])
    digest = hashlib.sha1(basis.encode('utf-8')).hexdigest()[:16].upper()
    return prefix + digest

def build_manifest_from_bindings(data: WorkbookData, bindings: Sequence[BoundRecord], manifest_path: Path) -> Dict[str, object]:
    manifest_base = manifest_path.parent
    id_map: Dict[Path, str] = {binding.record.path: fold_safe_id(binding.record) for binding in bindings}
    children_map: Dict[Path, List[Path]] = defaultdict(list)
    for binding in bindings:
        if binding.parent is not None:
            children_map[binding.parent].append(binding.record.path)
        for trim_path in binding.parent_trims:
            if trim_path != binding.parent:
                children_map[trim_path].append(binding.record.path)
    manifest: Dict[str, object] = {
        'workbook': manifest_relpath(data.path, manifest_base),
        'vehicle_name': data.vehicle_name,
        'vehicle_key': vehicle_key(data),
        'year': normalize_text(data.year),
        'make': normalize_text(data.make),
        'model': normalize_text(data.model),
        'files': [],
    }
    for binding in bindings:
        record = binding.record
        entry: Dict[str, object] = {
            'objecttype': record.objecttype,
            'type': record.type,
            'node_id': id_map[record.path],
        }
        entry.update(record.metadata)
        if binding.collection is not None:
            entry['collection'] = id_map[binding.collection]
            entry['collection_path'] = manifest_relpath(binding.collection, manifest_base)
        if binding.parent is not None:
            entry['parent'] = id_map[binding.parent]
            entry['parent_path'] = manifest_relpath(binding.parent, manifest_base)
        child_paths = children_map.get(record.path, [])
        if child_paths:
            entry['child'] = [id_map[path] for path in child_paths]
            entry['child_paths'] = [manifest_relpath(path, manifest_base) for path in child_paths]
        if binding.parent_vehicle is not None:
            entry['parent_vehicle'] = manifest_relpath(binding.parent_vehicle, manifest_base)
            entry['parent_vehicle_id'] = id_map[binding.parent_vehicle]
        if binding.parent_trims:
            if len(binding.parent_trims) == 1:
                entry['parent_trim'] = manifest_relpath(binding.parent_trims[0], manifest_base)
                entry['parent_trim_id'] = id_map[binding.parent_trims[0]]
            else:
                entry['parent_trims'] = [manifest_relpath(p, manifest_base) for p in binding.parent_trims]
                entry['parent_trim_ids'] = [id_map[p] for p in binding.parent_trims]
        entry['path'] = manifest_relpath(record.path, manifest_base)
        entry = {k: v for k, v in entry.items() if v not in ('', [], None)}
        manifest['files'].append(entry)
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding='utf-8')
    return manifest

def manifest_text_is_meaningful(value: Optional[str], *, require_letters: bool = True) -> bool:
    text = normalize_text(value)
    if not text:
        return False
    if require_letters and len(re.sub(r'[^A-Za-z]', '', text)) < 2:
        return False
    if re.fullmatch(r'[0-9. ()/\-]+', text):
        return False
    return True

def looks_like_engine_value(value: Optional[str]) -> bool:
    text = normalize_text(value)
    if not manifest_text_is_meaningful(text):
        return False
    tokens = ['engine', 'motor', 'drive unit', 'electric drive', 'turbomax', 'turbo', 'diesel', 'v6', 'v8', 'l4', 'ecotec', 'duramax']
    return any(token in text.lower() for token in tokens)

def looks_like_fuel_value(value: Optional[str]) -> bool:
    text = normalize_text(value)
    if not manifest_text_is_meaningful(text):
        return False
    tokens = ['fuel', 'gasoline', 'diesel', 'electric', 'unleaded', 'ethanol', 'none']
    return any(token in text.lower() for token in tokens)

def spec_column_manifest_metadata(data: WorkbookData, column: SpecColumn, trim: Optional[TrimDef] = None) -> Dict[str, object]:
    metadata: Dict[str, object] = {
        'name': spec_column_context_text(column) or normalize_text(column.sheet_name),
        'title': article_heading(data.vehicle_name, f'Configuration specifications | {column.sheet_name} | {column.header or column.top_label}'),
        'source_tab': normalize_text(column.sheet_name),
        'configuration_kind': CONFIG_KIND_SPEC_COLUMN,
        'configuration_top_label': normalize_text(column.top_label),
        'configuration_header': normalize_text(column.header),
        'configuration_header_lines': list(column.header_lines),
        'guide_sections': section_names_for_column(column),
    }
    seating = spec_column_seating_value(column)
    if manifest_text_is_meaningful(seating, require_letters=False):
        metadata['seating'] = seating
    body_style = spec_column_body_style_value(column)
    if manifest_text_is_meaningful(body_style):
        metadata['body_style'] = body_style
    engine = spec_column_engine_value(column)
    if looks_like_engine_value(engine):
        metadata['engine'] = engine
    drivetrain = spec_column_drivetrain_value(column)
    if manifest_text_is_meaningful(drivetrain):
        metadata['drivetrain'] = drivetrain
    fuel = spec_column_fuel_value(column)
    if looks_like_fuel_value(fuel):
        metadata['fuel_type'] = fuel
    if trim is not None:
        metadata['trim_code'] = normalize_text(trim.code)
        metadata['trim_header_from_guide'] = normalize_text(trim.raw_header)
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def render_spec_group_page(data: WorkbookData, group: SpecGroupDoc, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    context_text = spec_group_context_text(group)
    model_code = spec_group_model_code(group)
    source_tabs = unique_preserve_order(column.sheet_name for column in group.columns)
    identity_extra: List[Tuple[str, str]] = []
    if model_code:
        identity_extra.append(('Model code', model_code))
    if group.top_label:
        identity_extra.append(('Configuration top label', group.top_label))
    if group.header:
        identity_extra.append(('Configuration header', group.header))
    if group.header_lines:
        identity_extra.append(('Configuration header lines', ' ; '.join(group.header_lines)))
    if context_text:
        identity_extra.append(('Configuration context', context_text))

    parts = [
        '<section class="configuration-identity">' + render_article(
            article_heading(entity, 'Configuration identity from guide'),
            identity_fields(
                data,
                trim,
                category='Dimensions and specifications',
                source_tabs='; '.join(source_tabs),
                extra_fields=identity_extra,
            ),
            [('Guide sections', spec_group_section_names(group))] if spec_group_section_names(group) else [],
        ) + '</section>'
    ]

    for column in group.columns:
        grouped_values: 'OrderedDict[str, List[str]]' = OrderedDict()
        for cell in column.cells:
            grouped_values.setdefault(cell.section or 'Data', []).append(f'{cell.label}: {cell.value}')
        for section_name, values in grouped_values.items():
            for idx, chunk in enumerate(chunk_list(values, max_words=125, max_items=10), start=1):
                title = article_heading(entity, f'{column.sheet_name} | {section_name} | guide values')
                if idx > 1:
                    title = article_heading(entity, f'{column.sheet_name} | {section_name} | guide values | part {idx}')
                extra_fields = [('Source tab', column.sheet_name)]
                if model_code:
                    extra_fields.append(('Model code', model_code))
                if group.header:
                    extra_fields.append(('Configuration header', group.header))
                parts.append(
                    '<section class="configuration-values">' + render_article(
                        title,
                        identity_fields(
                            data,
                            trim,
                            category='Dimensions and specifications',
                            source_tabs=column.sheet_name,
                            extra_fields=extra_fields,
                        ),
                        [('Guide values', chunk)],
                    ) + '</section>'
                )

    return html_document(
        article_heading(entity, f'Configuration dimensions and specifications | {group.header or group.top_label}'),
        ''.join(parts),
    )

def spec_group_manifest_metadata(data: WorkbookData, group: SpecGroupDoc, trim: Optional[TrimDef] = None) -> Dict[str, object]:
    model_code = spec_group_model_code(group)
    metadata: Dict[str, object] = {
        'name': spec_group_context_text(group) or group.top_label or group.header,
        'title': article_heading(data.vehicle_name, f'Configuration dimensions and specifications | {group.header or group.top_label}'),
        'source_tabs': unique_preserve_order(column.sheet_name for column in group.columns),
        'configuration_kind': CONFIG_KIND_SPEC_GROUP,
        'configuration_top_label': group.top_label,
        'configuration_header': group.header,
        'configuration_header_lines': list(group.header_lines),
        'guide_sections': spec_group_section_names(group),
    }
    if model_code:
        metadata['model_code'] = model_code
    body_style = spec_group_first_value(group, spec_column_body_style_value) or group.top_label
    if manifest_text_is_meaningful(body_style):
        metadata['body_style'] = body_style
    seating = spec_group_first_value(group, spec_column_seating_value)
    if manifest_text_is_meaningful(seating, require_letters=False):
        metadata['seating'] = seating
    engine = spec_group_first_value(group, spec_column_engine_value)
    if looks_like_engine_value(engine):
        metadata['engine'] = engine
    drivetrain = spec_group_first_value(group, spec_column_drivetrain_value)
    if manifest_text_is_meaningful(drivetrain):
        metadata['drivetrain'] = drivetrain
    fuel = spec_group_first_value(group, spec_column_fuel_value)
    if looks_like_fuel_value(fuel):
        metadata['fuel_type'] = fuel
    if trim is not None:
        metadata['trim_code'] = normalize_text(trim.code)
        metadata['trim_header_from_guide'] = normalize_text(trim.raw_header)
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def render_powertrain_trailering_group_page(data: WorkbookData, group: PowertrainTraileringGroup, trim: Optional[TrimDef] = None) -> str:
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    source_tabs = unique_preserve_order([entry.sheet_name for entry in group.engine_entries] + [record.sheet_name for record in group.trailering_records])
    engines = unique_preserve_order([entry.engine for entry in group.engine_entries] + [record.engine for record in group.trailering_records])
    identity_extra: List[Tuple[str, str]] = [('Model code', group.model_code)]
    if group.top_labels:
        identity_extra.append(('Configuration top labels', ' ; '.join(group.top_labels)))
    if engines:
        identity_extra.append(('Engines from guide', ' ; '.join(engines)))

    parts = [
        '<section class="configuration-identity">' + render_article(
            article_heading(entity, 'Configuration identity from guide'),
            identity_fields(
                data,
                trim,
                category='Powertrain and trailering',
                source_tabs='; '.join(source_tabs),
                extra_fields=identity_extra,
            ),
        ) + '</section>'
    ]

    engine_lines: List[str] = []
    for entry in group.engine_entries:
        for item in entry.items:
            line = f'Engine {entry.engine} | {item.category}: {item.status_label or item.raw_status}'
            if item.notes:
                line += f" | Notes: {' ; '.join(unique_preserve_order(item.notes))}"
            engine_lines.append(line)
    engine_lines = unique_preserve_order(engine_lines)
    if engine_lines:
        for idx, chunk in enumerate(chunk_list(engine_lines, max_words=125, max_items=10), start=1):
            title = article_heading(entity, 'Powertrain, axle and GVWR from guide')
            if idx > 1:
                title = article_heading(entity, f'Powertrain, axle and GVWR from guide | part {idx}')
            parts.append(
                '<section class="powertrain-values">' + render_article(
                    title,
                    identity_fields(
                        data,
                        trim,
                        category='Powertrain and trailering',
                        source_tabs='; '.join(unique_preserve_order(entry.sheet_name for entry in group.engine_entries)),
                        extra_fields=[('Model code', group.model_code)],
                    ),
                    [('Guide values', chunk)],
                ) + '</section>'
            )

    trailering_lines: List[str] = []
    rating_types: List[str] = []
    for record in group.trailering_records:
        line = f'{record.rating_type} | Engine {record.engine} | Axle ratio {record.axle_ratio} | Maximum trailer weight {record.max_trailer_weight}'
        if normalize_text(record.note_text):
            line += f' | Note heading: {record.note_text}'
        if record.footnotes:
            line += f" | Notes: {' ; '.join(unique_preserve_order(record.footnotes))}"
        trailering_lines.append(line)
        rating_types.append(record.rating_type)
    trailering_lines = unique_preserve_order(trailering_lines)
    if trailering_lines:
        for idx, chunk in enumerate(chunk_list(trailering_lines, max_words=125, max_items=8), start=1):
            title = article_heading(entity, 'Trailering values from guide')
            if idx > 1:
                title = article_heading(entity, f'Trailering values from guide | part {idx}')
            parts.append(
                '<section class="trailering-values">' + render_article(
                    title,
                    identity_fields(
                        data,
                        trim,
                        category='Powertrain and trailering',
                        source_tabs='; '.join(unique_preserve_order(record.sheet_name for record in group.trailering_records)),
                        extra_fields=[
                            ('Model code', group.model_code),
                            ('Trailering rating types', ' ; '.join(unique_preserve_order(rating_types))),
                        ],
                    ),
                    [('Guide values', chunk)],
                ) + '</section>'
            )

    return html_document(
        article_heading(entity, f'Configuration powertrain and trailering | {group.model_code}'),
        ''.join(parts),
    )

def powertrain_trailering_manifest_metadata(data: WorkbookData, group: PowertrainTraileringGroup, trim: Optional[TrimDef] = None) -> Dict[str, object]:
    engines = unique_preserve_order([entry.engine for entry in group.engine_entries] + [record.engine for record in group.trailering_records])
    axle_ratios = unique_preserve_order(record.axle_ratio for record in group.trailering_records if normalize_text(record.axle_ratio))
    max_trailer_weights = unique_preserve_order(record.max_trailer_weight for record in group.trailering_records if normalize_text(record.max_trailer_weight))
    guide_categories = unique_preserve_order(item.category for entry in group.engine_entries for item in entry.items if normalize_text(item.category))
    metadata: Dict[str, object] = {
        'name': group.model_code,
        'title': article_heading(data.vehicle_name, f'Configuration powertrain and trailering | {group.model_code}'),
        'source_tabs': unique_preserve_order([entry.sheet_name for entry in group.engine_entries] + [record.sheet_name for record in group.trailering_records]),
        'configuration_kind': CONFIG_KIND_POWERTRAIN_TRAILERING_GROUP,
        'model_code': group.model_code,
        'configuration_top_labels': unique_preserve_order(group.top_labels),
        'guide_categories': guide_categories,
        'trailering_rating_types': unique_preserve_order(record.rating_type for record in group.trailering_records if normalize_text(record.rating_type)),
    }
    if engines:
        metadata['engines'] = engines
    if axle_ratios:
        metadata['axle_ratios'] = axle_ratios
    if max_trailer_weights:
        metadata['max_trailer_weights'] = max_trailer_weights
    if trim is not None:
        metadata['trim_code'] = normalize_text(trim.code)
        metadata['trim_header_from_guide'] = normalize_text(trim.raw_header)
    return {k: v for k, v in metadata.items() if v not in ('', [], None)}

def render_gcwr_reference_page(data: WorkbookData, records: Sequence[GCWRRecord]) -> str:
    entity = data.vehicle_name
    records = list(records)
    source_tabs = unique_preserve_order(record.sheet_name for record in records)
    table_titles = unique_preserve_order(record.table_title for record in records)
    parts = [
        '<section class="gcwr-identity">' + render_article(
            article_heading(entity, 'GCWR identity from guide'),
            identity_fields(
                data,
                category='Trailering and GCWR',
                source_tabs='; '.join(source_tabs),
                extra_fields=[('Table titles', ' ; '.join(table_titles))],
            ),
        ) + '</section>'
    ]
    lines: List[str] = []
    for record in records:
        line = f'Engine {record.engine} | GCWR {record.gcwr} | Axle ratio {record.axle_ratio}'
        if record.footnotes:
            line += f" | Notes: {' ; '.join(unique_preserve_order(record.footnotes))}"
        lines.append(line)
    for idx, chunk in enumerate(chunk_list(unique_preserve_order(lines), max_words=125, max_items=10), start=1):
        title = article_heading(entity, 'GCWR values from guide')
        if idx > 1:
            title = article_heading(entity, f'GCWR values from guide | part {idx}')
        parts.append(
            '<section class="gcwr-values">' + render_article(
                title,
                identity_fields(
                    data,
                    category='Trailering and GCWR',
                    source_tabs='; '.join(source_tabs),
                    extra_fields=[('Table titles', ' ; '.join(table_titles))],
                ),
                [('Guide values', chunk)],
            ) + '</section>'
        )
    return html_document(article_heading(entity, 'GCWR reference from guide'), ''.join(parts))

def gcwr_reference_manifest_metadata(data: WorkbookData, records: Sequence[GCWRRecord]) -> Dict[str, object]:
    records = list(records)
    return {
        'name': f'{data.vehicle_name} GCWR reference',
        'title': article_heading(data.vehicle_name, 'GCWR reference from guide'),
        'source_tabs': unique_preserve_order(record.sheet_name for record in records),
        'configuration_kind': CONFIG_KIND_GCWR_REFERENCE,
        'table_titles': unique_preserve_order(record.table_title for record in records),
        'engines': unique_preserve_order(record.engine for record in records if normalize_text(record.engine)),
        'gcwr_values': unique_preserve_order(record.gcwr for record in records if normalize_text(record.gcwr)),
        'axle_ratios': unique_preserve_order(record.axle_ratio for record in records if normalize_text(record.axle_ratio)),
    }

def build_model_and_comparison_records(
    data: WorkbookData,
    output_dir: Path,
    used_names: set[str],
    bindings: List[BoundRecord],
) -> Path:
    model_filename = f'model_{data.year}_{slugify(data.make)}_{slugify(data.model)}.html'
    model_path = unique_output_path(output_dir, model_filename, used_names)
    model_path.write_text(render_model_overview_page(data), encoding='utf-8')
    model_record = OutputFileRecord(
        objecttype=MODEL_OBJECTTYPE,
        type='model',
        path=model_path,
        metadata=model_overview_manifest_metadata(data),
    )
    add_bound_record(bindings, model_record, collection=model_path, parent=None, parent_vehicle=None, parent_trims=[])

    for category, features in model_feature_groups_by_category(data).items():
        if normalize_domain_value(category) == 'Other guide content':
            continue
        selected_features = [feature for feature in features if comparison_varies_by_trim(feature)]
        if not selected_features:
            continue
        domain_filename = f'compare_domain_{data.year}_{slugify(data.make)}_{slugify(data.model)}_{short_slug(category)}.html'
        domain_path = unique_output_path(output_dir, domain_filename, used_names)
        domain_path.write_text(render_comparison_domain_page(data, category, selected_features), encoding='utf-8')
        domain_record = OutputFileRecord(
            objecttype=COMPARISON_OBJECTTYPE,
            type='comparison',
            path=domain_path,
            metadata=comparison_domain_manifest_metadata(data, category, selected_features),
        )
        add_bound_record(bindings, domain_record, collection=model_path, parent=model_path, parent_vehicle=model_path, parent_trims=[])
    return model_path

def build_configuration_records(
    data: WorkbookData,
    output_dir: Path,
    used_names: set[str],
    bindings: List[BoundRecord],
    model_path: Path,
    trim_paths: Dict[str, Path],
) -> None:
    trim_path_by_name = {normalize_text(trim.name): trim_paths[trim.key] for trim in data.trim_defs if trim.key in trim_paths}

    def parent_paths_for_trims(matched_trims: List[TrimDef]) -> Tuple[Path, List[Path]]:
        paths = [p for trim in matched_trims if (p := trim_path_by_name.get(normalize_text(trim.name))) is not None]
        if not paths:
            return model_path, []
        parent = paths[0] if len(paths) == 1 else model_path
        return parent, paths

    def enrich_config_metadata(metadata: Dict[str, object], matched_trim: Optional[TrimDef], domain: str) -> Dict[str, object]:
        extra: Dict[str, object] = {
            'entity_name': full_trim_heading(data, matched_trim) if matched_trim is not None else data.vehicle_name,
            'entity_key': trim_entity_key(data, matched_trim) if matched_trim is not None else vehicle_key(data),
            'vehicle_key': vehicle_key(data),
            'guide_domains': [normalize_domain_value(domain)],
            'retrieval_granularity': 'coarse',
        }
        if matched_trim is not None:
            extra['trim_key'] = trim_entity_key(data, matched_trim)
        return with_flat_doc_metadata(metadata, **extra)

    for group in group_spec_columns_for_cpr(data):
        matched_trims = all_trim_matches_for_spec_group(data, group)
        matched_trim = matched_trims[0] if len(matched_trims) == 1 else None
        trim_slug = slugify(matched_trim.name) if matched_trim is not None else ''
        model_code = spec_group_model_code(group)
        base_name = f'spec_{data.year}_{slugify(data.make)}_{slugify(data.model)}_dimensions_specifications'
        if trim_slug:
            base_name += f'_{trim_slug}'
        if model_code:
            base_name += f'_{slugify(model_code)}'
        elif slugify(group.header or group.top_label):
            base_name += f'_{short_slug(group.header or group.top_label)}'
        page_path = unique_output_path(output_dir, base_name + '.html', used_names)
        page_path.write_text(render_spec_group_page(data, group, trim=matched_trim), encoding='utf-8')
        domain = 'Dimensions and specifications'
        metadata = enrich_config_metadata(
            config_parent_manifest_metadata(
                spec_group_manifest_metadata(data, group, trim=matched_trim),
                domain=domain,
                doc_type='configuration-spec',
            ),
            matched_trim,
            domain,
        )
        record = OutputFileRecord(objecttype=CONFIG_OBJECTTYPE, type=CONFIG_TYPE, path=page_path, metadata=metadata)
        parent_path, parent_trims = parent_paths_for_trims(matched_trims)
        add_bound_record(bindings, record, collection=model_path, parent=parent_path, parent_vehicle=model_path, parent_trims=parent_trims)

    for group in group_powertrain_trailering_for_cpr(data):
        matched_trims = all_trim_matches(data, group.model_code, *group.top_labels)
        matched_trim = matched_trims[0] if len(matched_trims) == 1 else None
        trim_slug = slugify(matched_trim.name) if matched_trim is not None else ''
        base_name = f'config_{data.year}_{slugify(data.make)}_{slugify(data.model)}_powertrain_trailering_{slugify(group.model_code)}'
        if trim_slug:
            base_name += f'_{trim_slug}'
        page_path = unique_output_path(output_dir, base_name + '.html', used_names)
        page_path.write_text(render_powertrain_trailering_group_page(data, group, trim=matched_trim), encoding='utf-8')
        domain = 'Powertrain and trailering'
        metadata = enrich_config_metadata(
            config_parent_manifest_metadata(
                powertrain_trailering_manifest_metadata(data, group, trim=matched_trim),
                domain=domain,
                doc_type='configuration-spec',
            ),
            matched_trim,
            domain,
        )
        record = OutputFileRecord(objecttype=CONFIG_OBJECTTYPE, type=CONFIG_TYPE, path=page_path, metadata=metadata)
        parent_path, parent_trims = parent_paths_for_trims(matched_trims)
        add_bound_record(bindings, record, collection=model_path, parent=parent_path, parent_vehicle=model_path, parent_trims=parent_trims)

    if data.gcwr_records:
        page_path = unique_output_path(
            output_dir,
            f'config_{data.year}_{slugify(data.make)}_{slugify(data.model)}_gcwr_reference.html',
            used_names,
        )
        page_path.write_text(render_gcwr_reference_page(data, data.gcwr_records), encoding='utf-8')
        domain = 'Trailering and GCWR'
        metadata = enrich_config_metadata(
            config_parent_manifest_metadata(
                gcwr_reference_manifest_metadata(data, data.gcwr_records),
                domain=domain,
                doc_type='configuration-spec',
            ),
            None,
            domain,
        )
        output_record = OutputFileRecord(objecttype=CONFIG_OBJECTTYPE, type=CONFIG_TYPE, path=page_path, metadata=metadata)
        add_bound_record(bindings, output_record, collection=model_path, parent=model_path, parent_vehicle=model_path, parent_trims=[])
