"""Section renderers for model and trim pages."""

import html as html_module
from collections import OrderedDict
from typing import TYPE_CHECKING, Dict, List, Sequence, Tuple

from ..aggregators.feature_aggregator import (
    aggregate_model_features,
    aggregate_trim_features,
    category_for_model_feature,
    category_for_trim_feature,
)
from ..models.aggregate_models import ModelFeatureAggregate, TrimFeatureAggregate
from ..parsers.common import parse_value_and_footnote_ids
from ..utils.constants import STATUS_PRIORITY, TRIM_SECTION_ORDER
from ..utils.text_utils import chunk_list, compact_text, htmlize_text, normalize_text, unique_preserve_order
from .common import (
    article_heading,
    availability_lines_for_model,
    availability_lines_for_trim,
    availability_summary_for_model,
    availability_summary_for_trim,
    chunk_feature_items,
    column_matches_trim,
    dedupe_fields,
    full_trim_heading,
    identity_fields,
    render_article,
    sort_category_key,
    trim_matches_decor,
)

if TYPE_CHECKING:
    from ..models.base_models import SpecColumn, TrimDef, WorkbookData


def feature_title(label: str, orderable_code: str = '', reference_code: str = '') -> str:
    """Generate feature title with codes."""
    prefix = orderable_code or reference_code
    label = normalize_text(label)
    if prefix:
        return f'{prefix} | {label}'
    return label


def render_page_identity_section(data: 'WorkbookData', trim: 'TrimDef | None' = None) -> str:
    """Render vehicle/trim identity block."""
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    parts = [f'<section class="guide-context"><h2>{html_module.escape(entity)} | Vehicle information</h2>']

    fields: List[Tuple[str, str]] = [('Vehicle', data.vehicle_name)]
    if data.vehicle_type:
        fields.append(('Vehicle type', data.vehicle_type.upper()))
    if data.propulsion:
        fields.append(('Propulsion', data.propulsion))

    if trim is not None:
        fields.append(('Trim', trim.name))
        if trim.code:
            fields.append(('Trim code', trim.code))
        
        from ..inference import infer_trim_drive_type
        trim_drive = infer_trim_drive_type(data.spec_columns, trim, data.drive_types)
        if trim_drive:
            fields.append(('Drive type', trim_drive))
        elif data.drive_types:
            fields.append(('Drive types', ', '.join(data.drive_types)))
    else:
        if data.drive_types:
            fields.append(('Drive types', ', '.join(data.drive_types)))
        trims_text = ' | '.join(
            f'{t.name} ({t.code})' if t.code and t.code != t.name else t.name
            for t in data.trim_defs if t.name
        )
        if trims_text:
            fields.append(('Available trims', trims_text))

    parts.append(render_article(article_heading(entity, 'Vehicle information'), dedupe_fields(fields)))
    parts.append('</section>')
    return ''.join(parts)


def render_model_feature_sections(data: 'WorkbookData') -> str:
    """Render model-level features grouped by category."""
    features = aggregate_model_features(data)
    if not features:
        return ''
    
    entity = data.vehicle_name
    category_groups: Dict[str, List[ModelFeatureAggregate]] = OrderedDict()
    
    for agg in features:
        cat = category_for_model_feature(agg)
        category_groups.setdefault(cat, []).append(agg)

    parts = [f'<section class="model-features"><h2>{html_module.escape(entity)} | Features and availability</h2>']
    
    for category in sorted(category_groups.keys(), key=sort_category_key):
        cat_feats = category_groups[category]
        items: List[str] = []
        for agg in cat_feats:
            desc = compact_text(agg.description or agg.title, max_words=25)
            avail = availability_summary_for_model(agg)
            items.append(f'{desc} — {avail}')
        
        for idx, chunk in enumerate(chunk_feature_items(items), start=1):
            title = category
            if idx > 1:
                title += f' | part {idx}'
            parts.append(
                render_article(
                    article_heading(entity, title),
                    identity_fields(data),
                    [('Feature availability from guide', chunk)],
                )
            )
    
    parts.append('</section>')
    return ''.join(parts)


def render_trim_feature_sections(data: 'WorkbookData', trim: 'TrimDef') -> str:
    """Render trim features grouped by availability section."""
    features = aggregate_trim_features(data, trim)
    if not features:
        return ''
    
    entity = full_trim_heading(data, trim)

    # Best status label for sorting
    def _best_status_label(agg: TrimFeatureAggregate) -> str:
        if not agg.availability_contexts:
            return 'Other'
        return min(
            (label for _raw, label in agg.availability_contexts.keys()),
            key=lambda lbl: STATUS_PRIORITY.get(lbl, 99),
        )

    # Bucket: section_label → source_sheet → [agg, ...]
    section_buckets: Dict[str, Dict[str, List[TrimFeatureAggregate]]] = OrderedDict(
        (label, OrderedDict()) for label in TRIM_SECTION_ORDER
    )
    
    for agg in features:
        label = _best_status_label(agg)
        bucket = section_buckets.setdefault(label, OrderedDict())
        primary_src = (
            normalize_text(agg.source_contexts[0]).split(' | ')[0]
            if agg.source_contexts else 'Other'
        )
        bucket.setdefault(primary_src, []).append(agg)

    parts = [f'<section class="trim-features"><h2>{html_module.escape(entity)} | Equipment and features</h2>']
    
    for section_label in TRIM_SECTION_ORDER:
        by_source = section_buckets.get(section_label, {})
        if not by_source:
            continue
        
        for source_name, source_feats in by_source.items():
            items: List[str] = []
            for agg in source_feats:
                desc = normalize_text(agg.description or agg.title)
                code = agg.orderable_code or agg.reference_code
                line = f'({code}) {desc}' if code else desc
                if agg.notes:
                    line += f' — {agg.notes[0]}'
                items.append(line)
            
            for idx, chunk in enumerate(chunk_list(items, max_words=110, max_items=8), start=1):
                title = f'{section_label} | {source_name}'
                if idx > 1:
                    title += f' | part {idx}'
                parts.append(
                    render_article(
                        article_heading(entity, title),
                        identity_fields(data, trim),
                        [('Features', chunk)],
                    )
                )
    
    parts.append('</section>')
    return ''.join(parts)


def render_model_color_sections(data: 'WorkbookData') -> str:
    """Render color sections for model."""
    if not data.color_sheets:
        return ''
    
    entity = data.vehicle_name
    parts = [f'<section class="colour-trim-sections"><h2>{html_module.escape(entity)} | Colour and trim from guide</h2>']
    
    for sheet in data.color_sheets:
        # Interior colors
        if sheet.interior_rows:
            decor_groups: OrderedDict = OrderedDict()
            for row in sheet.interior_rows:
                color_pairs = [
                    (color_name, code)
                    for color_name, code in row.colors.items()
                    if normalize_text(code) and normalize_text(code) != '--'
                ]
                if not color_pairs:
                    continue
                decor_groups.setdefault(row.decor_level, []).append(
                    (row.seat_type, row.seat_trim, row.seat_code, color_pairs)
                )

            interior_group_lines: List[str] = []
            for decor_level, seat_rows in decor_groups.items():
                for seat_type, seat_trim, seat_code, color_pairs in seat_rows:
                    if len(color_pairs) == 1:
                        color_name, code = color_pairs[0]
                        line = f'{decor_level} | {seat_trim} — {color_name}: {code}'
                    else:
                        color_text = '; '.join(f'{cn}: {code}' for cn, code in color_pairs)
                        line = f'{decor_level} | {seat_trim} — {color_text}'
                    interior_group_lines.append(line)
                    
                    # Atomic record
                    color_items = [f'{cn}: {code}' for cn, code in color_pairs]
                    parts.append(
                        render_article(
                            article_heading(entity, feature_title(f'Interior trim | {decor_level} | {seat_trim}', seat_code)),
                            identity_fields(
                                data,
                                extra_fields=[
                                    ('Decor level', decor_level),
                                    ('Seat type', seat_type),
                                    ('Seat trim', seat_trim),
                                ],
                            ),
                            [('Interior colours and guide values', color_items)],
                        )
                    )

            if interior_group_lines:
                for idx, chunk in enumerate(chunk_feature_items(interior_group_lines), start=1):
                    title = 'Colour and trim | interior grouped passage'
                    if idx > 1:
                        title += f' | part {idx}'
                    parts.append(
                        render_article(
                            article_heading(entity, title),
                            identity_fields(data),
                            [('Interior colour and trim lines from guide', chunk)],
                        )
                    )

        # Exterior colors
        if sheet.exterior_rows:
            exterior_group_lines: List[str] = []
            for row in sheet.exterior_rows:
                available_interiors = [
                    color_name
                    for color_name, status in row.colors.items()
                    if normalize_text(status) and normalize_text(status).upper() in ('A', 'S')
                ]
                title_value, title_note_ids = parse_value_and_footnote_ids(row.title)
                note_texts = [sheet.footnotes[nid] for nid in title_note_ids if nid in sheet.footnotes]
                paint_name = title_value or row.title
                line = ' | '.join(x for x in [paint_name, row.color_code] if normalize_text(x))
                if available_interiors:
                    line += ' — Available with: ' + ', '.join(available_interiors)
                exterior_group_lines.append(line)

                bullet_groups: List[Tuple[str, Sequence[str]]] = []
                if available_interiors:
                    bullet_groups.append(('Available with Interior Colours', available_interiors))
                if note_texts:
                    bullet_groups.append(('Guide notes', note_texts))
                parts.append(
                    render_article(
                        article_heading(entity, feature_title(f'Exterior paint | {paint_name}', row.color_code)),
                        identity_fields(
                            data,
                            extra_fields=[('Touch-Up Paint Number', row.touch_up_paint_number)],
                        ),
                        bullet_groups,
                    )
                )

            if exterior_group_lines:
                for idx, chunk in enumerate(chunk_feature_items(exterior_group_lines), start=1):
                    title = 'Colour and trim | exterior paint grouped passage'
                    if idx > 1:
                        title += f' | part {idx}'
                    parts.append(
                        render_article(
                            article_heading(entity, title),
                            identity_fields(data),
                            [('Exterior paint lines from guide', chunk)],
                        )
                    )

        general_notes = unique_preserve_order(list(sheet.footnotes.values()) + list(sheet.bullet_notes))
        if general_notes:
            parts.append(
                render_article(
                    article_heading(entity, f'{sheet.name} | colour and trim notes'),
                    identity_fields(data),
                    [('Guide notes', general_notes)],
                )
            )
    
    parts.append('</section>')
    return ''.join(parts)


def render_trim_color_sections(data: 'WorkbookData', trim: 'TrimDef') -> str:
    """Render color sections for trim."""
    if not data.color_sheets:
        return ''
    
    entity = full_trim_heading(data, trim)
    parts = [f'<section class="trim-colours"><h2>{html_module.escape(entity)} | Colour and trim from guide</h2>']
    
    for sheet in data.color_sheets:
        trim_interior_rows = [row for row in sheet.interior_rows if trim_matches_decor(trim, row.decor_level)]
        relevant_interior_columns: List[str] = []

        if trim_interior_rows:
            interior_group_lines: List[str] = []
            for row in trim_interior_rows:
                color_pairs = [
                    (color_name, code)
                    for color_name, code in row.colors.items()
                    if normalize_text(code) and normalize_text(code) != '--'
                ]
                if not color_pairs:
                    continue
                relevant_interior_columns.extend(cn for cn, _ in color_pairs)
                
                if len(color_pairs) == 1:
                    color_name, code = color_pairs[0]
                    line = f'{row.decor_level} | {row.seat_trim} — {color_name}: {code}'
                else:
                    color_text = '; '.join(f'{cn}: {code}' for cn, code in color_pairs)
                    line = f'{row.decor_level} | {row.seat_trim} — {color_text}'
                interior_group_lines.append(line)
                
                color_items = [f'{cn}: {code}' for cn, code in color_pairs]
                parts.append(
                    render_article(
                        article_heading(entity, feature_title(f'Interior trim | {row.decor_level} | {row.seat_trim}', row.seat_code)),
                        identity_fields(
                            data,
                            trim,
                            extra_fields=[
                                ('Decor level', row.decor_level),
                                ('Seat type', row.seat_type),
                                ('Seat trim', row.seat_trim),
                            ],
                        ),
                        [('Interior colours and guide values', color_items)],
                    )
                )

            relevant_interior_columns = unique_preserve_order(relevant_interior_columns)

            if interior_group_lines:
                for idx, chunk in enumerate(chunk_feature_items(interior_group_lines), start=1):
                    title = 'Colour and trim | interior grouped passage'
                    if idx > 1:
                        title += f' | part {idx}'
                    parts.append(
                        render_article(
                            article_heading(entity, title),
                            identity_fields(data, trim),
                            [('Interior colour and trim lines from guide', chunk)],
                        )
                    )

        # Exterior colors for this trim
        grouped_exterior_lines: List[str] = []
        for row in sheet.exterior_rows:
            available_interiors = [
                color_name
                for color_name in relevant_interior_columns
                if normalize_text(row.colors.get(color_name, '')).upper() in ('A', 'S')
            ] if relevant_interior_columns else [
                color_name
                for color_name, status in row.colors.items()
                if normalize_text(status).upper() in ('A', 'S')
            ]
            if not available_interiors and relevant_interior_columns:
                continue
            
            title_value, title_note_ids = parse_value_and_footnote_ids(row.title)
            note_texts = [sheet.footnotes[nid] for nid in title_note_ids if nid in sheet.footnotes]
            paint_name = title_value or row.title
            line = ' | '.join(x for x in [paint_name, row.color_code] if normalize_text(x))
            if available_interiors:
                line += ' — Available with: ' + ', '.join(available_interiors)
            grouped_exterior_lines.append(line)

            bullet_groups: List[Tuple[str, Sequence[str]]] = []
            if available_interiors:
                bullet_groups.append(('Available with Interior Colours', available_interiors))
            if note_texts:
                bullet_groups.append(('Guide notes', note_texts))
            parts.append(
                render_article(
                    article_heading(entity, feature_title(f'Exterior paint | {paint_name}', row.color_code)),
                    identity_fields(
                        data,
                        trim,
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
                        identity_fields(data, trim),
                        [('Exterior paint lines from guide', chunk)],
                    )
                )

        general_notes = unique_preserve_order(list(sheet.footnotes.values()) + list(sheet.bullet_notes))
        if general_notes:
            parts.append(
                render_article(
                    article_heading(entity, f'{sheet.name} | colour and trim notes'),
                    identity_fields(data, trim),
                    [('Guide notes', general_notes)],
                )
            )
    
    parts.append('</section>')
    return ''.join(parts)


def render_spec_records(data: 'WorkbookData', columns: List['SpecColumn'], *, trim: 'TrimDef | None' = None) -> str:
    """Render specification records."""
    if not columns:
        return ''
    
    entity = full_trim_heading(data, trim) if trim is not None else data.vehicle_name
    parts = [f'<section class="spec-sections"><h2>{html_module.escape(entity)} | Specifications and dimensions from guide</h2>']
    
    for column in columns:
        grouped: OrderedDict = OrderedDict()
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
                            extra_fields=[('Column context', header_text)],
                        ),
                        [('Guide values', value_chunk)],
                    )
                )
    
    parts.append('</section>')
    return ''.join(parts)


def render_trim_spec_sections(data: 'WorkbookData', trim: 'TrimDef') -> str:
    """Render specifications for trim."""
    direct_columns = [column for column in data.spec_columns if column_matches_trim(column, trim)]
    if direct_columns:
        return render_spec_records(data, direct_columns, trim=trim)
    return render_spec_records(data, data.spec_columns, trim=trim)


def render_engine_axles_section(data: 'WorkbookData') -> str:
    """Render engine and axle section."""
    if not data.engine_axle_entries:
        return ''
    
    entity = data.vehicle_name
    parts = [f'<section class="engine-axles"><h2>{html_module.escape(entity)} | Engine, axle and GVWR from guide</h2>']
    
    for entry in data.engine_axle_entries:
        grouped: OrderedDict = OrderedDict()
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
                            extra_fields=[('Top label', entry.top_label)],
                        ),
                        [('Guide values', chunk)],
                    )
                )
    
    parts.append('</section>')
    return ''.join(parts)


def render_trailering_section(data: 'WorkbookData') -> str:
    """Render trailering section."""
    if not data.trailering_records and not data.gcwr_records:
        return ''
    
    entity = data.vehicle_name
    parts = [f'<section class="trailering"><h2>{html_module.escape(entity)} | Trailering and GCWR from guide</h2>']
    
    from ..utils.text_utils import sentence_chunks
    
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
                    extra_fields=[('Table title', record.table_title), ('Axle ratio', record.axle_ratio)],
                ),
                bullet_groups,
            )
        )
    
    parts.append('</section>')
    return ''.join(parts)
