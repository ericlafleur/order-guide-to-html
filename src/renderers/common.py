"""Common rendering utilities."""

import html as html_module
import re
from typing import TYPE_CHECKING, Dict, List, Optional, Sequence, Tuple

from ..aggregators.feature_aggregator import (
    category_for_model_feature,
    category_for_trim_feature,
)
from ..inference import infer_trim_drive_type
from ..models.aggregate_models import ModelFeatureAggregate, TrimFeatureAggregate
from ..parsers.common import parse_status_value, parse_value_and_footnote_ids
from ..utils.constants import CATEGORY_ORDER, TRIM_SECTION_ORDER
from ..utils.text_utils import (
    chunk_list,
    compact_text,
    htmlize_text,
    normalize_text,
    sentence_chunks,
    unique_preserve_order,
)

if TYPE_CHECKING:
    from ..models.base_models import MatrixSheet, TrimDef, WorkbookData


def page_entity(data: 'WorkbookData', trim: Optional['TrimDef'] = None) -> str:
    """Get entity name for this page.
    
    Args:
        data: Workbook data
        trim: Optional trim definition
        
    Returns:
        Entity name string
    """
    if trim is None:
        return data.vehicle_name
    return normalize_text(f'{data.vehicle_name} {trim.name}')


def full_trim_heading(data: 'WorkbookData', trim: 'TrimDef') -> str:
    """Get full heading for trim page.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        Full heading string
    """
    base = page_entity(data, trim)
    if trim.code and trim.code.lower() != trim.name.lower():
        return f'{base} ({trim.code})'
    return base


def article_heading(entity: str, base_title: str) -> str:
    """Generate article heading.
    
    Args:
        entity: Entity name
        base_title: Base title
        
    Returns:
        Combined heading
    """
    base_title = normalize_text(base_title)
    if not base_title:
        return entity
    return f'{entity} | {base_title}'


def source_tabs_from_contexts(contexts: Sequence[str]) -> str:
    """Extract source tabs from contexts.
    
    Args:
        contexts: Context strings
        
    Returns:
        Semicolon-separated source tabs
    """
    tabs: List[str] = []
    for context in contexts:
        context = normalize_text(context)
        if not context:
            continue
        first = normalize_text(context.split('|', 1)[0])
        if first:
            tabs.append(first)
    return '; '.join(unique_preserve_order(tabs))


def dedupe_fields(fields: Sequence[Tuple[str, str]]) -> List[Tuple[str, str]]:
    """Deduplicate field list.
    
    Args:
        fields: List of (label, value) tuples
        
    Returns:
        Deduplicated list
    """
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
    data: 'WorkbookData',
    trim: Optional['TrimDef'] = None,
    *,
    extra_fields: Sequence[Tuple[str, str]] = (),
    drive_type: Optional[str] = None,
) -> List[Tuple[str, str]]:
    """Generate identity fields for RAG chunks.
    
    Args:
        data: Workbook data
        trim: Optional trim definition
        extra_fields: Extra fields to include
        drive_type: Override drive type
        
    Returns:
        List of (label, value) tuples
    """
    fields: List[Tuple[str, str]] = []
    
    if data.propulsion:
        fields.append(('Propulsion', data.propulsion))
    if data.vehicle_type:
        fields.append(('Vehicle type', data.vehicle_type.upper()))
    
    if trim is not None:
        resolved_drive = drive_type or infer_trim_drive_type(data.spec_columns, trim, data.drive_types)
        if resolved_drive:
            fields.append(('Drive type', resolved_drive))
        elif data.drive_types:
            fields.append(('Drive types', ', '.join(data.drive_types)))
    else:
        if data.drive_types:
            fields.append(('Drive types', ', '.join(data.drive_types)))
    
    fields.extend(extra_fields)
    return dedupe_fields(fields)


def render_article(
    title: str, 
    fields: Sequence[Tuple[str, str]], 
    bullet_groups: Sequence[Tuple[str, Sequence[str]]] = ()
) -> str:
    """Render an article block.
    
    Args:
        title: Article title
        fields: List of (label, value) field tuples
        bullet_groups: List of (label, items) bullet list tuples
        
    Returns:
        HTML string
    """
    parts = [f'<article class="guide-record"><h3>{htmlize_text(title)}</h3>']
    
    for label, value in fields:
        value = normalize_text(value)
        if value:
            parts.append(f'<p><strong>{html_module.escape(label)}:</strong> {htmlize_text(value)}</p>')
    
    for label, items in bullet_groups:
        clean_items = [normalize_text(item) for item in items if normalize_text(item)]
        if not clean_items:
            continue
        parts.append(f'<div class="record-list"><p><strong>{html_module.escape(label)}:</strong></p><ul>')
        for item in clean_items:
            parts.append(f'<li>{htmlize_text(item)}</li>')
        parts.append('</ul></div>')
    
    parts.append('</article>')
    return ''.join(parts)


def sort_category_key(category: str) -> Tuple[int, str]:
    """Sort key for categories.
    
    Args:
        category: Category name
        
    Returns:
        Sort key tuple
    """
    return (CATEGORY_ORDER.get(category, 999), normalize_text(category).lower())


def model_status_summary_lines(
    signature: Tuple[Tuple[str, str, Tuple[str, ...]], ...]
) -> List[str]:
    """Generate status summary lines for model.
    
    Args:
        signature: Status signature tuple
        
    Returns:
        List of summary lines
    """
    lines = []
    for raw, label, names in signature:
        lines.append(f'{label} [{raw}]: {", ".join(names)}')
    return lines


def availability_lines_for_trim(agg: TrimFeatureAggregate) -> List[str]:
    """Generate availability lines for trim feature.
    
    Args:
        agg: Trim feature aggregate
        
    Returns:
        List of availability lines
    """
    lines: List[str] = []
    for (raw, label), contexts in agg.availability_contexts.items():
        context_text = '; '.join(contexts)
        if context_text:
            lines.append(f'{label} [{raw}] — {context_text}')
        else:
            lines.append(f'{label} [{raw}]')
    return lines


def availability_summary_for_trim(agg: TrimFeatureAggregate) -> str:
    """Generate availability summary for trim feature.
    
    Args:
        agg: Trim feature aggregate
        
    Returns:
        Availability summary string
    """
    parts: List[str] = []
    for (raw, label), contexts in agg.availability_contexts.items():
        tabs = source_tabs_from_contexts(contexts)
        bit = f'{label} [{raw}]'
        if tabs:
            bit += f' ({tabs})'
        parts.append(bit)
    return '; '.join(parts)


def availability_lines_for_model(agg: ModelFeatureAggregate) -> List[str]:
    """Generate availability lines for model feature.
    
    Args:
        agg: Model feature aggregate
        
    Returns:
        List of availability lines
    """
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
    """Generate availability summary for model feature.
    
    Args:
        agg: Model feature aggregate
        
    Returns:
        Availability summary string
    """
    parts: List[str] = []
    for signature, _contexts in agg.availability_contexts.items():
        parts.append(' ; '.join(model_status_summary_lines(signature)))
    return ' / '.join(parts)


def chunk_feature_items(items: Sequence[str]) -> List[List[str]]:
    """Chunk feature items.
    
    Args:
        items: List of items
        
    Returns:
        List of chunked items
    """
    return chunk_list(items, max_words=135, max_items=8)


def trim_matches_decor(trim: 'TrimDef', decor_value: str) -> bool:
    """Check if trim matches decor level.
    
    Args:
        trim: Trim definition
        decor_value: Decor level value
        
    Returns:
        True if matches
    """
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


def column_matches_trim(column, trim: 'TrimDef') -> bool:
    """Check if spec column matches trim.
    
    Args:
        column: Spec column
        trim: Trim definition
        
    Returns:
        True if matches
    """
    blob = ' '.join([column.top_label, column.header] + column.header_lines).lower()
    if trim.name.lower() in blob or trim.code.lower() in blob:
        return True
    return False
