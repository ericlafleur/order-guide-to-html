from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Mapping, MutableMapping, Sequence, Tuple

from .utils import htmlize_text, normalize_text, unique_preserve_order
import html


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
WITH_CODES_BRACKETS_RE = re.compile(r'\[w/\s*([^\]]+)\]', re.I)

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
LOW_SIGNAL_COMPARISON_DOMAINS = {
    'Colour and trim',
}


@dataclass
class GuideTextCleaner:
    glossary: Dict[str, str] = field(default_factory=dict)

    def load_glossary(self, glossary: Mapping[str, str]) -> None:
        self.glossary = {
            normalize_text(k): normalize_text(v)
            for k, v in glossary.items()
            if normalize_text(k) and normalize_text(v)
        }

    def is_code_only(self, text: object) -> bool:
        text = normalize_text(text)
        if not text or not CODE_ONLY_RE.fullmatch(text):
            return False
        return any(ch.isdigit() for ch in text) or len(text) > 4

    def expand_code_sequence(self, seq: str) -> str:
        parts = [normalize_text(x) for x in re.split(r'\s+and\s+|\s*,\s*|\s*;\s*', normalize_text(seq)) if normalize_text(x)]
        descriptions: List[str] = []
        leftovers: List[str] = []
        for part in parts:
            if normalize_text(part) in self.glossary:
                descriptions.append(self.clean_customer_text(self.glossary[normalize_text(part)]))
            elif self.is_code_only(part):
                leftovers.append('')
            else:
                leftovers.append(self.clean_customer_text(part))
        descriptions = [d for d in descriptions if d]
        leftovers = [x for x in leftovers if x]
        joined = unique_preserve_order(descriptions + leftovers)
        return ' with ' + ' and '.join(joined) if joined else ''

    def _expand_with_codes(self, match: re.Match[str]) -> str:
        return self.expand_code_sequence(match.group(1))

    def _strip_code_parenthetical(self, match: re.Match[str]) -> str:
        content = normalize_text(match.group(1))
        if re.fullmatch(r'[A-Z0-9-]{2,10}', content) and re.search(r'[A-Z]', content):
            return ''
        return f'({content})'

    def clean_customer_text(self, text: object) -> str:
        text = normalize_text(text)
        if not text:
            return ''
        text = WITH_CODES_BRACKETS_RE.sub(self._expand_with_codes, text)
        text = CODE_PARENS_RE.sub(self._strip_code_parenthetical, text)
        text = FILE_CODE_RE.sub('', text)
        text = STATUS_BRACKET_RE.sub('', text)
        text = re.sub(
            r'\b(?:Trim code|Model code|Feature code|Reference code|Orderable code)\s*:?\s*[A-Z0-9-]{2,10}\b',
            '',
            text,
            flags=re.I,
        )
        if SOURCE_PARENS_HINT_RE.search(text):
            text = re.sub(
                r'\s+\(([^()]*)\)$',
                lambda m: '' if (';' in m.group(1) or SOURCE_PARENS_HINT_RE.search(m.group(1))) else m.group(0),
                text,
            )
        text = text.replace('includes ,', 'includes ')
        text = re.sub(r'\s+,', ',', text)
        text = re.sub(r',\s*,', ', ', text)
        text = re.sub(r'\s{2,}', ' ', text)
        text = re.sub(r'\|\s*\|', '|', text)
        text = re.sub(r'\s*\|\s*$', '', text)
        text = re.sub(r'^\s*\|\s*', '', text)
        text = re.sub(r'\s+([.;:])', r'\1', text)
        return normalize_text(text).strip(' |,;')

    def clean_heading_text(self, text: object) -> str:
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
            seg = self.clean_customer_text(seg)
            if not seg:
                continue
            if SHEET_SEGMENT_RE.fullmatch(seg):
                continue
            if self.is_code_only(seg) or FILE_CODE_RE.fullmatch(seg):
                continue
            key = seg.lower()
            if key in seen:
                continue
            seen.add(key)
            segments.append(seg)
        return ' | '.join(segments)

    def article_heading(self, entity: str, base_title: str) -> str:
        base_title = normalize_text(base_title)
        if not base_title:
            return self.clean_heading_text(entity)
        return self.clean_heading_text(f'{entity} | {base_title}')

    def dedupe_fields(self, fields: Sequence[Tuple[str, str]]) -> List[Tuple[str, str]]:
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
        self,
        data,
        trim=None,
        *,
        category: str = '',
        extra_fields: Sequence[Tuple[str, str]] = (),
    ) -> List[Tuple[str, str]]:
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
            value = self.clean_customer_text(value)
            if not value or self.is_code_only(value):
                continue
            fields.append((label, value))
        return self.dedupe_fields(fields)

    def cleaned_render_article(
        self,
        title: str,
        fields: Sequence[Tuple[str, str]],
        bullet_groups: Sequence[Tuple[str, Sequence[str]]] = (),
    ) -> str:
        title = self.clean_heading_text(title)
        parts = [f'<article class="guide-record"><h3>{htmlize_text(title)}</h3>']
        clean_fields: List[Tuple[str, str]] = []
        seen_fields = set()
        for label, value in fields:
            label = normalize_text(label)
            if label in FIELD_DROP_LABELS:
                continue
            label = FIELD_RENAME.get(label, label)
            value = self.clean_customer_text(value)
            if not label or not value or self.is_code_only(value):
                continue
            item = (label, value)
            if item in seen_fields:
                continue
            seen_fields.add(item)
            clean_fields.append(item)
        for label, value in clean_fields:
            parts.append(f'<p><strong>{html.escape(label)}:</strong> {htmlize_text(value)}</p>')
        for label, items in bullet_groups:
            label = normalize_text(label)
            if label in FIELD_DROP_LABELS:
                continue
            label = FIELD_RENAME.get(label, label)
            clean_items: List[str] = []
            seen_items = set()
            for item in items:
                item = self.clean_customer_text(item)
                if not item or self.is_code_only(item):
                    continue
                if item in seen_items:
                    continue
                seen_items.add(item)
                clean_items.append(item)
            if not clean_items:
                continue
            parts.append(f'<div class="record-list"><p><strong>{html.escape(label)}:</strong></p><ul>')
            for item in clean_items:
                parts.append(f'<li>{htmlize_text(item)}</li>')
            parts.append('</ul></div>')
        parts.append('</article>')
        return ''.join(parts)

    def clean_trim_heading(self, data, trim) -> str:
        return normalize_text(f'{data.vehicle_name} {trim.name}')

    def clean_feature_title(self, label: str, orderable_code: str = '', reference_code: str = '') -> str:
        return self.clean_customer_text(label)

    def clean_model_status_summary_lines(self, signature) -> List[str]:
        lines = []
        for _raw, label, names in signature:
            if names:
                clean_names = [self.clean_customer_text(name) for name in names if self.clean_customer_text(name)]
                if clean_names:
                    lines.append(f'{self.clean_customer_text(label)}: {", ".join(clean_names)}')
        return [line for line in lines if normalize_text(line)]

    def clean_availability_summary_for_trim(self, agg) -> str:
        labels = [self.clean_customer_text(label) for (_raw, label), _contexts in agg.availability_contexts.items()]
        labels = unique_preserve_order(labels)
        return '; '.join(labels)

    def clean_availability_lines_for_trim(self, agg) -> List[str]:
        labels = [self.clean_customer_text(label) for (_raw, label), _contexts in agg.availability_contexts.items()]
        return unique_preserve_order(labels)

    def clean_availability_summary_for_model(self, agg) -> str:
        parts: List[str] = []
        for signature, _contexts in agg.availability_contexts.items():
            parts.extend(self.clean_model_status_summary_lines(signature))
        return ' / '.join(unique_preserve_order(parts))

    def clean_availability_lines_for_model(self, agg) -> List[str]:
        lines: List[str] = []
        for signature, _contexts in agg.availability_contexts.items():
            lines.extend(self.clean_model_status_summary_lines(signature))
        return unique_preserve_order(lines)

    def clean_trim_group_line(self, agg) -> str:
        feature_text = self.clean_customer_text(agg.description or agg.title)
        availability = self.clean_availability_summary_for_trim(agg)
        return f'{feature_text} — {availability}' if availability else feature_text

    def clean_model_group_line(self, agg) -> str:
        feature_text = self.clean_customer_text(agg.description or agg.title)
        availability = self.clean_availability_summary_for_model(agg)
        return f'{feature_text} — {availability}' if availability else feature_text

    def trim_feature_is_unavailable(self, agg) -> bool:
        statuses = []
        for (raw, label), _contexts in agg.availability_contexts.items():
            raw_value = normalize_text(raw)
            label_value = normalize_text(self.clean_customer_text(label)).lower()
            statuses.append((raw_value, label_value))
        if not statuses:
            return False
        return all(label == 'not available' or raw == '--' for raw, label in statuses)

    def clean_title_document(self, title: str, *body_parts: str) -> str:
        title = self.clean_heading_text(title)
        parts = ['<html><head><meta charset="utf-8"></head><body>', f'<h1>{html.escape(normalize_text(title))}</h1>']
        parts.extend(part for part in body_parts if part)
        parts.append('</body></html>')
        return ''.join(parts)

    def matrix_row_label(self, row) -> str:
        text = row.description_main or row.description_raw
        text = normalize_text(text).split('\n')[0]
        text = re.sub(r'^NEW!\s+', '', text, flags=re.I)
        if ', includes ' in text.lower():
            text = text.split(',', 1)[0]
        elif '. ' in text and len(text.split('. ', 1)[0]) >= 12:
            text = text.split('. ', 1)[0]
        return self.clean_customer_text(text)

    def clean_manifest_value(self, value):
        if isinstance(value, str):
            return self.clean_heading_text(value) if ('|' in value or value.lower().endswith('overview') or 'comparison' in value.lower()) else self.clean_customer_text(value)
        if isinstance(value, list):
            cleaned = []
            for item in value:
                item = self.clean_manifest_value(item)
                if item in ('', [], None):
                    continue
                cleaned.append(item)
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
            return self.clean_manifest_entry(value)
        return value

    def clean_manifest_entry(self, entry: Dict[str, object]) -> Dict[str, object]:
        cleaned: Dict[str, object] = {}
        for key, value in entry.items():
            if key in MANIFEST_DROP_KEYS:
                continue
            key = MANIFEST_RENAME_KEYS.get(key, key)
            value = self.clean_manifest_value(value)
            if value in ('', [], None):
                continue
            cleaned[key] = value
        if cleaned.get('type') == 'model' and 'trim_names_from_guide' in cleaned:
            cleaned['available_trims'] = cleaned.pop('trim_names_from_guide')
        return cleaned

    def clean_manifest(self, manifest: Dict[str, object]) -> Dict[str, object]:
        cleaned = {
            'vehicle_name': self.clean_customer_text(manifest.get('vehicle_name', '')),
            'vehicle_key': manifest.get('vehicle_key', ''),
            'files': [self.clean_manifest_entry(entry) for entry in manifest.get('files', [])],
        }
        return {k: v for k, v in cleaned.items() if v not in ('', [], None)}
