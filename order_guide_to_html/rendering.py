from __future__ import annotations

from collections import OrderedDict
from typing import Dict, List, Optional, Sequence, Tuple

from .utils import chunk_feature_items, chunk_list
from .models import (
    GCWRRecord, ModelFeatureAggregate, PowertrainTraileringGroup, SpecGroupDoc,
)
from .parsing import parse_status_value, parse_value_and_footnote_ids
from .classification import (
    DOMAIN_COLOR, category_for_model_feature, category_for_trim_feature,
    sort_category_key,
)
from .configuration import trim_colour_context
import html
from .aggregation import FeatureAggregationService
from .cleaning import GuideTextCleaner, normalize_text, unique_preserve_order


class HtmlRenderer:
    """Render clean HTML pages from parsed order-guide data."""

    def __init__(self, cleaner: GuideTextCleaner, aggregation: FeatureAggregationService):
        self.cleaner = cleaner
        self.aggregation = aggregation

    def render_page_identity_section(self, data, trim=None) -> str:
        entity = self.cleaner.clean_trim_heading(data, trim) if trim is not None else data.vehicle_name
        if trim is None:
            fields = [('Vehicle', data.vehicle_name)]
            trim_names = [normalize_text(t.name) for t in data.trim_defs if normalize_text(t.name)]
            bullets = [(self.cleaner.t('Available trims'), unique_preserve_order(trim_names))] if trim_names else []
            title = f"{entity} | {self.cleaner.t('Model overview')}"
        else:
            fields = [('Vehicle', data.vehicle_name), ('Trim', trim.name)]
            bullets = []
            title = f"{entity} | {self.cleaner.t('Trim overview')}"
        return '<section class="vehicle-identity">' + self.cleaner.cleaned_render_article(title, fields, bullets) + '</section>'

    def render_trim_lineup_section(self, data) -> str:
        trim_names = [normalize_text(trim.name) for trim in data.trim_defs if normalize_text(trim.name)]
        if not trim_names:
            return ''
        article = self.cleaner.cleaned_render_article(
            self.cleaner.article_heading(data.vehicle_name, self.cleaner.t('Trim lineup')),
            self.cleaner.filtered_identity_fields(data, category='Other guide content'),
            [(self.cleaner.t('Available trims'), unique_preserve_order(trim_names))],
        )
        return '<section class="trim-lineup">' + article + '</section>'

    def render_feature_section(
        self,
        data,
        entity: str,
        category_groups: Dict[str, List[object]],
        *,
        trim=None,
        model_mode: bool,
        section_class: str,
        section_heading: str,
        bullet_label: str,
        title_suffix: str,
    ) -> str:
        if not category_groups:
            return ''
        parts = [f'<section class="{section_class}"><h2>{html.escape(entity)} | {html.escape(section_heading)}</h2>']
        for category in sorted(category_groups.keys(), key=sort_category_key):
            items = category_groups[category]
            display_category = self.cleaner.t(category) if self.cleaner.language == 'fr' else category
            lines = [self.cleaner.clean_model_group_line(item) if model_mode else self.cleaner.clean_trim_group_line(item) for item in items]
            for idx, line_chunk in enumerate(chunk_feature_items(lines), start=1):
                title = f'{display_category} | {title_suffix}'
                if idx > 1:
                    title += f' | {self.cleaner.t("part")} {idx}'
                parts.append(
                    self.cleaner.cleaned_render_article(
                        self.cleaner.article_heading(entity, title),
                        self.cleaner.filtered_identity_fields(data, trim, category=category),
                        [(bullet_label, line_chunk)],
                    )
                )
        parts.append('</section>')
        return ''.join(parts)

    def render_grouped_feature_sections(self, data, features, *, trim=None, model_mode=False) -> str:
        if not features:
            return ''
        entity = self.cleaner.clean_trim_heading(data, trim) if trim is not None else data.vehicle_name
        if trim is not None and not model_mode:
            positive_groups: Dict[str, List[object]] = OrderedDict()
            unavailable_groups: Dict[str, List[object]] = OrderedDict()
            for feature in features:
                category = category_for_trim_feature(feature)
                target = unavailable_groups if self.cleaner.trim_feature_is_unavailable(feature) else positive_groups
                target.setdefault(category, []).append(feature)
            return ''.join([
                self.render_feature_section(
                    data,
                    entity,
                    positive_groups,
                    trim=trim,
                    model_mode=False,
                    section_class='grouped-feature-passages',
                    section_heading=self.cleaner.t('Feature highlights'),
                    bullet_label=self.cleaner.t('Feature highlights'),
                    title_suffix=self.cleaner.t('Highlights'),
                ),
                self.render_feature_section(
                    data,
                    entity,
                    unavailable_groups,
                    trim=trim,
                    model_mode=False,
                    section_class='unavailable-feature-passages',
                    section_heading=self.cleaner.t('Features explicitly not offered on this trim'),
                    bullet_label=self.cleaner.t('Unavailable features'),
                    title_suffix=self.cleaner.t('Unavailable features'),
                ),
            ])

        category_groups: Dict[str, List[object]] = OrderedDict()
        for feature in features:
            category = category_for_model_feature(feature) if model_mode else category_for_trim_feature(feature)
            category_groups.setdefault(category, []).append(feature)
        return self.render_feature_section(
            data,
            entity,
            category_groups,
            trim=trim,
            model_mode=model_mode,
            section_class='grouped-feature-passages',
            section_heading=self.cleaner.t('Feature highlights'),
            bullet_label=self.cleaner.t('Feature highlights'),
            title_suffix=self.cleaner.t('Highlights'),
        )

    def clean_model_colour_group_lines(self, data) -> List[str]:
        lines: List[str] = []
        for sheet in data.color_sheets:
            for row in sheet.interior_rows:
                colors = [self.cleaner.clean_customer_text(color) for color, code in row.colors.items() if normalize_text(code) and normalize_text(code) != '--']
                parts = [self.cleaner.t('Colour and trim'), row.decor_level, row.seat_type, row.seat_trim]
                summary = ' | '.join(self.cleaner.clean_customer_text(x) for x in parts if self.cleaner.clean_customer_text(x))
                if colors:
                    summary += (' — Couleurs offertes: ' if self.cleaner.language == 'fr' else ' — Available colours: ') + '; '.join(unique_preserve_order(colors))
                lines.append(summary)
            for row in sheet.exterior_rows:
                title_value, _title_note_ids = parse_value_and_footnote_ids(row.title)
                availability_lines = []
                for color, status in row.colors.items():
                    status_code, status_label, _ = parse_status_value(status, {}, sheet.footnotes)
                    color = self.cleaner.clean_customer_text(color)
                    if not color or not status_label:
                        continue
                    if status_code == '--':
                        availability_lines.append(f"{color}: {'Non livrable' if self.cleaner.language == 'fr' else 'Not available'}")
                    else:
                        availability_lines.append(f"{color}: {self.cleaner.clean_customer_text(self.cleaner.translate_status_label(status_label))}")
                summary = ' | '.join(self.cleaner.clean_customer_text(x) for x in [self.cleaner.t('Colour and trim'), title_value or row.title] if self.cleaner.clean_customer_text(x))
                if availability_lines:
                    summary += ' — ' + '; '.join(unique_preserve_order(availability_lines))
                lines.append(summary)
        return unique_preserve_order(lines)

    def clean_trim_colour_group_lines(self, data, trim) -> List[str]:
        ctx = trim_colour_context(data, trim)
        lines: List[str] = []
        for _sheet, row, color_lines in ctx['interior_items']:
            colors = [self.cleaner.clean_customer_text(line.split(':', 1)[0]) for line in color_lines if self.cleaner.clean_customer_text(line.split(':', 1)[0])]
            summary = ' | '.join(self.cleaner.clean_customer_text(x) for x in [self.cleaner.t('Colour and trim'), row.decor_level, row.seat_type, row.seat_trim] if self.cleaner.clean_customer_text(x))
            if colors:
                summary += (' — Couleurs offertes: ' if self.cleaner.language == 'fr' else ' — Available colours: ') + '; '.join(unique_preserve_order(colors))
            lines.append(summary)
        for sheet, row, availability_lines, _note_texts in ctx['exterior_items']:
            title_value, _title_note_ids = parse_value_and_footnote_ids(row.title)
            clean_lines = []
            for line in availability_lines:
                color, _, raw_status = line.partition(':')
                status_code, status_label, _ = parse_status_value(raw_status, {}, sheet.footnotes)
                color = self.cleaner.clean_customer_text(color)
                if not color or not status_label:
                    continue
                if status_code == '--':
                    clean_lines.append(f"{color}: {'Non livrable' if self.cleaner.language == 'fr' else 'Not available'}")
                else:
                    clean_lines.append(f"{color}: {self.cleaner.clean_customer_text(self.cleaner.translate_status_label(status_label))}")
            summary = ' | '.join(self.cleaner.clean_customer_text(x) for x in [self.cleaner.t('Colour and trim'), title_value or row.title] if self.cleaner.clean_customer_text(x))
            if clean_lines:
                summary += ' — ' + '; '.join(unique_preserve_order(clean_lines))
            lines.append(summary)
        return unique_preserve_order(lines)

    def render_grouped_colour_summary(self, data, trim=None) -> str:
        if not data.color_sheets:
            return ''
        entity = self.cleaner.clean_trim_heading(data, trim) if trim is not None else data.vehicle_name
        lines = self.clean_trim_colour_group_lines(data, trim) if trim is not None else self.clean_model_colour_group_lines(data)
        if not lines:
            return ''
        colour_and_trim_label = self.cleaner.t('Colour and trim')
        parts = [f'<section class="grouped-colour-passages"><h2>{html.escape(entity)} | {html.escape(colour_and_trim_label)}</h2>']
        for idx, chunk in enumerate(chunk_feature_items(lines), start=1):
            chunk_title = self.cleaner.article_heading(entity, self.cleaner.t('Colour and trim')) if idx == 1 else self.cleaner.article_heading(entity, f"{self.cleaner.t('Colour and trim')} | {self.cleaner.t('part')} {idx}")
            parts.append(
                self.cleaner.cleaned_render_article(
                    chunk_title,
                    self.cleaner.filtered_identity_fields(data, trim, category=DOMAIN_COLOR),
                    [(self.cleaner.t('Colour and trim combinations'), chunk)],
                )
            )
        parts.append('</section>')
        return ''.join(parts)

    def render_model_overview_page(self, data) -> str:
        entity = data.vehicle_name
        return self.cleaner.clean_title_document(
            f"{entity} | {self.cleaner.t('Model overview')}",
            self.render_page_identity_section(data),
            self.render_trim_lineup_section(data),
            self.render_grouped_feature_sections(data, self.aggregation.aggregate_model_features(data), model_mode=True),
            self.render_grouped_colour_summary(data),
        )

    def render_trim_overview_page(self, data, trim) -> str:
        entity = self.cleaner.clean_trim_heading(data, trim)
        return self.cleaner.clean_title_document(
            f"{entity} | {self.cleaner.t('Trim overview')}",
            self.render_page_identity_section(data, trim=trim),
            self.render_grouped_feature_sections(data, self.aggregation.aggregate_trim_features(data, trim), trim=trim, model_mode=False),
            self.render_grouped_colour_summary(data, trim=trim),
        )

    def render_comparison_domain_page(self, data, category: str, features: Sequence[ModelFeatureAggregate]) -> str:
        entity = data.vehicle_name
        lines = [self.cleaner.clean_model_group_line(feature) for feature in features]
        category_label = self.cleaner.t(category) if self.cleaner.language == 'fr' else category
        trim_comparison_label = self.cleaner.t('Trim comparison')
        parts = [f'<section class="comparison-domain"><h2>{html.escape(entity)} | {html.escape(category_label)} {html.escape(trim_comparison_label)}</h2>']
        for idx, chunk in enumerate(chunk_feature_items(lines), start=1):
            title = self.cleaner.article_heading(entity, f"{self.cleaner.t(category) if self.cleaner.language == 'fr' else category} | {self.cleaner.t('Trim comparison highlights')}")
            if idx > 1:
                title = self.cleaner.article_heading(entity, f"{self.cleaner.t(category) if self.cleaner.language == 'fr' else category} | {self.cleaner.t('Trim comparison highlights')} | {self.cleaner.t('part')} {idx}")
            parts.append(
                self.cleaner.cleaned_render_article(
                    title,
                    self.cleaner.filtered_identity_fields(
                        data,
                        category=category,
                        extra_fields=[(self.cleaner.t('Comparison axis'), self.cleaner.t('Trim'))],
                    ),
                    [(self.cleaner.t('Comparison lines from guide'), chunk)],
                )
            )
        parts.append('</section>')
        return self.cleaner.clean_title_document(f"{entity} | {self.cleaner.t(category) if self.cleaner.language == 'fr' else category} {self.cleaner.t('Trim comparison')}", ''.join(parts))

    def render_spec_group_page(self, data, group: SpecGroupDoc, trim=None) -> str:
        entity = self.cleaner.clean_trim_heading(data, trim) if trim is not None else data.vehicle_name
        parts = [
            '<section class="configuration-identity">'
            + self.cleaner.cleaned_render_article(
                self.cleaner.article_heading(entity, self.cleaner.t('Configuration identity')),
                self.cleaner.filtered_identity_fields(data, trim, category=self.cleaner.t('Dimensions and specifications') if self.cleaner.language == 'fr' else 'Dimensions and specifications'),
            )
            + '</section>'
        ]
        for column in group.columns:
            grouped_values: 'OrderedDict[str, List[str]]' = OrderedDict()
            for cell in column.cells:
                grouped_values.setdefault(cell.section or 'Data', []).append(f'{cell.label}: {cell.value}')
            for section_name, values in grouped_values.items():
                for idx, chunk in enumerate(chunk_list(values, max_words=125, max_items=10), start=1):
                    title_bits = []
                    if normalize_text(section_name) and normalize_text(section_name).lower() != 'data':
                        title_bits.append(section_name)
                    title_bits.append(self.cleaner.t('Specifications'))
                    title = self.cleaner.article_heading(entity, ' | '.join(title_bits))
                    if idx > 1:
                        title = self.cleaner.article_heading(entity, ' | '.join(title_bits + [f'part {idx}']))
                    parts.append(
                        '<section class="configuration-values">'
                        + self.cleaner.cleaned_render_article(
                            title,
                            self.cleaner.filtered_identity_fields(data, trim, category=self.cleaner.t('Dimensions and specifications') if self.cleaner.language == 'fr' else 'Dimensions and specifications'),
                            [(self.cleaner.t('Specifications'), chunk)],
                        )
                        + '</section>'
                    )
        title_context = group.header or group.top_label
        return self.cleaner.clean_title_document(
            self.cleaner.article_heading(entity, f"{self.cleaner.t('Dimensions and specifications')} | {title_context}"),
            ''.join(parts),
        )

    def render_powertrain_trailering_group_page(self, data, group: PowertrainTraileringGroup, trim=None) -> str:
        entity = self.cleaner.clean_trim_heading(data, trim) if trim is not None else data.vehicle_name
        engines = unique_preserve_order([entry.engine for entry in group.engine_entries] + [record.engine for record in group.trailering_records])
        parts = [
            '<section class="configuration-identity">'
            + self.cleaner.cleaned_render_article(
                self.cleaner.article_heading(entity, self.cleaner.t('Configuration identity')),
                self.cleaner.filtered_identity_fields(
                    data,
                    trim,
                    category=self.cleaner.t('Powertrain and trailering') if self.cleaner.language == 'fr' else 'Powertrain and trailering',
                    extra_fields=[(self.cleaner.t('Engines'), ' ; '.join(engines))] if engines else [],
                ),
            )
            + '</section>'
        ]
        engine_lines: List[str] = []
        for entry in group.engine_entries:
            for item in entry.items:
                line = f"{('Moteur' if self.cleaner.language == 'fr' else 'Engine')} {item.engine or entry.engine if hasattr(item, 'engine') else entry.engine} | {item.category}: {self.cleaner.translate_status_label(item.status_label) or item.raw_status}"
                if item.notes:
                    line += f" | {self.cleaner.t('Notes')}: {' ; '.join(unique_preserve_order(item.notes))}"
                engine_lines.append(line)
        engine_lines = unique_preserve_order(engine_lines)
        if engine_lines:
            for idx, chunk in enumerate(chunk_list(engine_lines, max_words=125, max_items=10), start=1):
                title = self.cleaner.article_heading(entity, self.cleaner.t('Powertrain, axle and GVWR'))
                if idx > 1:
                    title = self.cleaner.article_heading(entity, f"{self.cleaner.t('Powertrain, axle and GVWR')} | {self.cleaner.t('part')} {idx}")
                parts.append(
                    '<section class="powertrain-values">'
                    + self.cleaner.cleaned_render_article(
                        title,
                        self.cleaner.filtered_identity_fields(data, trim, category=self.cleaner.t('Powertrain and trailering') if self.cleaner.language == 'fr' else 'Powertrain and trailering'),
                        [(self.cleaner.t('Specifications'), chunk)],
                    )
                    + '</section>'
                )
        trailering_lines: List[str] = []
        rating_types: List[str] = []
        for record in group.trailering_records:
            line = f"{record.rating_type} | {('Moteur' if self.cleaner.language == 'fr' else 'Engine')} {record.engine} | {'Rapport de pont' if self.cleaner.language == 'fr' else 'Axle ratio'} {record.axle_ratio} | {'Poids maximal de la remorque' if self.cleaner.language == 'fr' else 'Maximum trailer weight'} {record.max_trailer_weight}"
            if normalize_text(record.note_text):
                line += f" | {'Note' if self.cleaner.language == 'fr' else 'Note heading'}: {record.note_text}"
            if record.footnotes:
                line += f" | {self.cleaner.t('Notes')}: {' ; '.join(unique_preserve_order(record.footnotes))}"
            trailering_lines.append(line)
            rating_types.append(record.rating_type)
        trailering_lines = unique_preserve_order(trailering_lines)
        if trailering_lines:
            rating_text = ' ; '.join(unique_preserve_order(rating_types))
            for idx, chunk in enumerate(chunk_list(trailering_lines, max_words=125, max_items=8), start=1):
                title = self.cleaner.article_heading(entity, self.cleaner.t('Trailering values'))
                if idx > 1:
                    title = self.cleaner.article_heading(entity, f"{self.cleaner.t('Trailering values')} | {self.cleaner.t('part')} {idx}")
                parts.append(
                    '<section class="trailering-values">'
                    + self.cleaner.cleaned_render_article(
                        title,
                        self.cleaner.filtered_identity_fields(
                            data,
                            trim,
                            category=self.cleaner.t('Powertrain and trailering') if self.cleaner.language == 'fr' else 'Powertrain and trailering',
                            extra_fields=[(self.cleaner.t('Trailering ratings'), rating_text)] if rating_text else [],
                        ),
                        [(self.cleaner.t('Specifications'), chunk)],
                    )
                    + '</section>'
                )
        return self.cleaner.clean_title_document(
            self.cleaner.article_heading(entity, f"{self.cleaner.t('Powertrain and trailering')} | {group.model_code}"),
            ''.join(parts),
        )

    def render_gcwr_reference_page(self, data, records: Sequence[GCWRRecord]) -> str:
        entity = data.vehicle_name
        parts = [
            '<section class="gcwr-identity">'
            + self.cleaner.cleaned_render_article(
                self.cleaner.article_heading(entity, self.cleaner.t('GCWR reference')),
                self.cleaner.filtered_identity_fields(data, category='Trailering and GCWR'),
            )
            + '</section>'
        ]
        lines: List[str] = []
        for record in records:
            line = f"{('Moteur' if self.cleaner.language == 'fr' else 'Engine')} {record.engine} | GCWR {record.gcwr} | {'Rapport de pont' if self.cleaner.language == 'fr' else 'Axle ratio'} {record.axle_ratio}"
            if record.footnotes:
                line += f" | {self.cleaner.t('Notes')}: {' ; '.join(unique_preserve_order(record.footnotes))}"
            lines.append(line)
        for idx, chunk in enumerate(chunk_list(unique_preserve_order(lines), max_words=125, max_items=10), start=1):
            title = self.cleaner.article_heading(entity, self.cleaner.t('GCWR values'))
            if idx > 1:
                title = self.cleaner.article_heading(entity, f"{self.cleaner.t('GCWR values')} | {self.cleaner.t('part')} {idx}")
            parts.append(
                '<section class="gcwr-values">'
                + self.cleaner.cleaned_render_article(
                    title,
                    self.cleaner.filtered_identity_fields(data, category='Trailering and GCWR'),
                    [(self.cleaner.t('Specifications'), chunk)],
                )
                + '</section>'
            )
        return self.cleaner.clean_title_document(self.cleaner.article_heading(entity, self.cleaner.t('GCWR reference')), ''.join(parts))
