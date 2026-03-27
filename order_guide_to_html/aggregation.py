from __future__ import annotations

from collections import OrderedDict
from typing import Iterable, List, Sequence

from .models import ModelFeatureAggregate, TrimFeatureAggregate
from .parsing import parse_status_value, referenced_codes_for_text
from .classification import (
    category_for_model_feature, category_for_trim_feature, collect_row_note_texts,
    sort_category_key, sort_trim_feature, source_context, summarize_model_status_groups,
)
from .cleaning import GuideTextCleaner, normalize_text, unique_preserve_order


class FeatureAggregationService:
    """Aggregate workbook matrix rows into model- and trim-level features without monkey-patching."""

    def __init__(self, cleaner: GuideTextCleaner):
        self.cleaner = cleaner

    def aggregate_model_features(self, data) -> List[ModelFeatureAggregate]:
        groups: 'OrderedDict[str, ModelFeatureAggregate]' = OrderedDict()
        for sheet in data.matrix_sheets:
            for row in sheet.rows:
                title = self.cleaner.clean_feature_title(self.cleaner.matrix_row_label(row), row.option_code or '', row.ref_code or '')
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

    def aggregate_trim_features(self, data, trim) -> List[TrimFeatureAggregate]:
        groups: 'OrderedDict[str, TrimFeatureAggregate]' = OrderedDict()
        for sheet in data.matrix_sheets:
            for row in sheet.rows:
                raw = normalize_text(row.status_by_trim.get(trim.key))
                if not raw:
                    continue
                _code, label, _notes = parse_status_value(raw, row.inline_footnotes, sheet.footnotes)
                title = self.cleaner.clean_feature_title(self.cleaner.matrix_row_label(row), row.option_code or '', row.ref_code or '')
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

    def model_feature_groups_by_category(self, data) -> 'OrderedDict[str, List[ModelFeatureAggregate]]':
        groups: 'OrderedDict[str, List[ModelFeatureAggregate]]' = OrderedDict()
        for feature in self.aggregate_model_features(data):
            groups.setdefault(category_for_model_feature(feature), []).append(feature)
        ordered = OrderedDict()
        for category in sorted(groups.keys(), key=sort_category_key):
            ordered[category] = groups[category]
        return ordered

    def trim_feature_groups_by_category(self, data, trim) -> 'OrderedDict[str, List[TrimFeatureAggregate]]':
        groups: 'OrderedDict[str, List[TrimFeatureAggregate]]' = OrderedDict()
        for feature in self.aggregate_trim_features(data, trim):
            groups.setdefault(category_for_trim_feature(feature), []).append(feature)
        ordered = OrderedDict()
        for category in sorted(groups.keys(), key=sort_category_key):
            ordered[category] = groups[category]
        return ordered
