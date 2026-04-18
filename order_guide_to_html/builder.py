from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from .utils import short_slug, slugify, unique_output_path
from .models import BoundRecord, OutputFileRecord
from .classification import (
    COMPARISON_OBJECTTYPE, CONFIG_OBJECTTYPE, CONFIG_TYPE, MODEL_OBJECTTYPE,
    TRIM_OBJECTTYPE, comparison_varies_by_trim, normalize_domain_value,
)
from .configuration import (
    all_trim_matches, all_trim_matches_for_spec_group, best_trim_match, group_powertrain_trailering_for_cpr, group_spec_columns_for_cpr,
    powertrain_group_trim_match, spec_group_model_code, strip_drive_tokens,
)
from .manifest import (
    add_bound_record, build_manifest_from_bindings,
    comparison_domain_manifest_metadata, config_parent_manifest_metadata,
    gcwr_reference_manifest_metadata, model_overview_manifest_metadata,
    powertrain_trailering_manifest_metadata, spec_group_manifest_metadata,
    trim_entity_key, trim_overview_manifest_metadata, vehicle_key,
    vehicle_manifest_filename, with_flat_doc_metadata,
)
from .aggregation import FeatureAggregationService
from .cleaning import GuideTextCleaner, LOW_SIGNAL_COMPARISON_DOMAINS, normalize_text
from .rendering import HtmlRenderer


class CorpusBuilder:
    """Explicit write pipeline for the cleaned flat CPR corpus."""

    def __init__(self, cleaner: GuideTextCleaner, aggregation: FeatureAggregationService, renderer: HtmlRenderer):
        self.cleaner = cleaner
        self.aggregation = aggregation
        self.renderer = renderer

    def _config_metadata(self, data, metadata: Dict[str, object], matched_trim, domain: str) -> Dict[str, object]:
        extra: Dict[str, object] = {
            'entity_name': self.cleaner.clean_trim_heading(data, matched_trim) if matched_trim is not None else data.vehicle_name,
            'entity_key': trim_entity_key(data, matched_trim) if matched_trim is not None else vehicle_key(data),
            'vehicle_key': vehicle_key(data),
            'year': normalize_text(data.year),
            'make': normalize_text(data.make),
            'model': normalize_text(data.model),
            'guide_domains': [normalize_domain_value(domain)],
            'retrieval_granularity': 'coarse',
        }
        if matched_trim is not None:
            extra['trim_key'] = trim_entity_key(data, matched_trim)
            extra['trim'] = strip_drive_tokens(normalize_text(matched_trim.name))
        return with_flat_doc_metadata(metadata, **extra)

    def build_model_and_comparisons(self, data, output_dir: Path, used_names: set[str], bindings: List[BoundRecord]) -> Path:
        model_filename = f'model_{data.year}_{slugify(data.make)}_{slugify(data.model)}.html'
        model_path = unique_output_path(output_dir, model_filename, used_names)
        model_path.write_text(self.renderer.render_model_overview_page(data), encoding='utf-8')
        model_record = OutputFileRecord(
            objecttype=MODEL_OBJECTTYPE,
            type='model',
            path=model_path,
            metadata=model_overview_manifest_metadata(data),
        )
        add_bound_record(bindings, model_record, collection=model_path, parent=None, parent_vehicle=None, parent_trims=[])

        for category, features in self.aggregation.model_feature_groups_by_category(data).items():
            domain = normalize_text(category)
            if domain == 'Other guide content':
                continue
            if domain in LOW_SIGNAL_COMPARISON_DOMAINS:
                continue
            if not features:
                continue
            selected_features = [feature for feature in features if comparison_varies_by_trim(feature)] or list(features)
            if not selected_features:
                continue
            domain_filename = f'compare_domain_{data.year}_{slugify(data.make)}_{slugify(data.model)}_{short_slug(category)}.html'
            domain_path = unique_output_path(output_dir, domain_filename, used_names)
            domain_path.write_text(self.renderer.render_comparison_domain_page(data, category, selected_features), encoding='utf-8')
            domain_record = OutputFileRecord(
                objecttype=COMPARISON_OBJECTTYPE,
                type='comparison',
                path=domain_path,
                metadata=comparison_domain_manifest_metadata(data, category, selected_features),
            )
            add_bound_record(bindings, domain_record, collection=model_path, parent=model_path, parent_vehicle=model_path, parent_trims=[])
        return model_path

    def build_trims(self, data, output_dir: Path, used_names: set[str], bindings: List[BoundRecord], model_path: Path) -> Dict[str, Path]:
        trim_paths: Dict[str, Path] = {}
        name_slug_to_path: Dict[str, Path] = {}
        for trim in data.trim_defs:
            filename = f'trim_{data.year}_{slugify(data.make)}_{slugify(data.model)}_{slugify(trim.name)}.html'
            if filename in name_slug_to_path:
                trim_paths[trim.key] = name_slug_to_path[filename]
                continue
            path = unique_output_path(output_dir, filename, used_names)
            path.write_text(self.renderer.render_trim_overview_page(data, trim), encoding='utf-8')
            record = OutputFileRecord(
                objecttype=TRIM_OBJECTTYPE,
                type='trim',
                path=path,
                metadata=trim_overview_manifest_metadata(data, trim),
            )
            add_bound_record(bindings, record, collection=model_path, parent=model_path, parent_vehicle=model_path, parent_trims=[])
            trim_paths[trim.key] = path
            name_slug_to_path[filename] = path
        return trim_paths

    def build_configurations(
        self,
        data,
        output_dir: Path,
        used_names: set[str],
        bindings: List[BoundRecord],
        model_path: Path,
        trim_paths: Dict[str, Path],
    ) -> None:
        trim_path_by_name = {normalize_text(trim.name): trim_paths[trim.key] for trim in data.trim_defs if trim.key in trim_paths}

        def parent_paths_for_trims(matched_trims) -> Tuple[Path, List[Path]]:
            paths = [p for trim in matched_trims if (p := trim_path_by_name.get(normalize_text(trim.name))) is not None]
            if not paths:
                return model_path, []
            parent = paths[0] if len(paths) == 1 else model_path
            return parent, paths

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
            page_path.write_text(self.renderer.render_spec_group_page(data, group, trim=matched_trim), encoding='utf-8')
            domain = 'Dimensions and specifications'
            metadata = self._config_metadata(
                data,
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
            model_code_slug = slugify(group.model_code)
            base_name = f'config_{data.year}_{slugify(data.make)}_{slugify(data.model)}_powertrain_trailering_{model_code_slug}'
            if trim_slug and trim_slug not in model_code_slug:
                base_name += f'_{trim_slug}'
            page_path = unique_output_path(output_dir, base_name + '.html', used_names)
            page_path.write_text(self.renderer.render_powertrain_trailering_group_page(data, group, trim=matched_trim), encoding='utf-8')
            domain = 'Powertrain and trailering'
            metadata = self._config_metadata(
                data,
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
            page_path.write_text(self.renderer.render_gcwr_reference_page(data, data.gcwr_records), encoding='utf-8')
            domain = 'Trailering and GCWR'
            metadata = self._config_metadata(
                data,
                config_parent_manifest_metadata(
                    gcwr_reference_manifest_metadata(data, data.gcwr_records),
                    domain=domain,
                    doc_type='configuration-spec',
                ),
                None,
                domain,
            )
            record = OutputFileRecord(objecttype=CONFIG_OBJECTTYPE, type=CONFIG_TYPE, path=page_path, metadata=metadata)
            add_bound_record(bindings, record, collection=model_path, parent=model_path, parent_vehicle=model_path, parent_trims=[])

    def write_outputs(self, data, output_dir: Path) -> Dict[str, object]:
        self.cleaner.set_language(getattr(data, 'language', 'en'))
        self.cleaner.load_glossary(getattr(data, 'glossary', {}))
        output_dir.mkdir(parents=True, exist_ok=True)
        used_names: set[str] = set()
        bindings: List[BoundRecord] = []

        model_path = self.build_model_and_comparisons(data, output_dir, used_names, bindings)
        trim_paths = self.build_trims(data, output_dir, used_names, bindings, model_path)
        self.build_configurations(data, output_dir, used_names, bindings, model_path, trim_paths)

        manifest_path = output_dir / vehicle_manifest_filename(data)
        manifest = build_manifest_from_bindings(data, bindings, manifest_path)
        cleaned_manifest = self.cleaner.clean_manifest(manifest)
        manifest_path.write_text(json.dumps(cleaned_manifest, indent=2, ensure_ascii=False), encoding='utf-8')
        return cleaned_manifest
