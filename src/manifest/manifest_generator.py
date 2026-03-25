"""Manifest generation and file output."""

import json
import os
from pathlib import Path
from typing import TYPE_CHECKING, Dict, List, Optional, Sequence

from ..aggregators.feature_aggregator import aggregate_model_features, aggregate_trim_features
from ..inference import infer_trim_drive_type
from ..parsers.common import parse_status_value
from ..renderers import render_model_page, render_trim_page
from ..utils.constants import MANIFEST_BODY_STYLE_TOKENS, MANIFEST_DRIVE_TOKENS, MANIFEST_STANDARDISH_CODES
from ..utils.text_utils import normalize_text, slugify, unique_preserve_order
from .metadata_extractor import (
    extract_manifest_metadata_for_model,
    extract_manifest_metadata_for_trim,
)

if TYPE_CHECKING:
    from ..models.base_models import TrimDef, WorkbookData


def vehicle_manifest_filename(data: 'WorkbookData') -> str:
    """Generate manifest filename for vehicle.
    
    Args:
        data: Workbook data
        
    Returns:
        Manifest filename
    """
    return f'manifest_{slugify(data.year)}_{slugify(data.make)}_{slugify(data.model)}.json'


def manifest_relpath(path: Path, base_dir: Path) -> str:
    """Get relative path for manifest.
    
    Args:
        path: Target path
        base_dir: Base directory
        
    Returns:
        Relative path string
    """
    try:
        return os.path.relpath(str(path), str(base_dir))
    except (ValueError, OSError):
        return str(path)


def write_outputs(data: 'WorkbookData', output_dir: Path) -> Dict[str, object]:
    """Write HTML files and manifest for workbook data.
    
    Args:
        data: Workbook data
        output_dir: Output directory path
        
    Returns:
        Manifest dictionary
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    # Write model page
    model_filename = f'model_{data.year}_{slugify(data.make)}_{slugify(data.model)}.html'
    model_path = output_dir / model_filename
    model_path.write_text(render_model_page(data), encoding='utf-8')

    # Write trim pages
    trim_paths: Dict[str, Path] = {}
    for trim in data.trim_defs:
        trim_filename = f'trim_{data.year}_{slugify(data.make)}_{slugify(data.model)}_{slugify(trim.name)}.html'
        trim_path = output_dir / trim_filename
        trim_path.write_text(render_trim_page(data, trim), encoding='utf-8')
        trim_paths[trim.key] = trim_path

    # Create manifest
    manifest_path = output_dir / vehicle_manifest_filename(data)
    manifest_base = manifest_path.parent

    manifest: Dict[str, object] = {
        'workbook': manifest_relpath(data.path, manifest_base),
        'vehicle_name': data.vehicle_name,
        'files': [],
    }

    # Model entry
    model_entry: Dict[str, object] = {
        'objecttype': 'Product',
        'type': 'model',
    }
    model_entry.update(extract_manifest_metadata_for_model(data))
    model_entry['path'] = manifest_relpath(model_path, manifest_base)
    manifest['files'].append(model_entry)

    # Trim entries
    model_relpath = manifest_relpath(model_path, manifest_base)
    for trim in data.trim_defs:
        trim_entry: Dict[str, object] = {
            'objecttype': 'Variant',
            'type': 'trim',
        }
        trim_entry.update(extract_manifest_metadata_for_trim(data, trim))
        trim_entry['parent_model'] = model_relpath
        trim_entry['path'] = manifest_relpath(trim_paths[trim.key], manifest_base)
        manifest['files'].append(trim_entry)

    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding='utf-8')
    return manifest
