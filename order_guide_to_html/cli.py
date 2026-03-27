from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Optional, Sequence

from .parsing import parse_workbook
from .manifest import vehicle_manifest_filename
from .aggregation import FeatureAggregationService
from .builder import CorpusBuilder
from .cleaning import GuideTextCleaner
from .rendering import HtmlRenderer


def build_pipeline() -> CorpusBuilder:
    cleaner = GuideTextCleaner()
    aggregation = FeatureAggregationService(cleaner)
    renderer = HtmlRenderer(cleaner, aggregation)
    return CorpusBuilder(cleaner, aggregation, renderer)


WORKBOOKS_DIR = Path("workbooks")
OUTPUT_DIR = Path("workbooks_html")


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description='Convert a GM Vehicle Order Guide workbook into chunk-budget-aware model and trim HTML files for RAG.'
    )
    parser.add_argument(
        'files',
        nargs='*',
        type=Path,
        help='Specific xlsx files to convert. Defaults to all files in workbooks/.',
    )
    parser.add_argument('-o', '--output-dir', help='Directory for generated HTML files')
    args = parser.parse_args(argv)

    output_dir = Path(args.output_dir) if args.output_dir else OUTPUT_DIR

    xlsx_files = args.files if args.files else list(WORKBOOKS_DIR.glob("*.xlsx"))

    if not xlsx_files:
        print(f"No xlsx files found in {WORKBOOKS_DIR}/")
        return 0

    for workbook_path in xlsx_files:
        if not workbook_path.exists():
            print(f'Workbook not found: {workbook_path}', file=sys.stderr)
            continue
        data = parse_workbook(workbook_path)
        builder = build_pipeline()
        manifest = builder.write_outputs(data, output_dir)
        summary = {
            'workbook': str(workbook_path),
            'vehicle_name': data.vehicle_name,
            'output_dir': str(output_dir),
            'file_count': len(manifest.get('files', [])),
        }
    return 0
