#!/usr/bin/env python3
"""Order Guide to HTML converter - Main entry point.

Converts GM Vehicle Order Guide Excel workbooks into structured HTML files
optimized for RAG (Retrieval-Augmented Generation) systems.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Optional, Sequence

from src.manifest import write_outputs
from src.parsers import parse_workbook

# Default directories
WORKBOOKS_DIR = Path("workbooks")
OUTPUT_DIR = Path("workbooks_html")


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Main entry point for order guide conversion.
    
    Args:
        argv: Command-line arguments
        
    Returns:
        Exit code (0 for success)
    """
    parser = argparse.ArgumentParser(
        description='Convert GM Vehicle Order Guide workbooks into HTML files for RAG.'
    )
    parser.add_argument(
        'files',
        nargs='*',
        type=Path,
        help='Specific xlsx files to convert. Defaults to all files in workbooks/.',
    )
    parser.add_argument(
        '-o', 
        '--output-dir', 
        type=Path,
        help='Directory for generated HTML files (default: workbooks_html/)'
    )
    args = parser.parse_args(argv)

    output_dir = args.output_dir if args.output_dir else OUTPUT_DIR
    xlsx_files = args.files if args.files else list(WORKBOOKS_DIR.glob("*.xlsx"))

    if not xlsx_files:
        print(f"No xlsx files found in {WORKBOOKS_DIR}/")
        return 0

    for workbook_path in xlsx_files:
        if not workbook_path.exists():
            print(f'Workbook not found: {workbook_path}', file=sys.stderr)
            continue
        
        try:
            # Parse workbook
            data = parse_workbook(workbook_path)
            
            # Write HTML files and manifest
            manifest = write_outputs(data, output_dir)
            
            # Print manifest
            print(json.dumps(manifest, indent=2))
        
        except Exception as e:
            print(f'Error processing {workbook_path}: {e}', file=sys.stderr)
            import traceback
            traceback.print_exc()
            continue
    
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
