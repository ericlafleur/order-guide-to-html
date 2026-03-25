# Order Guide to HTML - Source Code Structure

This directory contains the refactored, modular codebase for converting GM Vehicle Order Guide Excel workbooks into structured HTML files.

## Directory Structure

```
src/
├── __init__.py              # Package initialization
├── utils/                   # Utility functions and constants
│   ├── __init__.py
│   ├── constants.py        # Constants (status labels, keywords, patterns)
│   └── text_utils.py       # Text processing utilities
├── models/                  # Data models
│   ├── __init__.py
│   ├── base_models.py      # Core dataclasses (WorkbookData, MatrixSheet, etc.)
│   └── aggregate_models.py # Aggregation dataclasses
├── parsers/                 # Excel sheet parsers
│   ├── __init__.py
│   ├── common.py           # Common parsing utilities
│   ├── matrix_parser.py    # Equipment matrix parser
│   ├── color_parser.py     # Color and trim parser
│   ├── spec_parser.py      # Specifications parser
│   ├── engine_axle_parser.py  # Engine/axle parser
│   ├── trailering_parser.py   # Trailering specs parser
│   ├── glossary_parser.py     # Glossary parser
│   └── workbook_parser.py     # Main workbook orchestrator
├── inference/              # Vehicle attribute inference
│   ├── __init__.py
│   ├── propulsion.py       # Propulsion type inference (EV/ICE/HYBRID)
│   ├── vehicle_type.py     # Vehicle type inference (SUV/truck)
│   └── drive_type.py       # Drive type inference (AWD/FWD/RWD)
├── aggregators/            # Feature aggregation logic
│   ├── __init__.py
│   ├── code_reference.py   # Code reference extraction
│   └── feature_aggregator.py  # Feature grouping and categorization
├── renderers/              # HTML rendering
│   ├── __init__.py
│   ├── common.py           # Common rendering utilities
│   ├── sections.py         # Section renderers (features, colors, specs)
│   ├── model_renderer.py   # Model page renderer
│   └── trim_renderer.py    # Trim page renderer
└── manifest/               # Manifest generation
    ├── __init__.py
    ├── manifest_generator.py  # Main manifest creation
    └── metadata_extractor.py  # Metadata extraction

```

## Design Principles

### Separation of Concerns
Each module has a single, well-defined responsibility:
- **Parsers** extract data from Excel sheets
- **Inference** determines vehicle attributes
- **Aggregators** group and categorize features
- **Renderers** generate HTML output
- **Manifest** creates metadata files

### DRY (Don't Repeat Yourself)
- Common utilities are centralized in `utils/`
- Shared constants live in `constants.py`
- Text processing functions are reusable across modules

### Modularity
- Each parser handles one sheet type
- Renderers are split by page type and sections
- Easy to test individual components

### Type Safety
- Dataclasses for structured data
- Type hints throughout
- Clear interfaces between modules

## Usage

The main entry point `order_guide_to_html.py` imports from this structure:

```python
from src.parsers import parse_workbook
from src.manifest import write_outputs

# Parse workbook
data = parse_workbook(path)

# Generate HTML and manifest
manifest = write_outputs(data, output_dir)
```

## Adding New Features

### New Sheet Type
1. Create parser in `parsers/new_sheet_parser.py`
2. Add data model in `models/base_models.py`
3. Import and use in `workbook_parser.py`

### New Rendering Section
1. Add section renderer in `renderers/sections.py`
2. Call from `model_renderer.py` or `trim_renderer.py`

### New Inference Logic
1. Create module in `inference/`
2. Call from `workbook_parser.py` or other appropriate location

## Testing

Each module can be tested independently:

```python
# Test parser
from src.parsers.matrix_parser import parse_matrix_sheet
result = parse_matrix_sheet(worksheet)

# Test renderer
from src.renderers.model_renderer import render_model_page
html = render_model_page(workbook_data)
```
