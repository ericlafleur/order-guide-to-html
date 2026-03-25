"""Model page renderer."""

import html as html_module
from typing import TYPE_CHECKING

from .sections import (
    render_engine_axles_section,
    render_model_color_sections,
    render_model_feature_sections,
    render_page_identity_section,
    render_spec_records,
    render_trailering_section,
)

if TYPE_CHECKING:
    from ..models.base_models import WorkbookData


def render_model_page(data: 'WorkbookData') -> str:
    """Render complete model page HTML.
    
    Args:
        data: Workbook data
        
    Returns:
        Complete HTML page string
    """
    entity = data.vehicle_name
    parts = [
        '<html><head><meta charset="utf-8"></head><body>',
        f'<h1>{html_module.escape(entity)} | Vehicle Order Guide</h1>',
        render_page_identity_section(data),
        render_model_feature_sections(data),
        render_model_color_sections(data),
        render_spec_records(data, data.spec_columns),
        render_engine_axles_section(data),
        render_trailering_section(data),
        '</body></html>',
    ]
    return ''.join(part for part in parts if part)
