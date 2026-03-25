"""Trim page renderer."""

import html as html_module
from typing import TYPE_CHECKING

from .common import full_trim_heading
from .sections import (
    render_page_identity_section,
    render_trim_color_sections,
    render_trim_feature_sections,
    render_trim_spec_sections,
)

if TYPE_CHECKING:
    from ..models.base_models import TrimDef, WorkbookData


def render_trim_page(data: 'WorkbookData', trim: 'TrimDef') -> str:
    """Render complete trim page HTML.
    
    Args:
        data: Workbook data
        trim: Trim definition
        
    Returns:
        Complete HTML page string
    """
    entity = full_trim_heading(data, trim)
    parts = [
        '<html><head><meta charset="utf-8"></head><body>',
        f'<h1>{html_module.escape(entity)} | Vehicle Order Guide</h1>',
        render_page_identity_section(data, trim=trim),
        render_trim_feature_sections(data, trim),
        render_trim_color_sections(data, trim),
        render_trim_spec_sections(data, trim),
        '</body></html>',
    ]
    return ''.join(part for part in parts if part)
