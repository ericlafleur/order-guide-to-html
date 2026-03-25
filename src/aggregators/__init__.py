"""Feature aggregation module."""

from .feature_aggregator import (
    aggregate_model_features,
    aggregate_trim_features,
    category_for_model_feature,
    category_for_trim_feature,
    sort_trim_feature,
)
from .code_reference import referenced_codes_for_text

__all__ = [
    "aggregate_model_features",
    "aggregate_trim_features",
    "category_for_model_feature",
    "category_for_trim_feature",
    "sort_trim_feature",
    "referenced_codes_for_text",
]
