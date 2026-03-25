"""Aggregate models for feature grouping."""

from collections import OrderedDict
from dataclasses import dataclass, field
from typing import List, Tuple


@dataclass
class ModelFeatureAggregate:
    """Aggregated feature data across all trims at model level."""
    
    title: str
    description: str
    orderable_code: str
    reference_code: str
    source_contexts: List[str] = field(default_factory=list)
    availability_contexts: OrderedDict[
        Tuple[Tuple[Tuple[str, str, Tuple[str, ...]], ...]], 
        List[str]
    ] = field(default_factory=OrderedDict)
    notes: List[str] = field(default_factory=list)
    referenced_codes: List[Tuple[str, str]] = field(default_factory=list)


@dataclass
class TrimFeatureAggregate:
    """Aggregated feature data for a specific trim."""
    
    title: str
    description: str
    orderable_code: str
    reference_code: str
    availability_contexts: OrderedDict[Tuple[str, str], List[str]] = field(default_factory=OrderedDict)
    source_contexts: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)
    referenced_codes: List[Tuple[str, str]] = field(default_factory=list)
