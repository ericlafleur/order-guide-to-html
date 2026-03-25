"""Propulsion type inference."""

import re
from typing import TYPE_CHECKING, List

from ..utils.constants import PROPULSION_KEYWORDS

if TYPE_CHECKING:
    from ..models.base_models import EngineAxleEntry


def infer_propulsion(vehicle_name: str, engine_axle_entries: 'List[EngineAxleEntry]') -> str:
    """Infer propulsion type (EV / HYBRID / PHEV / ICE).
    
    Engine data is checked first because it is the most reliable signal.
    Short keywords like 'ev' are matched as whole words in the vehicle name to
    avoid false positives (e.g. "Silv**er**ado" containing 'ev').
    
    Args:
        vehicle_name: Full vehicle name
        engine_axle_entries: List of engine/axle entries
        
    Returns:
        Propulsion type string
    """
    # 1. Engine descriptions are the most reliable source
    for entry in engine_axle_entries:
        eng = entry.engine.lower()
        for kw, prop in PROPULSION_KEYWORDS.items():
            if kw in eng:
                return prop

    # 2. Vehicle name — use word-boundary matching for short tokens
    name_lower = vehicle_name.lower()
    for kw, prop in PROPULSION_KEYWORDS.items():
        if len(kw) <= 3:
            if re.search(r'\b' + re.escape(kw) + r'\b', name_lower):
                return prop
        else:
            if kw in name_lower:
                return prop

    return 'ICE'
