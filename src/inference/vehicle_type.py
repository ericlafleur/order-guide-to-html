"""Vehicle type inference."""

from ..utils.constants import VEHICLE_TYPE_SUV_KEYWORDS, VEHICLE_TYPE_TRUCK_KEYWORDS


def infer_vehicle_type(vehicle_name: str) -> str:
    """Infer vehicle type (suv, truck, van, etc.) from vehicle name.
    
    Args:
        vehicle_name: Full vehicle name
        
    Returns:
        Vehicle type string
    """
    name_lower = vehicle_name.lower()
    
    for kw in VEHICLE_TYPE_SUV_KEYWORDS:
        if kw in name_lower:
            return 'suv'
    
    for kw in VEHICLE_TYPE_TRUCK_KEYWORDS:
        if kw in name_lower:
            return 'truck'
    
    return 'suv'
