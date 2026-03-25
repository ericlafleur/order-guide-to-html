"""Vehicle attribute inference module."""

from .propulsion import infer_propulsion
from .vehicle_type import infer_vehicle_type
from .drive_type import infer_drive_types, infer_trim_drive_type, parse_drive_type_from_text

__all__ = [
    "infer_propulsion",
    "infer_vehicle_type",
    "infer_drive_types",
    "infer_trim_drive_type",
    "parse_drive_type_from_text",
]
