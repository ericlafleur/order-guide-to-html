"""Data models for order guide parsing."""

from .base_models import (
    TrimDef,
    MatrixRow,
    MatrixSheet,
    ColorInteriorRow,
    ColorExteriorRow,
    ColorSheet,
    SpecCell,
    SpecColumn,
    EngineAxleItem,
    EngineAxleEntry,
    TraileringRecord,
    GCWRRecord,
    WorkbookData,
)
from .aggregate_models import (
    ModelFeatureAggregate,
    TrimFeatureAggregate,
)

__all__ = [
    # Base models
    "TrimDef",
    "MatrixRow",
    "MatrixSheet",
    "ColorInteriorRow",
    "ColorExteriorRow",
    "ColorSheet",
    "SpecCell",
    "SpecColumn",
    "EngineAxleItem",
    "EngineAxleEntry",
    "TraileringRecord",
    "GCWRRecord",
    "WorkbookData",
    # Aggregate models
    "ModelFeatureAggregate",
    "TrimFeatureAggregate",
]
